#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod parse;
mod db;

use std::time::{Duration, Instant};
use eframe::{NativeOptions};
use eframe::egui::*;
use egui_extras::{Column, TableBuilder};
use rust_xlsxwriter::{Format, Workbook};
use crate::db::{DBManager, ViewQuery};
use crate::parse::{parse_import_files, parse_sifrant_file};

struct App {
    db_manager: DBManager,
    retry_import: bool,
    successfully_parsed: Option<bool>,
    successfully_stored_data: Option<bool>,
    successfully_stored_sifrant: Option<bool>,

    row_data: Option<Vec<ViewQuery>>,
    successfully_loaded_query: Option<bool>,

    successfully_exported: Option<bool>,
    export_counter: Option<Instant>,

    opomba_material: String,
    opomba_opomba: String,
    successfully_stored_opomba: Option<bool>,

    dobavni_rok_material: String,
    dobavni_rok_dobavni_rok: String,
    successfully_stored_dobavni_rok: Option<bool>,


    /* --Columns-- */



    /* --Filters-- */
    filter_rdeca: bool,
    filter_oranzna: bool,
    filter_rumena: bool,
    filter_modra_zelena: bool,

    filter_sifra_materiala: String,
    filter_naziv_materiala: String,
    filter_nabavnik: String,
    filter_zaloga_vecja: bool,
    filter_zaloga_gt: String,

    filter_nabavnik_definiran: bool,
    filter_nabavnik_ni_definiran: bool,

    filter_poraba_vecja: bool,
    filter_poraba_gt: String,

    filter_odprta_narocila: bool,
    filter_odprta_narocila_gt: String,

    filter_dobavni_rok: bool,
    filter_dobavni_rok_gt: String,

    filter_opomba: bool,
}

impl App {
    pub fn new<'a>(cc: &'a eframe::CreationContext<'a>) -> Self {
        cc.egui_ctx.send_viewport_cmd(ViewportCommand::Maximized(true));
        cc.egui_ctx.set_visuals(Visuals::light());

        cc.egui_ctx.style_mut(|s| {
           s.override_text_style = Some(TextStyle::Heading);
            s.visuals.override_text_color = Some(Color32::from_rgb(5, 5, 5));
        });

        let mut row_data = None;
        let db_manager = DBManager { db_name: "magneti_db.sqlite3".to_string() };
        let _ = db_manager.try_create_tables();

        let result = db_manager.get_data();

        let successfully_loaded_query = match result {
            Err(err) => {
                println!("initial_load error: {:?}", err.to_string());
                Some(false)
            },
            Ok(rows) => {
                row_data = Some(rows);
                println!("row_data loaded: {}", row_data.as_ref().unwrap().len());
                Some(true)
            }
        };

        Self {
            db_manager,
            retry_import: false,
            successfully_parsed: None,
            successfully_stored_data: None,
            successfully_stored_sifrant: None,
            row_data,
            successfully_loaded_query,

            successfully_exported: None,
            export_counter: None,

            opomba_material: String::new(),
            opomba_opomba: String::new(),
            successfully_stored_opomba: None,

            dobavni_rok_material: String::new(),
            dobavni_rok_dobavni_rok: String::new(),
            successfully_stored_dobavni_rok: None,

            filter_rdeca: false,
            filter_oranzna: false,
            filter_rumena: false,
            filter_modra_zelena: false,


            filter_sifra_materiala: String::new(),
            filter_naziv_materiala: String::new(),
            filter_nabavnik: String::new(),
            filter_zaloga_vecja: true,
            filter_zaloga_gt: String::from("0"),

            filter_nabavnik_definiran: true,
            filter_nabavnik_ni_definiran: false,

            filter_poraba_vecja: true,
            filter_poraba_gt: String::from("0"),

            filter_odprta_narocila: false,
            filter_odprta_narocila_gt: String::from("0"),

            filter_dobavni_rok: false,
            filter_dobavni_rok_gt: String::from("0"),

            filter_opomba: false,
        }
    }

}


impl App {
    fn apply_filters(&self, rows: &Vec<ViewQuery>) -> Vec<ViewQuery> {

        rows.iter()
            .filter(|&row| {
                // --Evil GPT hacked-- //
                let any_color_filter_active = self.filter_rumena || self.filter_oranzna || self.filter_rdeca || self.filter_modra_zelena;
                let mut color_matches = false;
                if any_color_filter_active {
                    let months_left = row.dobavni_rok.map_or(0., |dr| row.trenutna_zaloga_zadostuje_za_mesecev.unwrap_or(0.) - dr);
                    let no_open_orders = row.odprta_narocila.is_some_and(|v| v == 0.);

                    color_matches |= self.filter_rumena && row.dobavni_rok.is_some() && months_left >= 1.5 && months_left < 3. && no_open_orders;
                    color_matches |= self.filter_oranzna && row.dobavni_rok.is_some() && months_left >= 0.5 && months_left < 1.5 && no_open_orders;
                    color_matches |= self.filter_rdeca && row.dobavni_rok.is_some() && months_left < 0.5 && no_open_orders;
                    color_matches |= self.filter_modra_zelena && row.dobavni_rok.is_some() && !no_open_orders;  // has open orders
                }
                // ---- //

                let condition = if self.filter_zaloga_vecja && self.filter_poraba_vecja {
                    // both are checked then OR
                    row.zaloga.is_some_and(|zal| zal > parse_string_to_optional_f64(self.filter_zaloga_gt.as_str()).unwrap_or(0.)) ||
                        row.poraba_3m.is_some_and(|por| por > parse_string_to_optional_f64(self.filter_poraba_gt.as_str()).unwrap_or(0.))
                } else if self.filter_zaloga_vecja && !self.filter_poraba_vecja {
                    row.zaloga.is_some_and(|zal| zal > parse_string_to_optional_f64(self.filter_zaloga_gt.as_str()).unwrap_or(0.))
                } else if !self.filter_zaloga_vecja && self.filter_poraba_vecja {
                    row.poraba_3m.is_some_and(|por| por > parse_string_to_optional_f64(self.filter_poraba_gt.as_str()).unwrap_or(0.))
                } else {
                    true
                };


                format!("{}", row.material).contains(self.filter_sifra_materiala.as_str()) &&
                    row.naziv_materiala.as_ref().is_some_and(|a| format!("{}", a.to_lowercase()).contains(self.filter_naziv_materiala.to_lowercase().as_str())) &&
                    row.nabavna_skupina.as_ref().is_some_and(|a| format!("{}", a.to_lowercase()).contains(self.filter_nabavnik.to_lowercase().as_str())) &&

                    condition &&

                    (!self.filter_odprta_narocila || row.odprta_narocila.is_some_and(|odp| odp > parse_string_to_optional_f64(self.filter_odprta_narocila_gt.as_str()).unwrap_or(0.))) &&
                    (!self.filter_dobavni_rok || row.dobavni_rok.is_some_and(|dob| dob > parse_string_to_optional_f64(self.filter_dobavni_rok_gt.as_str()).unwrap_or(0.))) &&
                    (color_matches || !any_color_filter_active) &&
                    (!self.filter_nabavnik_ni_definiran || row.nabavna_skupina.as_ref().is_some_and(|str| str.eq(""))) &&
                    (!self.filter_nabavnik_definiran || !row.nabavna_skupina.as_ref().is_some_and(|str| str.eq(""))) &&
                    (!self.filter_opomba || row.opomba.as_ref().is_some_and(|s| !s.eq("")))
        })
            .cloned()
            .collect()


    }


    pub fn render_table(&self, ui: &mut Ui, data: &Vec<ViewQuery>) {
        let number_width = 100.;
        let string_width = 550.;

        ScrollArea::both().show(ui, |ui| {
            ui.style_mut().visuals.faint_bg_color = Color32::from_rgb(200, 200, 200);
            TableBuilder::new(ui)
                .striped(true)
                .cell_layout(Layout::left_to_right(Align::Center))
                .columns(Column::exact(number_width), 1) // Material
                .columns(Column::exact(string_width * 0.6), 1) // Naziv materiala
                .columns(Column::exact(number_width), 2) // nabavna_skupina, mrp_karakteristika
                .columns(Column::exact(number_width), 5) // Zaloga, Poraba, Odprta narocila, Dobavni rok
                .columns(Column::exact(number_width * 1.8), 2) // trenutni zalogi
                .columns(Column::remainder(), 1) // Opomba
                .header(50.0, |mut header| {
                    header.col(|ui| {ui.heading("Material"); });
                    header.col(|ui| {ui.heading("Naziv"); });
                    header.col(|ui| {ui.heading("Nabavnik").on_hover_text("002 Neli\n008 Alenka/Viktorija\n010 Boštjan"); });
                    header.col(|ui| {ui.heading("MRP"); });
                    header.col(|ui| {ui.heading("Zaloga").on_hover_text("Trenutna zaloga v SAP-u"); });
                    header.col(|ui| {ui.heading("Poraba 3M").on_hover_text("Povprečna mesečna poraba za zadnje 3 mesece"); });
                    header.col(|ui| {ui.heading("Poraba 12M").on_hover_text("Povprečna mesečna poraba za zadnjih 12 mesecev"); });
                    header.col(|ui| {ui.heading("Odprto").on_hover_text("Odprta naročila dobaviteljem"); });
                    header.col(|ui| {ui.heading("Dobava").on_hover_text("Predviden dobavni rok v mesecih"); });
                    header.col(|ui| {ui.heading("Zaloga SAP").on_hover_text("Trenutna zaloga v SAP-u, ki zadostuje za X mesecev na osnovi povprečne porabe preteklih 3 mesecev"); });
                    header.col(|ui| {ui.heading("Zaloga SAP in odprto").on_hover_text("Seštevek trenutne zaloge v SAP-u in odprtih naročil, ki zadostuje za X mesecev na osnovi povprečne porabe preteklih 3 mesecev"); });
                    header.col(|ui| {ui.heading("Opomba"); });
                })
                .body(|body| {
                    body.rows(25., data.len(), |mut table_row| {
                        let index = table_row.index().clone();

                        let row = &data[index];

                        let mut row_color = Color32::TRANSPARENT;
                        if !row.dobavni_rok.is_none() {
                            if row.odprta_narocila.is_some_and(|v| v != 0.) {
                                if row.trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev.unwrap_or(0.) <= row.dobavni_rok.unwrap_or(0.) {
                                    row_color = Color32::LIGHT_BLUE;
                                } else {
                                    row_color = Color32::GREEN;
                                }
                            }


                            if row.trenutna_zaloga_zadostuje_za_mesecev.unwrap_or(0.) - row.dobavni_rok.unwrap_or(0.) < 3. &&
                                row.odprta_narocila.is_some_and(|v| v == 0.) {
                                row_color = Color32::YELLOW;
                            }

                            if row.trenutna_zaloga_zadostuje_za_mesecev.unwrap_or(0.) - row.dobavni_rok.unwrap_or(0.) < 1.5 &&
                                row.odprta_narocila.is_some_and(|v| v == 0.) {
                                row_color = Color32::ORANGE;
                            }

                            if row.trenutna_zaloga_zadostuje_za_mesecev.unwrap_or(0.) - row.dobavni_rok.unwrap_or(0.) < 0.5 &&
                                row.odprta_narocila.is_some_and(|v| v == 0.) {
                                row_color = Color32::RED;
                            }
                        }

                        /*
                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(format!("{}", index + 1));
                        });

                         */

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.material.to_string());
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.naziv_materiala.clone().unwrap_or_else(|| "".to_string()));
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.nabavna_skupina.clone().unwrap_or_else(|| "".to_string()));
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.mrp_karakteristika.clone().unwrap_or_else(|| "".to_string()));
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.zaloga.map_or("".to_string(), |v| format_number_custom(v)));
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.poraba_3m.map_or("".to_string(), |v| format_number_custom(v)));
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.poraba_12m.map_or("".to_string(), |v| format_number_custom(v)));
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.odprta_narocila.map_or("".to_string(), |v| format_number_custom(v)));
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.dobavni_rok.map_or("".to_string(), |v| format_number_custom(v)));
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.trenutna_zaloga_zadostuje_za_mesecev.map_or("".to_string(), |v| format_number_custom(v)));
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev.map_or("".to_string(), |v| format_number_custom(v)));
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            let t = row.opomba.clone().unwrap_or_else(|| "".to_string());
                            ui.label(&t).on_hover_text(t);
                        });



                    });
                });
        });
    }
}

fn format_number_custom(value: f64) -> String {
    let s = format!("{:.2}", value);
    let parts = s.split('.').collect::<Vec<_>>();
    if parts.len() == 2 {
        let int_part = parts[0];
        let dec_part = parts[1];
        let int_part_with_sep = int_part.chars()
            .rev()
            .collect::<Vec<_>>()
            .chunks(3)
            .map(|chunk| chunk.iter().rev().collect::<String>())
            .collect::<Vec<_>>()
            .into_iter()
            .rev()
            .collect::<Vec<_>>()
            .join(".");
        format!("{},{}", int_part_with_sep, dec_part)
    } else {
        s
    }
}

impl eframe::App for App {
    fn update(&mut self, ctx: &Context, _frame: &mut eframe::Frame) {
        //ctx.request_repaint();

        CentralPanel::default().show(ctx, |ui| {
            ui.vertical(|ui| {
                if let Some(success) = self.successfully_loaded_query {
                    let color = if success { Color32::GREEN } else { Color32::RED };
                    ui.colored_label(color, "■");
                }
                if ui.button("Naloži").clicked() {
                    match self.db_manager.get_data() {
                        Ok(rows) => {
                            self.row_data = Some(rows);
                            self.successfully_loaded_query = Some(true);
                        }
                        Err(err) => {
                            self.successfully_loaded_query = Some(false);
                            println!("query_load error: {:?}", err);
                        }
                    }
                }

                ui.horizontal(|ui| {
                    if ui.button("Izvozi").clicked() {
                        let mut resp: Result<(), Box<dyn std::error::Error>> = Ok(());

                        let _ = self.row_data.as_ref().map(|d| { resp = export_filtered_to_excel(&self.apply_filters(&d)); });
                        self.export_counter = Some(Instant::now());
                        match resp {
                            Err(err) => {
                                println!("export error: {:?}", err);
                                self.successfully_exported = Some(false);
                            },
                            Ok(_) => {
                                self.successfully_exported = Some(true);
                            }
                        }
                    }

                    if self.successfully_exported.is_some() && self.export_counter.is_some_and(|start| start.elapsed() < Duration::from_secs(3)){
                        if self.successfully_exported.unwrap() {
                            ui.colored_label(Color32::GREEN, "Izvozil!");
                        } else {
                            ui.colored_label(Color32::GREEN, "Napaka!");
                        }
                    } else {
                        self.export_counter = None;
                    }
                });

            });

            println!("Before filters: {}", self.row_data.as_ref().unwrap_or(&Vec::new()).len());
            let data = match &self.row_data {
                Some(d) => self.apply_filters(d),
                None => Vec::new(),
            };


            ui.label(format!("Število zadetkov: {}", &data.len()));

            ui.add_space(4.0);

            ScrollArea::vertical().show(ui, |ui| {
                self.render_table(ui, &data);
            });
        });


        Window::new("Filtri")
            .resizable(false)
            .collapsible(true)
            .default_open(false)
            .fixed_pos(pos2(80., 5.))
            .order(Order::Middle)
            .default_size(vec2(300., 400.))
            .show(ctx, |ui| {
                ui.vertical(|ui| {
                    ui.add(
                        TextEdit::singleline(&mut self.filter_sifra_materiala)
                            .hint_text("Iskanje po šifri materiala...")
                    );
                    ui.add(
                        TextEdit::singleline(&mut self.filter_naziv_materiala)
                            .hint_text("Iskanje po nazivu materiala...")
                    );
                    ui.add(
                        TextEdit::singleline(&mut self.filter_nabavnik)
                            .hint_text("Iskanje po nabavniku...")
                    );

                    ui.horizontal(|ui| {
                        ui.checkbox(&mut self.filter_zaloga_vecja, "zaloga večja kot");
                        ui.add(
                            TextEdit::singleline(&mut self.filter_zaloga_gt)
                                .desired_width(50.)
                        );
                    });

                    ui.horizontal(|ui| {
                        ui.checkbox(&mut self.filter_poraba_vecja, "poraba večja kot");
                        ui.add(
                            TextEdit::singleline(&mut self.filter_poraba_gt)
                                .desired_width(50.)
                        );
                    });


                    ui.horizontal(|ui| {
                        ui.checkbox(&mut self.filter_nabavnik_definiran, "Nabavnik je definiran");
                    });
                    ui.horizontal(|ui| {
                        ui.checkbox(&mut self.filter_nabavnik_ni_definiran, "Nabavnik ni definiran");
                    });



                    ui.horizontal(|ui| {
                        ui.checkbox(&mut self.filter_odprta_narocila, "odprta naročila večja kot");
                        ui.add(
                            TextEdit::singleline(&mut self.filter_odprta_narocila_gt)
                                .desired_width(50.)
                        );
                    });

                    ui.horizontal(|ui| {
                        ui.checkbox(&mut self.filter_dobavni_rok, "dobavni rok večji kot");
                        ui.add(
                            TextEdit::singleline(&mut self.filter_dobavni_rok_gt)
                                .desired_width(50.)
                        );
                    });

                    ui.checkbox(&mut self.filter_opomba, "Artikel ima opombo");

                    ui.checkbox(&mut self.filter_rumena, "Rumeni");
                    ui.checkbox(&mut self.filter_oranzna, "Oranžni");
                    ui.checkbox(&mut self.filter_rdeca, "Rdeči");
                    ui.checkbox(&mut self.filter_modra_zelena, "Modri/Zeleni");


                    let reset = ui.button("Ponastavi filtre");
                    if reset.clicked() {
                        self.filter_sifra_materiala = String::new();
                        self.filter_naziv_materiala = String::new();
                        self.filter_nabavnik = String::new();
                        self.filter_zaloga_vecja = true;
                        self.filter_zaloga_gt = String::from("0");

                        self.filter_poraba_vecja = true;
                        self.filter_poraba_gt = String::from("0");

                        self.filter_nabavnik_definiran = true;
                        self.filter_nabavnik_ni_definiran = false;

                        self.filter_odprta_narocila = false;
                        self.filter_odprta_narocila_gt = String::from("0");

                        self.filter_dobavni_rok = false;
                        self.filter_dobavni_rok_gt = String::from("0");

                        self.filter_opomba = false;

                        self.filter_rumena = false;
                        self.filter_oranzna = false;
                        self.filter_rdeca = false;
                        self.filter_modra_zelena = false;
                    }


                });

            });

        Window::new("Dobavni rok")
            .resizable(false)
            .collapsible(true)
            .default_open(false)
            .fixed_pos(pos2(185., 5.))
            .order(Order::Middle)
            .default_size(vec2(300., 200.))
            .show(ctx, |ui| {
                ui.vertical_centered(|ui| {
                    ui.add(
                        TextEdit::singleline(&mut self.dobavni_rok_material)
                            .hint_text("Šifra materiala...")
                    );
                    ui.add(
                        TextEdit::singleline(&mut self.dobavni_rok_dobavni_rok)
                            .hint_text("Dobavni rok...")
                    );

                    let dobavni_rok = ui.button("Shrani dobavni rok");
                    if dobavni_rok.clicked() {
                        let resp = self.db_manager.store_dobavni_rok((
                            self.dobavni_rok_material.parse().unwrap_or(0),
                            parse_string_to_optional_f64(self.dobavni_rok_dobavni_rok.as_str()),
                        ));
                        match resp {
                            Err(err) => {
                                println!("error storing opomba: {:?}", err.to_string());
                                self.successfully_stored_dobavni_rok = Some(false);
                            },
                            Ok(_) => {
                                println!("stored opomba!");
                                self.successfully_stored_dobavni_rok = Some(true);
                            }
                        }

                        let result = self.db_manager.get_data();

                        match result {
                            Err(err) => {
                                println!("query_load error after dobavni rok update: {:?}", err.to_string());
                                self.successfully_loaded_query = Some(false);
                            },
                            Ok(rows) => {
                                self.row_data = Some(rows);
                                println!("row_data loaded: {}", self.row_data.as_ref().unwrap().len());
                                self.successfully_loaded_query = Some(true);
                            }
                        }
                    }
                    if self.successfully_stored_dobavni_rok.is_some() {
                        if self.successfully_stored_dobavni_rok.unwrap() {
                            ui.colored_label(Color32::GREEN, "Shranil");
                        } else {
                            ui.colored_label(Color32::RED, "Napaka pri shranjevanju dobavnega roka!");
                        }
                    }

                });

            });


        Window::new("Opombe")
            .resizable(false)
            .collapsible(true)
            .default_open(false)
            .fixed_pos(pos2(350., 5.))
            .order(Order::Middle)
            .default_size(vec2(300., 200.))
            .show(ctx, |ui| {
                ui.vertical_centered(|ui| {
                    ui.add(
                        TextEdit::singleline(&mut self.opomba_material)
                            .hint_text("Šifra materiala...")
                    );
                    ui.add(
                        TextEdit::singleline(&mut self.opomba_opomba)
                            .hint_text("Opomba...")
                    );

                    let store_opomba = ui.button("Shrani opombo");
                    if store_opomba.clicked() {


                        let resp = self.db_manager.store_opomba_to_db((self.opomba_material.parse().unwrap_or(0), self.opomba_opomba.clone()));
                        match resp {
                            Err(err) => {
                                println!("error storing opomba: {:?}", err.to_string());
                                self.successfully_stored_opomba = Some(false);
                            },
                            Ok(_) => {
                                println!("stored opomba!");
                                self.successfully_stored_opomba = Some(true);
                            }
                        }

                        let result = self.db_manager.get_data();

                        match result {
                            Err(err) => {
                                println!("query_load error after opomba update: {:?}", err.to_string());
                                self.successfully_loaded_query = Some(false);
                            },
                            Ok(rows) => {
                                self.row_data = Some(rows);
                                println!("row_data loaded: {}", self.row_data.as_ref().unwrap().len());
                                self.successfully_loaded_query = Some(true);
                            }
                        }
                    }
                    if self.successfully_stored_opomba.is_some() {
                        if self.successfully_stored_opomba.unwrap() {
                            ui.colored_label(Color32::GREEN, "Shranil");
                        } else {
                            ui.colored_label(Color32::RED, "Napaka pri shranjevanju opombe!");
                        }
                    }

                });

            });




        Window::new("Vnos Excel-ov")
            .resizable(false)
            .collapsible(true)
            .default_open(false)
            .fixed_pos(pos2(490., 5.))
            .order(Order::Middle)
            .default_size(vec2(300., 200.))
            .show(ctx, |ui| {
                let sifrant_button = ui.button("Vnos šifrant");
                let import_button = ui.button("Vnos Poraba 12M, Poraba 3M, Zaloga, Odprta naročila");
                let mut file_input_error: Option<Box<dyn std::error::Error>> = None;
                if import_button.clicked() {
                    let downloads_dir = dirs_next::download_dir().unwrap_or_else(|| std::path::PathBuf::from("C:\\"));
                    let files = rfd::FileDialog::new()
                        .set_title("Izberi 4 datoteke!")
                        .set_directory(downloads_dir)
                        .pick_files();
                    if files.is_some() {
                        let result = parse_import_files(files.unwrap());
                        match result {
                            Ok(row_data) => {
                                self.retry_import = false;
                                self.successfully_parsed = Some(true);
                                let db_result = self.db_manager.store_to_data(row_data);
                                match db_result {
                                    Ok(_) => {
                                        self.successfully_stored_data = Some(true);
                                    },
                                    Err(err) => {
                                        println!("data: {:?}", err);
                                        self.successfully_stored_data = Some(false);
                                    }
                                }
                            },
                            Err(e) => {
                                self.retry_import = true;
                                self.successfully_parsed = Some(false);
                                file_input_error = Some(e);
                            }
                        }
                    } else {
                        self.retry_import = true;

                    }
                }


                if sifrant_button.clicked() {
                    let downloads_dir = dirs_next::download_dir().unwrap_or_else(|| std::path::PathBuf::from("C:\\"));
                    let file = rfd::FileDialog::new()
                        .set_title("Izberi 1 datoteko!")
                        .set_directory(downloads_dir)
                        .pick_file();
                    if file.is_some() {
                        let result = parse_sifrant_file(file.unwrap());
                        match result {
                            Ok(rows) => {
                                let db_result = self.db_manager.store_sifrant_to_db(rows);
                                match db_result {
                                    Ok(_) => {
                                        self.successfully_stored_sifrant = Some(true);
                                    },
                                    Err(err) => {
                                        println!("Šifrant: {:?}", err);
                                        self.successfully_stored_sifrant = Some(false);
                                    }
                                }
                            },
                            Err(e) => {
                                file_input_error = Some(e);
                            }
                        }
                    } else {
                        self.retry_import = true;
                    }
                }



                if self.retry_import {
                    if file_input_error.is_none() {
                        ui.colored_label(Color32::RED, "Nepravilen vnos datotek");
                    } else {
                        ui.colored_label(Color32::RED, file_input_error.unwrap().to_string());
                    }
                }
                if self.successfully_parsed.is_some() {
                    if self.successfully_parsed.unwrap() {
                        ui.colored_label(Color32::GREEN, "Brez napak v Excel-u!");
                    } else {
                        ui.colored_label(Color32::ORANGE, "Napaka v Excel-u");
                    }
                }
                if self.successfully_stored_data.is_some() {
                    if self.successfully_stored_data.unwrap() {
                        ui.colored_label(Color32::GREEN, "Uspešno shranil v bazo 3 Excel-ov");
                    } else {
                        ui.colored_label(Color32::ORANGE, "Neuspešno shranjevanje v bazo 3 Excel-ov");
                    }
                }

                if self.successfully_stored_sifrant.is_some() {
                    if self.successfully_stored_sifrant.unwrap() {
                        ui.colored_label(Color32::GREEN, "Uspešno shranil v bazo šifranta");
                    } else {
                        ui.colored_label(Color32::ORANGE, "Neuspešno shranjevanje v bazo šifranta");
                    }
                }


                let delete = ui.button("Izbriši baze").on_hover_text("Pred posodabljanjem!");
                if delete.clicked() {
                    let _ = self.db_manager.drop_all_tables();
                }
            });
    }
}


pub fn export_filtered_to_excel(
    data: &[ViewQuery],
) -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();



    worksheet.write_string(0, 0, "Material")?;
    worksheet.write_string(0, 1, "Naziv")?;
    worksheet.write_string(0, 2, "Nabavnik")?;
    worksheet.write_string(0, 3, "MRP")?;
    worksheet.write_string(0, 4, "Zaloga")?;
    worksheet.write_string(0, 5, "Poraba, Povprečna mesečna poraba za zadnje 3 mesece")?;
    worksheet.write_string(0, 5, "Poraba, Povprečna mesečna poraba za zadnjih 12 mesecev")?;
    worksheet.write_string(0, 6, "Odprto, Odprta naročila dobaviteljem")?;
    worksheet.write_string(0, 7, "Dobava, Predviden dobavni rok v mesecih")?;
    worksheet.write_string(0, 8, "Zaloga SAP, Trenutna zaloga v SAP-u, ki zadostuje za X mesecev na osnovi povprečne porabe preteklih 3 mesecev")?;
    worksheet.write_string(0, 9, "Zaloga SAP in odprto, Seštevek trenutne zaloge v SAP-u in odprtih naročil, ki zadostuje za X mesecev na osnovi povprečne porabe preteklih 3 mesecev")?;
    worksheet.write_string(0, 10, "Opomba")?;


    for (row_idx, item) in data.iter().enumerate() {
        let row = (row_idx + 1) as u32;

        worksheet.write_number(row, 0, item.material as f64)?;

        match &item.naziv_materiala {
            Some(s) => worksheet.write_string(row, 1, s)?,
            None => worksheet.write_blank(row, 1, &Format::default())?,
        };

        match &item.nabavna_skupina {
            Some(s) => worksheet.write_string(row, 2, s)?,
            None => worksheet.write_blank(row, 2, &Format::default())?,
        };

        match &item.mrp_karakteristika {
            Some(s) => worksheet.write_string(row, 3, s)?,
            None => worksheet.write_blank(row, 3, &Format::default())?,
        };

        match item.zaloga {
            Some(v) => worksheet.write_number(row, 4, v)?,
            None => worksheet.write_blank(row, 4, &Format::default())?,
        };

        match item.poraba_3m {
            Some(v) => worksheet.write_number(row, 5, v)?,
            None => worksheet.write_blank(row, 5, &Format::default())?,
        };

        match item.poraba_12m {
            Some(v) => worksheet.write_number(row, 6, v)?,
            None => worksheet.write_blank(row, 6, &Format::default())?,
        };

        match item.odprta_narocila {
            Some(v) => worksheet.write_number(row, 7, v)?,
            None => worksheet.write_blank(row, 7, &Format::default())?,
        };

        match item.dobavni_rok {
            Some(v) => worksheet.write_number(row, 8, v)?,
            None => worksheet.write_blank(row, 8, &Format::default())?,
        };

        match item.trenutna_zaloga_zadostuje_za_mesecev {
            Some(v) => worksheet.write_number(row, 9, v)?,
            None => worksheet.write_blank(row, 9, &Format::default())?,
        };

        match item.trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev {
            Some(v) => worksheet.write_number(row, 10, v)?,
            None => worksheet.write_blank(row, 10, &Format::default())?,
        };

        match &item.opomba {
            Some(s) => worksheet.write_string(row, 11, s)?,
            None => worksheet.write_blank(row, 11, &Format::default())?,
        };
    }

    workbook.save("Analitika.xlsx")?;
    Ok(())
}


fn parse_string_to_optional_f64(s: &str) -> Option<f64> {
    if s.eq("") {
        None
    } else {
        Some(s.replace(",", ".").parse().unwrap_or(0.))
    }
}

fn main() {

    eframe::run_native(
        "Magneti Excel",
        NativeOptions::default(),
        Box::new(|cc| Ok(Box::new(App::new(cc))))
    ).unwrap();
}
