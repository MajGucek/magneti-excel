#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod parse;
mod db;

use eframe::{NativeOptions};
use eframe::egui::*;
use egui_extras::{Column, TableBuilder};
use rfd::MessageLevel;
use rust_xlsxwriter::{Format, Workbook};
use crate::db::{DBManager, SortColumn, SortState, ViewQuery};
use crate::parse::{parse_dobavitelji_file, parse_import_files, parse_sifrant_file};

struct App {
    db_manager: DBManager,

    row_data: Rows,
    successfully_loaded_query: Option<bool>,
    sort_state: SortState,


    opomba_material: String,
    opomba_opomba: String,

    dobavni_rok_material: String,
    dobavni_rok_dobavni_rok: String,

    min_zaloga_material: String,
    min_zaloga_min_zaloga: String,

    max_zaloga_material: String,
    max_zaloga_max_zaloga: String,

    blagovna_skupina_material: String,
    blagovna_skupina_blagovna_skupina: String,

    pakiranje_material: String,
    pakiranje_pakiranje: String,


    /* --Columns-- */



    /* --Filters-- */
    filter_rdeca: bool,
    filter_oranzna: bool,
    filter_rumena: bool,
    filter_modra: bool,
    filter_zelena: bool,
    filter_viola: bool,
    filter_teal: bool,
    filter_indigo: bool,

    filter_sifra_materiala: String,
    filter_naziv_materiala: String,
    filter_nabavnik: String,
    filter_dobavitelj: String,
    filter_blagovna_skupina: String,

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

    filter_min_zaloga: bool,
    filter_min_zaloga_gt: String,

    filter_max_zaloga: bool,
    filter_max_zaloga_gt: String,

    filter_opomba: bool,
    filter_no_dobavni_rok: bool,


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
        let sort_state = SortState::default();
        let _ = db_manager.try_drop_view();
        let _ = db_manager.try_create_view().unwrap();

        let result = db_manager.get_data(&sort_state);

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
            row_data: Rows {row_data},
            successfully_loaded_query,
            sort_state,


            opomba_material: String::new(),
            opomba_opomba: String::new(),

            dobavni_rok_material: String::new(),
            dobavni_rok_dobavni_rok: String::new(),

            min_zaloga_material: String::new(),
            min_zaloga_min_zaloga: String::new(),

            max_zaloga_material: String::new(),
            max_zaloga_max_zaloga: String::new(),

            blagovna_skupina_material: String::new(),
            blagovna_skupina_blagovna_skupina: String::new(),

            pakiranje_material: String::new(),
            pakiranje_pakiranje: String::new(),


            filter_rdeca: false,
            filter_oranzna: false,
            filter_rumena: false,
            filter_modra: false,
            filter_zelena: false,
            filter_viola: false,
            filter_teal: false,
            filter_indigo: false,


            filter_sifra_materiala: String::new(),
            filter_naziv_materiala: String::new(),
            filter_nabavnik: String::new(),
            filter_dobavitelj: String::new(),
            filter_blagovna_skupina: String::new(),

            filter_zaloga_vecja: false,
            filter_zaloga_gt: String::from("0"),

            filter_nabavnik_definiran: false,
            filter_nabavnik_ni_definiran: false,

            filter_poraba_vecja: false,
            filter_poraba_gt: String::from("0"),

            filter_odprta_narocila: false,
            filter_odprta_narocila_gt: String::from("0"),

            filter_dobavni_rok: true,
            filter_dobavni_rok_gt: String::from("0"),

            filter_min_zaloga: false,
            filter_min_zaloga_gt: String::from("0"),

            filter_max_zaloga: false,
            filter_max_zaloga_gt: String::from("0"),

            filter_opomba: false,
            filter_no_dobavni_rok: false,
        }
    }

}


struct Rows {
    row_data: Option<Vec<ViewQuery>>
}

impl Rows {
    fn query(&mut self, db_manager: &DBManager, sort_state: &SortState) -> Option<bool> {
        match db_manager.get_data(&sort_state) {
            Ok(rows) => {
                self.row_data = Some(rows);
                Some(true)
            }
            Err(err) => {
                println!("query_load error: {:?}", err);
                Some(false)
            }
        }
    }
}


impl App {



    fn apply_filters(&self, rows: &Vec<ViewQuery>) -> Vec<ViewQuery> {

        rows.iter()
            .filter(|&row| {
                // --Evil GPT hacked-- //
                let any_color_filter_active = self.filter_rumena || 
                    self.filter_oranzna || 
                    self.filter_rdeca || 
                    self.filter_zelena || 
                    self.filter_modra || 
                    self.filter_viola ||
                    self.filter_teal || 
                    self.filter_indigo
                    ;
                let mut color_matches = false;
                if any_color_filter_active {
                    let months_left = row.dobavni_rok.map_or(0., |dr| row.trenutna_zaloga_zadostuje_za_mesecev.unwrap_or(0.) - dr);
                    let no_open_orders = row.odprta_narocila.is_some_and(|v| v == 0.);

                    color_matches |= self.filter_rumena && row.dobavni_rok.is_some() && months_left >= 1.0 && months_left < 1.5 && no_open_orders;
                    color_matches |= self.filter_oranzna && row.dobavni_rok.is_some() && months_left >= 0.3 && months_left < 1.0 && no_open_orders;
                    color_matches |= self.filter_rdeca && row.dobavni_rok.is_some() && months_left < 0.3 && no_open_orders && row.dobavni_rok.unwrap_or(0.) < 90.;

                    color_matches |= self.filter_viola && row.dobavni_rok.is_some() && row.dobavni_rok.unwrap_or(0.) >= 90. &&
                        no_open_orders;

                    color_matches |= self.filter_modra && row.dobavni_rok.is_some() && !no_open_orders &&
                        row.trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev.unwrap_or(0.) <= row.dobavni_rok.unwrap_or(0.);
                    color_matches |= self.filter_zelena && row.dobavni_rok.is_some() && !no_open_orders &&
                        !(row.trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev.unwrap_or(0.) <= row.dobavni_rok.unwrap_or(0.));

                    color_matches |= self.filter_teal && row.minimalna_zaloga.is_some_and(|val| val > row.zaloga.unwrap_or(0.));
                    color_matches |= self.filter_indigo && row.maximalna_zaloga.is_some_and(|val| val < row.zaloga.unwrap_or(0.));
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

                    (row.nabavna_skupina.as_ref().is_some_and(|a| format!("{}", a.to_lowercase()).contains(self.filter_nabavnik.to_lowercase().as_str())) ||
                        row.nabavna_skupina.as_ref().is_some_and(|a| format!("{}", format_nabavnik(a).unwrap_or(a).to_lowercase()).contains(self.filter_nabavnik.to_lowercase().as_str())
                        )
                    ) &&

                    row.blagovna_skupina.as_ref().unwrap_or(&String::new()).to_lowercase().contains(self.filter_blagovna_skupina.to_lowercase().as_str()) &&

                    row.dobavitelji.as_ref().is_some_and(|a| format!("{}", a.to_lowercase()).contains(self.filter_dobavitelj.to_lowercase().as_str())) &&
                    condition &&

                    (!self.filter_odprta_narocila || row.odprta_narocila.is_some_and(|odp| odp > parse_string_to_optional_f64(self.filter_odprta_narocila_gt.as_str()).unwrap_or(0.))) &&
                    (!self.filter_dobavni_rok || row.dobavni_rok.is_some_and(|dob| dob > parse_string_to_optional_f64(self.filter_dobavni_rok_gt.as_str()).unwrap_or(0.))) &&
                    (!self.filter_min_zaloga || row.minimalna_zaloga.is_some_and(|dob| dob > parse_string_to_optional_f64(self.filter_min_zaloga_gt.as_str()).unwrap_or(0.))) &&
                    (!self.filter_max_zaloga || row.maximalna_zaloga.is_some_and(|dob| dob > parse_string_to_optional_f64(self.filter_max_zaloga_gt.as_str()).unwrap_or(0.))) &&

                    (color_matches || !any_color_filter_active) &&
                    (!self.filter_nabavnik_ni_definiran || row.nabavna_skupina.as_ref().is_some_and(|str| str.eq(""))) &&
                    (!self.filter_nabavnik_definiran || !row.nabavna_skupina.as_ref().is_some_and(|str| str.eq(""))) &&
                    (!self.filter_opomba || row.opomba.as_ref().is_some_and(|s| !s.eq(""))) &&
                    (!self.filter_no_dobavni_rok || row.dobavni_rok.is_none())
        })
            .cloned()
            .collect()


    }


    pub fn render_table(&self, ui: &mut Ui, data: &Vec<ViewQuery>) {
        ScrollArea::both().show(ui, |ui| {
            ui.style_mut().visuals.faint_bg_color = Color32::from_rgb(200, 200, 200);
            TableBuilder::new(ui)
                .striped(true)
                .cell_layout(Layout::left_to_right(Align::Center))
                .columns(Column::exact(80.), 1) // Material
                .columns(Column::exact(300.), 1)// Naziv
                .columns(Column::exact(85.), 1) // Zaloga
                .columns(Column::exact(100.5), 1)// Poraba 3M
                .columns(Column::exact(100.5), 1)// Poraba 24M
                .columns(Column::exact(90.), 1)// Odprto
                .columns(Column::exact(90.), 1)// Dobava
                .columns(Column::exact(110.), 1)// Zaloga SAP
                .columns(Column::exact(120.), 1)// Sum Zaloga
                .columns(Column::exact(120.), 1)// Enota
                .columns(Column::exact(120.), 1)// Minimalna zaloga
                .columns(Column::exact(120.), 1)// Maximalna zaloga
                .columns(Column::exact(120.), 1)// Pakiranje
                .columns(Column::exact(120.), 1)// Blagovna Skupina
                .columns(Column::exact(300.), 1)// Opomba
                .columns(Column::exact(90.), 1)// Nabavnik
                .columns(Column::remainder(), 1)//Dobavitelji

                .header(50.0, |mut header| {
                    header.col(|ui| {ui.heading("Material"); });
                    header.col(|ui| {ui.heading("Naziv"); });


                    header.col(|ui| {ui.heading("Zaloga").on_hover_text("Trenutna zaloga v SAP-u"); });
                    header.col(|ui| {ui.heading("Poraba 3M").on_hover_text("Povprečna mesečna poraba za zadnje 3 mesece"); });
                    header.col(|ui| {ui.heading("Poraba 24M").on_hover_text("Povprečna mesečna poraba za zadnjih 12 mesecev"); });
                    header.col(|ui| {ui.heading("Odprto").on_hover_text("Odprta naročila dobaviteljem"); });
                    header.col(|ui| {ui.heading("Dobava").on_hover_text("Predviden dobavni rok v mesecih"); });

                    header.col(|ui| {ui.heading("Zaloga SAP").on_hover_text("Trenutna zaloga v SAP-u, ki zadostuje za X mesecev na osnovi povprečne porabe preteklih 3 mesecev, če artikel nima 3M porabe računa na osnovi 24M porabe"); });

                    header.col(|ui| {ui.heading("Sum Zaloga").on_hover_text("Seštevek trenutne zaloge v SAP-u in odprtih naročil, ki zadostuje za X mesecev na osnovi povprečne porabe preteklih 3 mesecev, če artikel nima 3M porabe računa na osnovi 24M porabe"); });
                    header.col(|ui| {ui.heading("Enota"); });
                    header.col(|ui| {ui.heading("Min Zaloga"); });
                    header.col(|ui| {ui.heading("Max Zaloga"); });
                    header.col(|ui| {ui.heading("Pakiranje"); });
                    header.col(|ui| {ui.heading("Blagovna Skupina"); });
                    header.col(|ui| {ui.heading("Opomba"); });
                    header.col(|ui| {ui.heading("Nabavnik").on_hover_text("002 Neli\n008 Viktorija\n010 Boštjan"); });
                    header.col(|ui| {ui.heading("Dobavitelji"); });
                    //header.col(|ui| {ui.heading("MRP"); });
                })
                .body(|body| {
                    body.rows(25., data.len(), |mut table_row| {
                        let index = table_row.index().clone();

                        let row = &data[index];

                        let mut row_color = Color32::TRANSPARENT;
                        if row.dobavni_rok.is_some() {
                            if row.odprta_narocila.is_some_and(|v| v != 0.) {
                                if row.trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev.unwrap_or(0.) <= row.dobavni_rok.unwrap_or(0.) {
                                    row_color = Color32::LIGHT_BLUE;
                                } else {
                                    row_color = Color32::LIGHT_GREEN;
                                }
                            }

                            if row.odprta_narocila.is_some_and(|v| v == 0.) {
                                if row.trenutna_zaloga_zadostuje_za_mesecev.unwrap_or(0.) - row.dobavni_rok.unwrap_or(0.) < 1.5 {
                                    // yellow
                                    row_color = Color32::YELLOW;
                                }

                                if row.trenutna_zaloga_zadostuje_za_mesecev.unwrap_or(0.) - row.dobavni_rok.unwrap_or(0.) < 1.0 {
                                    // orange
                                    row_color = Color32::from_rgb(255, 153, 51);
                                }

                                if row.trenutna_zaloga_zadostuje_za_mesecev.unwrap_or(0.) - row.dobavni_rok.unwrap_or(0.) < 0.3 {
                                    // red
                                    row_color = Color32::from_rgb(255, 135, 135);
                                }

                                if row.dobavni_rok.unwrap_or(0.) >= 90. {
                                    // VIOLET
                                    row_color = Color32::from_rgb(191, 136, 187);
                                }
                            }
                        }






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
                            ui.label(row.zaloga.map_or("".to_string(), |v| format_number_custom(v)));
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            let poraba_3m = row.poraba_3m.map_or("".to_string(), |v| format_number_custom(v));
                            let poraba_24m = row.poraba_24m.map_or("".to_string(), |v| format_number_custom(v));

                            let (arrow, color) = if poraba_3m > poraba_24m {
                                ("🔺", Color32::BLACK)
                            } else if !poraba_3m.eq("0,00") && !poraba_24m.eq("0,00") && !poraba_3m.eq(poraba_24m.as_str()) {
                                ("🔻", Color32::BLACK)
                            } else if poraba_3m.eq("0,00") && !poraba_24m.eq("0,00") {
                                ("🔻", Color32::BLACK)
                            } else {
                                ("     ", Color32::TRANSPARENT)
                            };


                            ui.colored_label(color, arrow);
                            ui.label(poraba_3m);
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.poraba_24m.map_or("".to_string(), |v| format_number_custom(v)));
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
                            let t = row.osnovna_merska_enota.clone().unwrap_or_else(|| "".to_string());
                            ui.label(&t);
                        });

                        table_row.col(|ui| {
                            let old = row_color;
                            if row.minimalna_zaloga.is_some_and(|val| val > row.zaloga.unwrap_or(0.))  {
                                // teal
                                row_color = Color32::from_rgb(0, 150, 136);
                            }
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.minimalna_zaloga.map_or("".to_string(), |v| format_number_custom(v)));
                            row_color = old;
                        });
                        table_row.col(|ui| {
                            let old = row_color;
                            if row.maximalna_zaloga.is_some_and(|val| val < row.zaloga.unwrap_or(0.)) {
                                // indigo
                                row_color = Color32::from_rgb(153,51,255);
                            }

                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.maximalna_zaloga.map_or("".to_string(), |v| format_number_custom(v)));
                            row_color = old;
                        });


                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            let t = row.pakiranje.clone().unwrap_or_else(|| "".to_string());
                            ui.label(&t);
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            let t = row.blagovna_skupina.clone().unwrap_or_else(|| "".to_string());
                            ui.label(&t);
                        });


                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            let t = row.opomba.clone().unwrap_or_else(|| "".to_string());
                            ui.label(&t);//.on_hover_text(t);
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            let nabavna_skupina = row.nabavna_skupina.clone().unwrap_or_else(|| "".to_string());

                            ui.label(format_nabavnik(nabavna_skupina.as_str()).unwrap_or(nabavna_skupina.as_str()));
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            let t = row.dobavitelji.clone().unwrap_or_else(|| "".to_string());
                            ui.label(&t);
                        });

                        /*
                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.mrp_karakteristika.clone().unwrap_or_else(|| "".to_string()));
                        });

                         */

                    });
                });
        });
    }
}

fn format_nabavnik<'a>(nabavna_skupina: &str) -> Option<&'a str> {
    match nabavna_skupina {
            "002" => Some("Neli"),
            "008" => Some("Viktoria"),
            "010" => Some("Boštjan"),
            _ => None,
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
        ctx.request_repaint_after(std::time::Duration::from_millis(100));


        CentralPanel::default().show(ctx, |ui| {
            ui.vertical(|ui| {
                if let Some(success) = self.successfully_loaded_query {
                    let color = if success { Color32::GREEN } else { Color32::RED };
                    ui.colored_label(color, "■");
                }
                if ui.button("Naloži").clicked() {
                    self.row_data.query(&self.db_manager, &self.sort_state);
                }

                ui.horizontal(|ui| {
                    if ui.button("Izvozi").clicked() {
                        let mut resp: Result<(), Box<dyn std::error::Error>> = Ok(());

                        let _ = self.row_data.row_data.as_ref().map(|d| { resp = export_filtered_to_excel(&self.apply_filters(&d)); });
                        match resp {
                            Err(err) => {
                                println!("export error: {:?}", err);
                                rfd::MessageDialog::new().set_title("Napaka").set_description("Napaka pri izvozu").set_level(MessageLevel::Error).show();
                            },
                            Ok(_) => {
                                rfd::MessageDialog::new().set_title("Uspeh").set_description("Uspešno izvozil").set_level(MessageLevel::Info).show();
                            }
                        }
                    }
                });

            });

            let data = match &self.row_data.row_data {
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

                    ui.add(
                        TextEdit::singleline(&mut self.filter_dobavitelj)
                            .hint_text("Iskanje po dobavitelju...")
                    );

                    ui.add(
                        TextEdit::singleline(&mut self.filter_blagovna_skupina)
                            .hint_text("Iskanje po blagovni skupini...")
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

                    ui.horizontal(|ui| {
                        ui.checkbox(&mut self.filter_min_zaloga, "Min zaloga večja kot");
                        ui.add(
                            TextEdit::singleline(&mut self.filter_min_zaloga_gt)
                                .desired_width(50.)
                        );
                    });

                    ui.horizontal(|ui| {
                        ui.checkbox(&mut self.filter_max_zaloga, "Max zaloga večja kot");
                        ui.add(
                            TextEdit::singleline(&mut self.filter_max_zaloga_gt)
                                .desired_width(50.)
                        );
                    });

                    ui.checkbox(&mut self.filter_opomba, "Artikel ima opombo");
                    ui.checkbox(&mut self.filter_no_dobavni_rok, "Nima dobavnega roka");


                    ui.collapsing("Barve", |ui| {
                        ui.vertical(|ui| {
                            ui.checkbox(&mut self.filter_rumena, "Rumeni");
                            ui.checkbox(&mut self.filter_oranzna, "Oranžni");
                            ui.checkbox(&mut self.filter_rdeca, "Rdeči");
                            ui.checkbox(&mut self.filter_modra, "Modri");
                            ui.checkbox(&mut self.filter_zelena, "Zeleni");
                            ui.checkbox(&mut self.filter_viola, "Viola");
                            ui.checkbox(&mut self.filter_teal, "Smaragdna");
                            ui.checkbox(&mut self.filter_indigo, "Indigo");
                        });
                    });




                    let reset = ui.button("Ponastavi filtre");
                    if reset.clicked() {
                        self.filter_sifra_materiala = String::new();
                        self.filter_naziv_materiala = String::new();
                        self.filter_nabavnik = String::new();
                        self.filter_dobavitelj = String::new();
                        self.filter_zaloga_vecja = false;
                        self.filter_zaloga_gt = String::from("0");
                        self.filter_blagovna_skupina = String::new();

                        self.filter_poraba_vecja = false;
                        self.filter_poraba_gt = String::from("0");

                        self.filter_nabavnik_definiran = false;
                        self.filter_nabavnik_ni_definiran = false;

                        self.filter_odprta_narocila = false;
                        self.filter_odprta_narocila_gt = String::from("0");

                        self.filter_dobavni_rok = true;
                        self.filter_dobavni_rok_gt = String::from("0");

                        self.filter_min_zaloga = false;
                        self.filter_min_zaloga_gt = String::from("0");

                        self.filter_max_zaloga = false;
                        self.filter_max_zaloga_gt = String::from("0");

                        self.filter_opomba = false;
                        self.filter_no_dobavni_rok = false;

                        self.filter_rumena = false;
                        self.filter_oranzna = false;
                        self.filter_rdeca = false;
                        self.filter_modra = false;
                        self.filter_zelena = false;
                        self.filter_viola = false;
                        self.filter_teal = false;
                        self.filter_indigo = false;


                        self.sort_state = SortState::default();
                        self.row_data.query(&self.db_manager, &self.sort_state);
                    }
                });
            });


        Window::new("Ročni Vnosi")
            .resizable(false)
            .collapsible(true)
            .default_open(false)
            .fixed_pos(pos2(185., 5.))
            .order(Order::Middle)
            .default_size(vec2(300., 200.))
            .show(ctx, |ui| {
                let spacing = 40.;


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
                                rfd::MessageDialog::new().set_title("Napaka").set_description("Napaka pri shranjevanju").set_level(MessageLevel::Error).show();
                            },
                            Ok(_) => {
                                println!("stored opomba!");
                                rfd::MessageDialog::new().set_title("Uspeh").set_description("Uspešno shranil dobavni rok").set_level(MessageLevel::Info).show();
                            }
                        }

                        self.row_data.query(&self.db_manager, &self.sort_state);
                    }


                    ui.add_space(spacing);



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
                                rfd::MessageDialog::new().set_title("Napaka").set_description("Napaka pri shranjevanju opombe").set_level(MessageLevel::Error).show();
                            },
                            Ok(_) => {
                                println!("stored opomba!");
                                rfd::MessageDialog::new().set_title("Uspeh").set_description("Uspešno shranil opombo").set_level(MessageLevel::Info).show();
                            }
                        }

                        self.row_data.query(&self.db_manager, &self.sort_state);
                    }



                    ui.add_space(spacing);



                    ui.add(
                        TextEdit::singleline(&mut self.min_zaloga_material)
                            .hint_text("Šifra materiala...")
                    );
                    ui.add(
                        TextEdit::singleline(&mut self.min_zaloga_min_zaloga)
                            .hint_text("Minimalna zaloga...")
                    );

                    let store_min_zaloga = ui.button("Shrani minimalno zalogo");
                    if store_min_zaloga.clicked() {
                        let resp = self.db_manager.store_min_zaloga((self.min_zaloga_material.parse().unwrap_or(0), self.min_zaloga_min_zaloga.clone().parse::<f64>().ok()));
                        match resp {
                            Err(err) => {
                                println!("error storing min zaloga: {:?}", err.to_string());
                                rfd::MessageDialog::new().set_title("Napaka").set_description("Napaka pri shranjevanju minimalne zaloge").set_level(MessageLevel::Error).show();
                            },
                            Ok(_) => {
                                println!("stored min zaloga!");
                                rfd::MessageDialog::new().set_title("Uspeh").set_description("Uspešno shranil minimalno zalogo").set_level(MessageLevel::Info).show();
                            }
                        }

                        self.row_data.query(&self.db_manager, &self.sort_state);
                    }



                    ui.add_space(spacing);



                    ui.add(
                        TextEdit::singleline(&mut self.max_zaloga_material)
                            .hint_text("Šifra materiala...")
                    );
                    ui.add(
                        TextEdit::singleline(&mut self.max_zaloga_max_zaloga)
                            .hint_text("Maximalna zaloga...")
                    );

                    let store_max_zaloga = ui.button("Shrani maximalno zalogo");
                    if store_max_zaloga.clicked() {
                        let resp = self.db_manager.store_max_zaloga((self.max_zaloga_material.parse().unwrap_or(0), self.max_zaloga_max_zaloga.clone().parse::<f64>().ok()));
                        match resp {
                            Err(err) => {
                                println!("error storing max zaloga: {:?}", err.to_string());
                                rfd::MessageDialog::new().set_title("Napaka").set_description("Napaka pri shranjevanju maximalno zalogo").set_level(MessageLevel::Error).show();
                            },
                            Ok(_) => {
                                println!("stored max zaloga!");
                                rfd::MessageDialog::new().set_title("Uspeh").set_description("Uspešno shranil maximalno zalogo").set_level(MessageLevel::Info).show();
                            }
                        }

                        self.row_data.query(&self.db_manager, &self.sort_state);
                    }



                    ui.add_space(spacing);



                    ui.add(
                        TextEdit::singleline(&mut self.blagovna_skupina_material)
                            .hint_text("Šifra materiala...")
                    );
                    ui.add(
                        TextEdit::singleline(&mut self.blagovna_skupina_blagovna_skupina)
                            .hint_text("Blagovna skupina...")
                    );

                    let store_blagovna_skupina = ui.button("Shrani blagovno skupino");
                    if store_blagovna_skupina.clicked() {
                        let resp = self.db_manager.store_blagovna_skupina((self.blagovna_skupina_material.parse().unwrap_or(0), self.blagovna_skupina_blagovna_skupina.clone()));
                        match resp {
                            Err(err) => {
                                println!("error storing blagovna skupina: {:?}", err.to_string());
                                rfd::MessageDialog::new().set_title("Napaka").set_description("Napaka pri shranjevanju blagovne skupine").set_level(MessageLevel::Error).show();
                            },
                            Ok(_) => {
                                println!("stored blagovna skupina!");
                                rfd::MessageDialog::new().set_title("Uspeh").set_description("Uspešno shranil blagovno skupino").set_level(MessageLevel::Info).show();
                            }
                        }

                        self.row_data.query(&self.db_manager, &self.sort_state);
                    }


                    ui.add_space(spacing);



                    ui.add(
                        TextEdit::singleline(&mut self.pakiranje_material)
                            .hint_text("Šifra materiala...")
                    );
                    ui.add(
                        TextEdit::singleline(&mut self.pakiranje_pakiranje)
                            .hint_text("Pakiranje...")
                    );

                    let store_pakiranje = ui.button("Shrani pakiranje");
                    if store_pakiranje.clicked() {
                        let resp = self.db_manager.store_pakiranje((self.pakiranje_material.parse().unwrap_or(0), self.pakiranje_pakiranje.clone()));
                        match resp {
                            Err(err) => {
                                println!("error storing pakiranje: {:?}", err.to_string());
                                rfd::MessageDialog::new().set_title("Napaka").set_description("Napaka pri shranjevanju pakiranja").set_level(MessageLevel::Error).show();
                            },
                            Ok(_) => {
                                println!("stored blagovna skupina!");
                                rfd::MessageDialog::new().set_title("Uspeh").set_description("Uspešno shranil pakiranje").set_level(MessageLevel::Info).show();
                            }
                        }

                        self.row_data.query(&self.db_manager, &self.sort_state);
                    }
                });

            });


        Window::new("Razvrstitev")
            .resizable(false)
            .collapsible(true)
            .default_open(false)
            .fixed_pos(pos2(345., 5.))
            .order(Order::Middle)
            .default_size(vec2(200., 150.))
            .show(ctx, |ui| {
                ui.horizontal(|ui| {
                    let old_column = self.sort_state.sort_column;

                    ComboBox::from_label("")
                        .selected_text(self.sort_state.sort_column.as_str())
                        .show_ui(ui, |ui| {
                            ui.selectable_value(&mut self.sort_state.sort_column, SortColumn::Material, "Material");
                            ui.selectable_value(&mut self.sort_state.sort_column, SortColumn::NazivMateriala, "Naziv");
                            ui.selectable_value(&mut self.sort_state.sort_column, SortColumn::NabavnaSkupina, "Nabavnik");
                            //ui.selectable_value(&mut self.sort_state.sort_column, SortColumn::MRP, "MRP");
                            ui.selectable_value(&mut self.sort_state.sort_column, SortColumn::Zaloga, "Zaloga");
                            ui.selectable_value(&mut self.sort_state.sort_column, SortColumn::Poraba3M, "Poraba 3M");
                            ui.selectable_value(&mut self.sort_state.sort_column, SortColumn::Poraba24M, "Poraba 24M");
                            ui.selectable_value(&mut self.sort_state.sort_column, SortColumn::OdprtaNarocila, "Odprto");
                            ui.selectable_value(&mut self.sort_state.sort_column, SortColumn::DobavniRok, "Dobava");
                            ui.selectable_value(&mut self.sort_state.sort_column, SortColumn::TrenutnaZalogaZadostujeZaMesecev, "Zaloga SAP");
                            ui.selectable_value(&mut self.sort_state.sort_column, SortColumn::TrenutnaZalogaInOdprtaNarocilaZadostujeZaMesecev, "Sum Zaloga");
                            ui.selectable_value(&mut self.sort_state.sort_column, SortColumn::OsnovnaMerskaEnota, "Enota");
                            ui.selectable_value(&mut self.sort_state.sort_column, SortColumn::Dobavitelji, "Dobavitelji");
                            ui.selectable_value(&mut self.sort_state.sort_column, SortColumn::MinimalnaZaloga, "Min zaloga");
                            ui.selectable_value(&mut self.sort_state.sort_column, SortColumn::MaximalnaZaloga, "Max zaloga");
                            ui.selectable_value(&mut self.sort_state.sort_column, SortColumn::Pakiranje, "Pakiranje");
                            ui.selectable_value(&mut self.sort_state.sort_column, SortColumn::BlagovnaSkupina, "Blagovna Skupina");
                            ui.selectable_value(&mut self.sort_state.sort_column, SortColumn::Opomba, "Opomba");
                        });
                    if old_column != self.sort_state.sort_column ||
                        ui.checkbox(&mut self.sort_state.descending, "Padajoče").changed()
                    {
                        self.row_data.query(&self.db_manager, &self.sort_state);
                    }
                });
            });



        Window::new("Vnos Excel-ov")
            .resizable(false)
            .collapsible(true)
            .default_open(false)
            .fixed_pos(pos2(500., 5.))
            .order(Order::Middle)
            .default_size(vec2(300., 200.))
            .show(ctx, |ui| {
                let sifrant_button = ui.button("Šifrant").on_hover_text("ŠIFRANT.XLSX");
                let import_button = ui.button("Poraba, Zaloga, Odprta naročila").on_hover_text("PORABA.XLSX, ZALOGA.XLSX, ODPRTA NAROČILA.XLSX");
                let dobavitelji_button = ui.button("Dobavitelji").on_hover_text("DOBAVITELJI.XLSX");
                if import_button.clicked() {
                    let downloads_dir = dirs_next::download_dir().unwrap_or_else(|| std::path::PathBuf::from("C:\\"));
                    let files = rfd::FileDialog::new()
                        .set_title("Izberi 4 datoteke!")
                        .set_directory(downloads_dir)
                        .pick_files();
                    if files.is_some() {
                        match self.db_manager.drop_data() {
                            Err(e) => println!("dropping error: {}", e),
                            Ok(_) => println!("Successfully dropped data")
                        }

                        let result = parse_import_files(files.unwrap());
                        match result {
                            Ok(row_data) => {
                                let db_result = self.db_manager.store_to_data(row_data);
                                match db_result {
                                    Ok(_) => {
                                        rfd::MessageDialog::new().set_title("Uspeh").set_description("Uspešno shranil podatke").set_level(MessageLevel::Info).show();
                                    },
                                    Err(err) => {
                                        println!("data: {:?}", err);
                                        rfd::MessageDialog::new().set_title("Napaka").set_description("Napaka pri shranjevanju podatkov").set_level(MessageLevel::Error).show();
                                    }
                                }
                            },
                            Err(_) => {
                                rfd::MessageDialog::new().set_title("Napaka").set_description("Napaka pri branju Excel-a").set_level(MessageLevel::Error).show();
                            }
                        }
                    } else {
                        rfd::MessageDialog::new().set_title("Napaka").set_description("Napaka pri branju pri dobivanju datotek").set_level(MessageLevel::Error).show();
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
                                        rfd::MessageDialog::new().set_title("Uspeh").set_description("Uspešno shranil šifrant").set_level(MessageLevel::Info).show();
                                    },
                                    Err(err) => {
                                        println!("Šifrant: {:?}", err);
                                        rfd::MessageDialog::new().set_title("Napaka").set_description("Napaka pri shranjevanju šifranta").set_level(MessageLevel::Error).show();
                                    }
                                }
                            },
                            Err(e) => {
                                println!("{}", e);
                                rfd::MessageDialog::new().set_title("Napaka").set_description("Napačno ime datoteke, ŠIFRANT.XLSX").set_level(MessageLevel::Error).show();
                            }
                        }
                    } else {
                        rfd::MessageDialog::new().set_title("Napaka").set_description("Napaka pri dobivanju datoteke").set_level(MessageLevel::Error).show();
                    }
                }

                if dobavitelji_button.clicked() {
                    let downloads_dir = dirs_next::download_dir().unwrap_or_else(|| std::path::PathBuf::from("C:\\"));
                    let file = rfd::FileDialog::new()
                        .set_title("Izberi 1 datoteko!")
                        .set_directory(downloads_dir)
                        .pick_file();
                    if file.is_some() {
                        let result = parse_dobavitelji_file(file.unwrap());
                        match result {
                            Ok(rows) => {
                                let db_result = self.db_manager.store_dobavitelji_to_db(rows);
                                match db_result {
                                    Ok(_) => {
                                        rfd::MessageDialog::new().set_title("Uspeh").set_description("Uspešno shranil dobavitelje").set_level(MessageLevel::Info).show();
                                    },
                                    Err(err) => {
                                        println!("Dobavitelji: {:?}", err);
                                        rfd::MessageDialog::new().set_title("Napaka").set_description("Napaka pri shranjevanju dobaviteljev").set_level(MessageLevel::Error).show();
                                    }
                                }
                            },
                            Err(e) => {
                                println!("{}", e);
                            }
                        }
                    } else {
                        rfd::MessageDialog::new().set_title("Napaka").set_description("Napaka pri dobivanju datoteke").set_level(MessageLevel::Error).show();
                    }
                }


                let delete = ui.button("Izbriši ne ročno vnesene tabele").on_hover_text("Izbriši pred posodabljanjem!");
                if delete.clicked() {
                    let res = self.db_manager.drop_all_tables();
                    match res {
                        Err(e) => {
                            if !e.to_string().eq("no such table: data (code 1)") {
                                println!("{}", e);
                                rfd::MessageDialog::new().set_title("Napaka").set_description("Napaka pri brisanju podatkov").set_level(MessageLevel::Error).show();
                            }
                        },
                        Ok(_) => {
                            rfd::MessageDialog::new().set_title("Uspeh").set_description("Uspešno zbrisal podatke").set_level(MessageLevel::Info).show();
                        }
                    }
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
    worksheet.write_string(0, 3, "Zaloga")?;
    worksheet.write_string(0, 4, "Poraba, Povprečna mesečna poraba za zadnje 3 mesece")?;
    worksheet.write_string(0, 5, "Poraba, Povprečna mesečna poraba za zadnjih 24 mesecev")?;
    worksheet.write_string(0, 6, "Odprto, Odprta naročila dobaviteljem")?;
    worksheet.write_string(0, 7, "Dobava, Predviden dobavni rok v mesecih")?;
    worksheet.write_string(0, 8, "Zaloga SAP, Trenutna zaloga v SAP-u, ki zadostuje za X mesecev na osnovi povprečne porabe preteklih 3 mesecev, če artikel nima 3M porabe računa na osnovi 24M porabe")?;
    worksheet.write_string(0, 9, "Sum Zaloga, Seštevek trenutne zaloge v SAP-u in odprtih naročil, ki zadostuje za X mesecev na osnovi povprečne porabe preteklih 3 mesecev, če artikel nima 3M porabe računa na osnovi 24M porabe")?;
    worksheet.write_string(0, 10, "Enota")?;
    worksheet.write_string(0, 11, "Dobavitelji")?;
    worksheet.write_string(0, 12, "Minimalna Zaloga")?;
    worksheet.write_string(0, 13, "Maximalna Zaloga")?;
    worksheet.write_string(0, 14, "Pakiranje")?;
    worksheet.write_string(0, 15, "Blagovna Skupina")?;
    worksheet.write_string(0, 16, "Opomba")?;


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

        match item.zaloga {
            Some(v) => worksheet.write_number(row, 3, v)?,
            None => worksheet.write_blank(row, 3, &Format::default())?,
        };

        match item.poraba_3m {
            Some(v) => worksheet.write_number(row, 4, v)?,
            None => worksheet.write_blank(row, 4, &Format::default())?,
        };

        match item.poraba_24m {
            Some(v) => worksheet.write_number(row, 5, v)?,
            None => worksheet.write_blank(row, 5, &Format::default())?,
        };

        match item.odprta_narocila {
            Some(v) => worksheet.write_number(row, 6, v)?,
            None => worksheet.write_blank(row, 6, &Format::default())?,
        };

        match item.dobavni_rok {
            Some(v) => worksheet.write_number(row, 7, v)?,
            None => worksheet.write_blank(row, 7, &Format::default())?,
        };

        match item.trenutna_zaloga_zadostuje_za_mesecev {
            Some(v) => worksheet.write_number(row, 8, v)?,
            None => worksheet.write_blank(row, 8, &Format::default())?,
        };

        match item.trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev {
            Some(v) => worksheet.write_number(row, 9, v)?,
            None => worksheet.write_blank(row, 9, &Format::default())?,
        };

        match &item.osnovna_merska_enota {
            Some(s) => worksheet.write_string(row, 10, s)?,
            None => worksheet.write_blank(row, 10, &Format::default())?,
        };

        match &item.dobavitelji {
            Some(s) => worksheet.write_string(row, 11, s)?,
            None => worksheet.write_blank(row, 11, &Format::default())?,
        };

        match item.minimalna_zaloga {
            Some(s) => worksheet.write_number(row, 12, s)?,
            None => worksheet.write_blank(row, 12, &Format::default())?,
        };

        match item.maximalna_zaloga {
            Some(s) => worksheet.write_number(row, 13, s)?,
            None => worksheet.write_blank(row, 13, &Format::default())?,
        };

        match &item.pakiranje {
            Some(s) => worksheet.write_string(row, 14, s)?,
            None => worksheet.write_blank(row, 14, &Format::default())?,
        };

        match &item.blagovna_skupina {
            Some(s) => worksheet.write_string(row, 15, s)?,
            None => worksheet.write_blank(row, 15, &Format::default())?,
        };

        match &item.opomba {
            Some(s) => worksheet.write_string(row, 16, s)?,
            None => worksheet.write_blank(row, 16, &Format::default())?,
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
