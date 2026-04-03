mod parse;
mod db;

use eframe::{Frame, NativeOptions};
use eframe::egui::*;
use egui_extras::{Column, TableBuilder};
use crate::db::{DBManager, ViewQuery};
use crate::parse::{parse_extra_config_files, parse_import_files, parse_sifrant_file};

struct App {
    db_manager: DBManager,
    retry_import: bool,
    successfully_parsed: Option<bool>,
    successfully_stored_data: Option<bool>,
    successfully_stored_sifrant: Option<bool>,

    row_data: Option<Vec<ViewQuery>>,
    successfully_loaded_query: Option<bool>,

    opomba_material: String,
    opomba_opomba: String,
    successfully_stored_opomba: Option<bool>,

    /* --Filters-- */
    filter_sifra_materiala: String,
    filter_naziv_materiala: String,
    filter_nabavnik: String,
    filter_zaloga_vecja: bool,
    filter_poraba_vecja: bool,
    filter_odprta_narocila: bool,
    filter_dobavni_rok: bool,
    filter_aktivni: bool,
}

impl App {
    pub fn new<'a>(cc: &'a eframe::CreationContext<'a>) -> Self {
        cc.egui_ctx.send_viewport_cmd(ViewportCommand::Maximized(true));

        let mut row_data = None;
        let db_manager = DBManager { db_name: "magneti_db.sqlite3".to_string() };
        
        let result = db_manager.get_data();
        let mut successfully_loaded_query: Option<bool> = None;

        match result {
            Err(err) => {
                println!("initial_load error: {:?}", err.to_string());
                successfully_loaded_query = Some(false);
            },
            Ok(rows) => {
                row_data = Some(rows);
                println!("row_data loaded: {}", row_data.as_ref().unwrap().len());
                successfully_loaded_query = Some(true);
            }
        }

        Self {
            db_manager,
            retry_import: false,
            successfully_parsed: None,
            successfully_stored_data: None,
            successfully_stored_sifrant: None,
            row_data,
            successfully_loaded_query,

            opomba_material: String::new(),
            opomba_opomba: String::new(),
            successfully_stored_opomba: None,

            filter_sifra_materiala: String::new(),
            filter_naziv_materiala: String::new(),
            filter_nabavnik: String::new(),
            filter_zaloga_vecja: false,
            filter_poraba_vecja: false,
            filter_odprta_narocila: false,
            filter_dobavni_rok: false,
            filter_aktivni: true,
        }
    }

}


impl App {
    fn apply_filters(&self, rows: &Vec<ViewQuery>) -> Vec<ViewQuery> {
        rows.iter()
            .filter(|&row| {
                format!("{}", row.material).contains(self.filter_sifra_materiala.as_str()) &&
                    row.naziv_materiala.as_ref().is_some_and(|a| format!("{}", a).contains(self.filter_naziv_materiala.as_str())) &&
                    row.nabavna_skupina.as_ref().is_some_and(|a| format!("{}", a).contains(self.filter_nabavnik.as_str())) &&
                    (!self.filter_aktivni || row.zaloga.is_some_and(|zal| zal > 0.0) || row.poraba.is_some_and(|por| por > 0.0)) &&
                    (!self.filter_zaloga_vecja || row.zaloga.is_some_and(|zal| zal > 0.)) &&
                    (!self.filter_poraba_vecja || row.poraba.is_some_and(|por| por > 0.)) &&
                    (!self.filter_odprta_narocila || row.odprta_narocila.is_some_and(|odp| odp > 0.)) &&
                    (!self.filter_dobavni_rok || row.dobavni_rok.is_some_and(|dob| dob > 0.))

        })
            .cloned()
            .collect()


    }


    pub fn render_table(&self, ui: &mut Ui) {
        let data = match &self.row_data {
            Some(d) => self.apply_filters(d),
            None => return,
        };

        let number_width = 120.;
        let string_width = 500.;

        ScrollArea::both().show(ui, |ui| {
            TableBuilder::new(ui)
                .striped(true)
                .cell_layout(Layout::left_to_right(Align::Center))
                .columns(Column::exact(number_width), 1) // Material
                .columns(Column::exact(string_width).at_least(string_width), 1) // Naziv materiala
                .columns(Column::exact(number_width), 2) // nabavna_skupina, mrp_karakteristika
                .columns(Column::exact(number_width), 4) // Zaloga, Poraba, Odprta narocila, Dobavni rok
                .columns(Column::exact(number_width).at_least(number_width), 2) // trenutni zalogi
                .columns(Column::remainder().at_least(string_width), 1) // Opomba
                .header(50.0, |mut header| {
                    header.col(|ui| {ui.heading("Material"); });
                    header.col(|ui| {ui.heading("Naziv"); });
                    header.col(|ui| {ui.heading("Nabavnik"); });
                    header.col(|ui| {ui.heading("MRP"); });
                    header.col(|ui| {ui.heading("Zaloga"); });
                    header.col(|ui| {ui.heading("Poraba").on_hover_text("Povprečna mesečna poraba za zadnjih 12 mesecev"); });
                    header.col(|ui| {ui.heading("Odprto").on_hover_text("Odprta naročila dobaviteljem"); });
                    header.col(|ui| {ui.heading("Dobava").on_hover_text("Predviden dobavni rok v mesecih"); });
                    header.col(|ui| {ui.heading("Zaloga SAP").on_hover_text("Trenutna zaloga v SAP-u"); });
                    header.col(|ui| {ui.heading("Zaloga SAP in odprto").on_hover_text("Seštevek trenutne zaloge v SAP-u in odprtih naročil"); });
                    header.col(|ui| {ui.heading("Opomba"); });
                })
                .body(|mut body| {
                    for row in data.iter().take(50) {
                        body.row(25.0, |mut row_ui| {
                            row_ui.col(|ui| { ui.label(row.material.to_string()); });
                            row_ui.col(|ui| { ui.label(row.naziv_materiala.clone().unwrap_or_else(|| "".to_string())); });
                            row_ui.col(|ui| { ui.label(row.nabavna_skupina.clone().unwrap_or_else(|| "".to_string())); });
                            row_ui.col(|ui| { ui.label(row.mrp_karakteristika.clone().unwrap_or_else(|| "".to_string())); });
                            row_ui.col(|ui| { ui.label(row.zaloga.map_or("".to_string(), |v| format_number_custom(v))); });
                            row_ui.col(|ui| { ui.label(row.poraba.map_or("".to_string(), |v| format_number_custom(v))); });
                            row_ui.col(|ui| { ui.label(row.odprta_narocila.map_or("".to_string(), |v| format_number_custom(v))); });
                            row_ui.col(|ui| { ui.label(row.dobavni_rok.map_or("".to_string(), |v| format_number_custom(v))); });
                            row_ui.col(|ui| { ui.label(row.trenutna_zaloga_zadostuje_za_mesecev.map_or("".to_string(), |v| format_number_custom(v))); });
                            row_ui.col(|ui| { ui.label(row.trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev.map_or("".to_string(), |v| format_number_custom(v))); });
                            row_ui.col(|ui| { ui.label(row.opomba.clone().unwrap_or_else(|| "".to_string())); });

                        });
                    }
                });
        });
    }
}

fn format_number_custom(value: f64) -> String {
    let s = format!("{:.2}", value);
    let mut parts = s.split('.').collect::<Vec<_>>();
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
    fn update(&mut self, ctx: &Context, _frame: &mut Frame) {
        //ctx.request_repaint();

        CentralPanel::default().show(ctx, |ui| {
            ui.horizontal(|ui| {
                if let Some(success) = self.successfully_loaded_query {
                    let color = if success { Color32::GREEN } else { Color32::RED };
                    ui.colored_label(color, "■");
                }
                if ui.button("Load").clicked() {
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


            });

            ui.add_space(4.0);

            ScrollArea::vertical().show(ui, |ui| {
                self.render_table(ui);
            });
        });


        Window::new("Filtri")
            .resizable(false)
            .collapsible(true)
            .default_open(true)
            .default_size(vec2(600., 400.))
            .show(ctx, |ui| {
                ui.vertical_centered(|ui| {
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


                    ui.checkbox(&mut self.filter_aktivni, "Pokaži le aktivne");
                    ui.checkbox(&mut self.filter_zaloga_vecja, "zaloga večja od 0");
                    ui.checkbox(&mut self.filter_poraba_vecja, "poraba večja od 0");
                    ui.checkbox(&mut self.filter_odprta_narocila, "odprta naročila večja od 0");
                    ui.checkbox(&mut self.filter_dobavni_rok, "dobavni rok večji kot 0");


                });

            });

        Window::new("Import")
            .resizable(false)
            .collapsible(true)
            .default_open(true)
            .default_size(vec2(300., 200.))
            .show(ctx, |ui| {
                let import_button = ui.button("Add excel files");
                let sifrant_button = ui.button("Add šifrant");
                let extra_config_button = ui.button("Add extra configuration");
                let mut file_input_error: Option<Box<dyn std::error::Error>> = None;
                if import_button.clicked() {
                    let downloads_dir = dirs_next::download_dir().unwrap_or_else(|| std::path::PathBuf::from("C:\\"));
                    let files = rfd::FileDialog::new()
                        .set_title("Select files")
                        .set_directory(downloads_dir)
                        .pick_files();
                    if files.is_some() {
                        let result = parse_import_files(files.unwrap());
                        match result {
                            Ok(row_data) => {
                                self.retry_import = false;
                                self.successfully_parsed = Some(true);
                                let db_result = self.db_manager.store_to_db(row_data);
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
                        .set_title("Select file!")
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


                if extra_config_button.clicked() {
                    let downloads_dir = dirs_next::download_dir().unwrap_or_else(|| std::path::PathBuf::from("C:\\"));
                    let file = rfd::FileDialog::new()
                        .set_title("Select file!")
                        .set_directory(downloads_dir)
                        .pick_file();
                    if file.is_some() {
                        let result = parse_extra_config_files(file.unwrap());
                        match result {
                            Ok(extra_config_rows) => {
                                let db_result = self.db_manager.store_extra_config_to_db(extra_config_rows);
                                match db_result {
                                    Ok(_) => {
                                        self.successfully_stored_data = Some(true);
                                    },
                                    Err(err) => {
                                        println!("config: {:?}", err);
                                        self.successfully_stored_data = Some(false);
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
                        ui.colored_label(Color32::RED, "File input Error");
                    } else {
                        ui.colored_label(Color32::RED, file_input_error.unwrap().to_string());
                    }
                }
                if self.successfully_parsed.is_some() {
                    if self.successfully_parsed.unwrap() {
                        ui.colored_label(Color32::GREEN, "Successful parse!");
                    } else {
                        ui.colored_label(Color32::ORANGE, "Failed to parse, errors in excel!");
                    }
                }
                if self.successfully_stored_data.is_some() {
                    if self.successfully_stored_data.unwrap() {
                        ui.colored_label(Color32::GREEN, "Successful store to database!");
                    } else {
                        ui.colored_label(Color32::ORANGE, "Failed to store to DB!");
                    }
                }

                if self.successfully_stored_sifrant.is_some() {
                    if self.successfully_stored_sifrant.unwrap() {
                        ui.colored_label(Color32::GREEN, "Successful store to database!");
                    } else {
                        ui.colored_label(Color32::ORANGE, "Failed to store to DB!");
                    }
                }


            });


        Window::new("Opomba")
            .resizable(false)
            .collapsible(true)
            .default_open(true)
            .default_size(vec2(300., 200.))
            .show(ctx, |ui| {
                ui.vertical_centered(|ui| {
                    ui.horizontal(|ui| {
                        ui.label("Material: ");
                        ui.text_edit_singleline(&mut self.opomba_material);
                    });
                    ui.horizontal(|ui| {
                        ui.label("Opomba: ");
                        ui.text_edit_singleline(&mut self.opomba_opomba);
                    });
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
                    }
                    if self.successfully_stored_opomba.is_some() {
                        if self.successfully_stored_opomba.unwrap() {
                            ui.colored_label(Color32::GREEN, "Success!");
                        } else {
                            ui.colored_label(Color32::RED, "Failed to store opomba!");
                        }
                    }
                    
                });

            });
    }
}


fn main() {

    eframe::run_native(
        "Magneti Excel",
        NativeOptions::default(),
        Box::new(|cc| Ok(Box::new(App::new(cc))))
    ).unwrap();
}
