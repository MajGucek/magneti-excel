use std::collections::HashMap;
use std::io::stderr;
use std::path::PathBuf;
use calamine::{open_workbook_auto, Reader};
use eframe::{Frame, NativeOptions};
use eframe::HardwareAcceleration::Preferred;
use eframe::egui::*;
use egui_extras::{Column, TableBuilder};
use calamine::DataType;
use rfd::{MessageDialog, MessageLevel};

struct App {
    retry_import: bool,
    successfully_parsed: Option<bool>,
    successfully_stored_data: Option<bool>,
}

impl App {
    pub fn new<'a>(cc: &'a eframe::CreationContext<'a>) -> Self {
        cc.egui_ctx.send_viewport_cmd(ViewportCommand::Maximized(true));

        Self {
            retry_import: false,
            successfully_parsed: None,
            successfully_stored_data: None,
        }
    }

    pub fn update_state(&mut self) {

    }
}

impl eframe::App for App {
    fn update(&mut self, ctx: &Context, _frame: &mut Frame) {
        ctx.request_repaint();

        CentralPanel::default().show(ctx, |ui| {
            let total_height = ui.available_height();
            let top_height = total_height * 0.15;
            let bottom_height = total_height * 0.85;

            // Search bar
            ui.allocate_ui(vec2(ui.available_width(), top_height), |ui| {
                ui.horizontal(|ui| {
                    ui.label("Search:");
                    ui.text_edit_singleline(&mut String::new());
                });
            });

            // Table
            ui.allocate_ui(vec2(ui.available_width(), bottom_height), |ui| {
                ScrollArea::vertical().show(ui, |ui| {
                    TableBuilder::new(ui)
                        .striped(true)
                        .cell_layout(Layout::left_to_right(Align::Center))
                        .columns(Column::remainder(), 2)
                        .header(20.0, |mut header| {
                            header.col(|ui| { ui.heading("ID"); });
                            header.col(|ui| { ui.heading("Material"); });
                        })
                        .body(|mut body| {
                            for row in 0..20 {
                                body.row(18.0, |mut row_ui| {
                                    row_ui.col(|ui| { ui.label("id"); });
                                    row_ui.col(|ui| { ui.label("material"); });
                                });
                            }
                        });
                });
            });
        });

        Window::new("Import")
            .resizable(false)
            .collapsible(true)
            .default_open(true)
            .default_size(vec2(300., 200.))
            .show(ctx, |ui| {
                let import_button = ui.button("Add excel files");
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
                                let db_result = store_to_db(row_data);
                                match db_result {
                                    Ok(_) => {
                                        self.successfully_stored_data = Some(true);
                                    },
                                    Err(err) => {
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



                if extra_config_button.clicked() {
                    let downloads_dir = dirs_next::download_dir().unwrap_or_else(|| std::path::PathBuf::from("C:\\"));
                    let file = rfd::FileDialog::new()
                        .set_title("Select file!")
                        .set_directory(downloads_dir)
                        .pick_file();
                    if file.is_some() {
                        let result = parse_extra_config_files(file.unwrap());
                        match result {
                            Ok(extra_config_row) => {
                                let db_result = store_extra_config_to_db(extra_config_row);
                                match db_result {
                                    Ok(_) => {
                                        self.successfully_stored_data = Some(true);
                                    },
                                    Err(err) => {
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


            });
    }
}

#[derive(Default)]
struct ExtraConfigRow {
    nabavnik: String,
    min_kolicina: f64,
    pakiranje: f64,
    dobavni_rok: f64,
}

fn parse_extra_config_files(path: PathBuf) -> Result<ExtraConfigRow, Box<dyn std::error::Error>> {
    let mut workbook = open_workbook_auto(path)?;
    let range= workbook.worksheet_range(workbook.sheet_names().get(0).ok_or("Workbook has no sheets")?)?;
    let mut extra_config = ExtraConfigRow::default();
    for row in range.rows().skip(1) {
        let nabavnik = row.get(0)
            .and_then(DataType::get_string)
            .unwrap_or("");
        let min_kolicina = row.get(1)
            .and_then(DataType::get_float)
            .unwrap_or(0.);
        let pakiranje = row.get(2)
            .and_then(DataType::get_float)
            .unwrap_or(0.);
        let dobavni_rok = row.get(3)
            .and_then(DataType::get_float)
            .unwrap_or(0.);
        extra_config = ExtraConfigRow {
            nabavnik: nabavnik.to_string(),
            min_kolicina,
            pakiranje,
            dobavni_rok
        };
    }

    Ok(extra_config)
}


fn store_extra_config_to_db(extra_config_row: ExtraConfigRow) -> Result<(), Box<dyn std::error::Error>> {
    let connection = sqlite::open("magneti_db.sqlite3")?;
    connection.execute("
        CREATE TABLE IF NOT EXISTS config (
            id INTEGER PRIMARY KEY,
            nabavnik TEXT NOT NULL,
            min_kolicina REAL NOT NULL,
            pakiranje REAL NOT NULL,
            dobavni_rok REAL NOT NULL
        );
    ")?;


    let mut statement = connection.prepare("
        INSERT INTO config (nabavnik, min_kolicina, pakiranje, dobavni_rok) VALUES (?, ?, ?, ?)
    ")?;

    statement.bind(&[(1, extra_config_row.nabavnik.as_str())][..])?;
    statement.bind((2, extra_config_row.min_kolicina))?;
    statement.bind((3, extra_config_row.pakiranje))?;
    statement.bind((4, extra_config_row.dobavni_rok))?;
    statement.next()?;
    statement.reset()?;

    Ok(())
}

fn store_to_db(row_data: Vec<RowData>) -> Result<(), Box<dyn std::error::Error>> {
    let connection = sqlite::open("magneti_db.sqlite3")?;
    connection.execute("
        CREATE TABLE IF NOT EXISTS data (
            id INTEGER PRIMARY KEY,
            material INTEGER NOT NULL,
            naziv_materiala TEXT,
            zaloga REAL,
            poraba REAL,
            odprta_narocila REAL,
            nabavnik TEXT,
            min_kolicina REAL,
            pakiranje REAL,
            dobavni_rok REAL
        );
    ")?;


    let mut statement = connection.prepare("
        INSERT INTO data (material, naziv_materiala, zaloga, poraba, odprta_narocila) VALUES (?, ?, ?, ?, ?)
    ")?;
    for row in row_data {
        statement.bind((1, row.material))?;
        statement.bind(&[(2, row.naziv_materiala.as_str())][..])?;
        statement.bind((3, row.zaloga))?;
        statement.bind((4, row.poraba))?;
        statement.bind((5, row.odprta_narocila))?;
        statement.next()?;
        statement.reset()?;
    }

    Ok(())
}


#[derive(Default, Debug)]
struct RowData {
    material: i64, // zaloga
    naziv_materiala: String, // poraba
    zaloga: f64, // zaloga
    poraba: f64,
    odprta_narocila: f64,
}



fn parse_import_files(files: Vec<PathBuf>) -> Result<Vec<RowData>, Box<dyn std::error::Error>> {
    if files.len() != 3 {
        MessageDialog::new()
            .set_title("Napaka, nepravilno število datotek != 3")
            .set_level(MessageLevel::Error)
            .show();
    }


    let poraba_file = files.iter().find(|&path_buf| {
        path_buf.file_name().unwrap() == "PORABA.XLSX"
    }).ok_or("File PORABA.XLSX not found")?;

    let mut workbook = open_workbook_auto(poraba_file)?;
    let range= workbook.worksheet_range(workbook.sheet_names().get(0).ok_or("Workbook has no sheets")?).unwrap();
    let mut material_map: HashMap<i64, (String, f64)> = HashMap::new();
    for row in range.rows().skip(1) {
        let material = row.get(1)
            .and_then(DataType::get_string)
            .map(|f| f.parse::<i64>().unwrap_or(0))
            .unwrap_or(0);
        if material == 0 { continue; }

        let opis_materiala = row.get(2)
            .and_then(DataType::get_string)
            .unwrap_or("");

        let klc_v_em_vnosa = row.get(9)
            .and_then(DataType::get_float)
            .map(|f| f)
            .unwrap_or(0.);

        let entry = material_map
            .entry(material)
            .or_insert((opis_materiala.to_string(), 0.0));

        (*entry).1 += klc_v_em_vnosa;
    }

    let odprta_narocila_file = files.iter().find(|&path_buf| {
        path_buf.file_name().unwrap() == "ODPRTA NAROČILA.XLSX"
    }).ok_or("File PORABA:XLSX not found")?;
    let mut workbook = open_workbook_auto(odprta_narocila_file)?;
    let range= workbook.worksheet_range(workbook.sheet_names().get(0).ok_or("Workbook has no sheets")?).unwrap();
    let mut dobava_map: HashMap<i64, f64> = HashMap::new();
    for row in range.rows().skip(1) {
        let material = row.get(0)
            .and_then(DataType::get_string)
            .map(|f| f.parse::<i64>().unwrap_or(0))
            .unwrap_or(0);
        if material == 0 { continue; }



        let se_za_dobavo = row.get(23)
            .and_then(DataType::get_float)
            .map(|f| f)
            .unwrap_or(0.);

        let entry = dobava_map
            .entry(material)
            .or_insert(0.);

        *entry += se_za_dobavo;
    }


    let zaloga_file = files.iter().find(|&path_buf| {
        path_buf.file_name().unwrap() == "ZALOGA.XLSX"
    }).ok_or("File PORABA:XLSX not found")?;
    let mut workbook = open_workbook_auto(zaloga_file)?;
    let range= workbook.worksheet_range(workbook.sheet_names().get(0).ok_or("Workbook has no sheets")?).unwrap();
    let mut row_data: Vec<RowData> = Vec::with_capacity(100);
    for row in range.rows().skip(1) {
        let material = row.get(1)
            .and_then(DataType::get_string)
            .map(|f| f.parse::<i64>().unwrap_or(0))
            .unwrap_or(0);
        if material == 0 { continue; }

        let zaloga = row.get(7)
            .and_then(DataType::get_float)
            .map(|f| f)
            .unwrap_or(0.);

        let emtpy_entry = ("".to_string(), 0.);
        let entry = material_map.get(&material).unwrap_or(&emtpy_entry);
        let opis_materiala = entry.0.clone();
        let poraba = entry.1.abs() / 12.;
        let odprta_narocila = *dobava_map.get(&material).unwrap_or(&0.);


        row_data.push(RowData {
            material,
            naziv_materiala: opis_materiala,
            zaloga,
            poraba,
            odprta_narocila
        });
    }


    Ok(row_data)
}



fn main() {

    eframe::run_native(
        "Konstil Excel",
        NativeOptions::default(),
        Box::new(|cc| Ok(Box::new(App::new(cc))))
    ).unwrap();
}
