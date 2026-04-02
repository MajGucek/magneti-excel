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
}

impl App {
    pub fn new<'a>(cc: &'a eframe::CreationContext<'a>) -> Self {
        cc.egui_ctx.send_viewport_cmd(ViewportCommand::Maximized(true));

        Self {
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
                let import_buttom = ui.button("Add excel files");
                if import_buttom.clicked() {
                    let downloads_dir = dirs_next::download_dir().unwrap_or_else(|| std::path::PathBuf::from("C:\\"));
                    let files = rfd::FileDialog::new()
                        .set_title("Select files")
                        .set_directory(downloads_dir)
                        .pick_files();

                    if files.is_some() {
                        let result = handle_import_files(files.unwrap());
                    } else {
                        ui.colored_label(Color32::RED, "Error");
                    }
                }



            });
    }
}


#[derive(Default, Debug)]
struct RowData {
    material: i64, // zaloga
    naziv_materiala: String, // poraba
    zaloga: f64, // zaloga
    poraba: f64,
    odprta_narocila: f64,
}



fn handle_import_files(files: Vec<PathBuf>) -> Result<(), Box<dyn std::error::Error>> {
    if files.len() != 3 {
        MessageDialog::new()
            .set_title("Napaka, nepravilno število datotek != 3")
            .set_level(MessageLevel::Error)
            .show();
    }


    let poraba_file = files.iter().find(|&path_buf| {
        path_buf.file_name().unwrap() == "PORABA.XLSX"
    }).unwrap();

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
    }).unwrap();
    let mut workbook = open_workbook_auto(odprta_narocila_file)?;
    let range= workbook.worksheet_range(workbook.sheet_names().get(0).ok_or("Workbook has no sheets")?).unwrap();
    let mut dobava_map: HashMap<i64, f64> = HashMap::new();
    for row in range.rows().skip(4) {
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
    }).unwrap();
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
        let poraba = entry.1.abs() / 6.;
        let odprta_narocila = *dobava_map.get(&material).unwrap_or(&0.);


        row_data.push(RowData {
            material,
            naziv_materiala: opis_materiala,
            zaloga,
            poraba,
            odprta_narocila
        });
    }


    println!("obdelal!");
    row_data.iter().for_each(|row| {
        println!("{:?}", row);
    });


    Ok(())
}



fn main() {

    eframe::run_native(
        "Konstil Excel",
        NativeOptions::default(),
        Box::new(|cc| Ok(Box::new(App::new(cc))))
    ).unwrap();
}
