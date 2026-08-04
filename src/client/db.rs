use std::time::{Duration, Instant};
use eframe::egui::{Color32, CornerRadius, CursorIcon, RichText};
use rfd::{MessageButtons, MessageDialog, MessageDialogResult, MessageLevel};
pub(crate) use magneti_excel::{NabavaQuery, PorabaQuery, SortState, ViewQuery, ViewQueryFields};
use magneti_excel::{format_elapsed_time, HandInput};
use crate::graph::PorabaNabavaRows;
use crate::{format_nabavnik, format_number_custom, parse_string_to_optional_f64, Rows, INDIGO, RED, TEAL};

pub struct DBManager {
    url: String,
    last_query: Option<Instant>,
}

impl DBManager {
    pub fn get_last_query_time(&self) -> String {
        match self.last_query {
            Some(i) => {
                format_elapsed_time(i)
            },
            None => {
                String::new()
            }
        }
    }

    pub fn create(url: &str) -> Self {
        Self {
            url: url.to_string(),
            last_query: None,
        }
    }
    pub fn update_url(&mut self, url: &str) {
        log::info!("Updating server IP");
        self.url = format!("http://{}:8080", url);
    }
    pub fn get_data(&mut self, sort: &SortState) -> Result<Vec<ViewQuery>, Box<dyn std::error::Error>> {
        log::info!("sending request to server");
        let sort_json = serde_json::to_string(sort)?;
        let req_url = format!("{}/data/0/{}/{}", self.url, usize::MAX,  urlencoding::encode(&sort_json));

        self.last_query = Some(Instant::now());
        let data: Vec<ViewQuery> = ureq::get(&req_url).call()?.body_mut().with_config().limit(u64::MAX).read_json()?;
        log::info!("Exiting get_data()");
        Ok(data)
    }

    pub fn get_poraba(&mut self, material: i64) -> Result<Vec<PorabaQuery>, Box<dyn std::error::Error>> {
        let req_url = format!("{}/poraba/{}", self.url, material);

        self.last_query = Some(Instant::now());
        let data: Vec<PorabaQuery> = ureq::get(&req_url).call()?.body_mut().read_json()?;
        Ok(data)
    }
    pub fn get_nabava(&mut self, material: i64) -> Result<Vec<NabavaQuery>, Box<dyn std::error::Error>> {
        let req_url = format!("{}/nabava/{}", self.url, material);

        self.last_query = Some(Instant::now());
        let data: Vec<NabavaQuery> = ureq::get(&req_url).call()?.body_mut().read_json()?;
        Ok(data)
    }

    fn store_url(&self) -> String {
        format!("{}/upload", self.url)
    }

    pub fn store_pakiranje_to_server(&self, pakiranje: (i64, String)) -> Result<(), Box<dyn std::error::Error>> {
        let hand_input = HandInput::Pakiranje(pakiranje.0, pakiranje.1);
        let json_value = serde_json::to_value(&hand_input)?;
        let response = ureq::post(&self.store_url()).send_json(json_value)?;
        log::info!("{:?}", response);
        Ok(())
    }

    pub fn store_opomba_to_server(&self, opomba: (i64, String)) -> Result<(), Box<dyn std::error::Error>> {
        let hand_input = HandInput::Opomba(opomba.0, opomba.1);
        let json_value = serde_json::to_value(&hand_input)?;
        let response = ureq::post(&self.store_url()).send_json(json_value)?;
        log::info!("{:?}", response);
        Ok(())
    }


    pub fn store_blagovna_skupina_to_server(&self, blagovna_skupina: (i64, String)) -> Result<(), Box<dyn std::error::Error>> {
        let hand_input = HandInput::BlagovnaSkupina(blagovna_skupina.0, blagovna_skupina.1);
        let json_value = serde_json::to_value(&hand_input)?;
        let response = ureq::post(&self.store_url()).send_json(json_value)?;
        log::info!("{:?}", response);
        Ok(())
    }



    pub fn store_min_zaloga_to_server(&self, min_zaloga_row: (i64, Option<f64>)) -> Result<(), Box<dyn std::error::Error>> {
        let hand_input = HandInput::MinZaloga(min_zaloga_row.0, min_zaloga_row.1);
        let json_value = serde_json::to_value(&hand_input)?;
        let response = ureq::post(&self.store_url()).send_json(json_value)?;
        log::info!("{:?}", response);
        Ok(())
    }


    pub fn store_max_zaloga_to_server(&self, max_zaloga_row: (i64, Option<f64>)) -> Result<(), Box<dyn std::error::Error>> {
        let hand_input = HandInput::MaxZaloga(max_zaloga_row.0, max_zaloga_row.1);
        let json_value = serde_json::to_value(&hand_input)?;
        let response = ureq::post(&self.store_url()).send_json(json_value)?;
        log::info!("{:?}", response);
        Ok(())
    }



    pub fn store_dobavni_rok_to_server(&self, dobavni_rok_row: (i64, Option<f64>)) -> Result<(), Box<dyn std::error::Error>> {
        let hand_input = HandInput::DobavniRok(dobavni_rok_row.0, dobavni_rok_row.1);
        let json_value = serde_json::to_value(&hand_input)?;
        let response = ureq::post(&self.store_url()).send_json(json_value)?;
        log::info!("{:?}", response);
        Ok(())
    }



}


pub fn construct_headers(field: ViewQueryFields, header: &mut egui_extras::TableRow, sort: &mut ViewQueryFields) {
    match field {
        ViewQueryFields::Material => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::Material, "Material"); });},
        ViewQueryFields::NazivMateriala => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::NazivMateriala, "Naziv"); });},
        ViewQueryFields::OsnovnaMerskaEnota => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::OsnovnaMerskaEnota, "Enota"); });},
        ViewQueryFields::NabavnaSkupina => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::NabavnaSkupina, "Nabavnik").on_hover_text("002 Neli\n008 Viktoriia\n010 Boštjan"); });},
        ViewQueryFields::MRP => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::MRP, "MRP"); });},
        ViewQueryFields::Zaloga => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::Zaloga, "Zaloga Sum").on_hover_text("Trenutna zaloga v SAP-u"); });},
        ViewQueryFields::Poraba3M => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::Poraba3M, "Poraba 3M").on_hover_text("Povprečna mesečna poraba za zadnje 3 mesece"); });},
        ViewQueryFields::Poraba24M => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::Poraba24M, "Poraba 24M").on_hover_text("Povprečna mesečna poraba za zadnjih 24 mesecev"); });},
        ViewQueryFields::OdprtaNarocila => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::OdprtaNarocila, "Odprto").on_hover_text("Odprta naročila dobaviteljem"); });},
        ViewQueryFields::DobavniRok => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::DobavniRok, "Dobava").on_hover_text("Predviden dobavni rok v mesecih"); });},
        ViewQueryFields::TrenutnaZalogaZadostujeZaMesecev => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::TrenutnaZalogaZadostujeZaMesecev, "Zaloga SAP").on_hover_text("Trenutna zaloga v SAP-u, ki zadostuje za X mesecev na osnovi povprečne porabe preteklih 3 mesecev, če artikel nima 3M porabe računa na osnovi 24M porabe"); });},
        ViewQueryFields::TrenutnaZalogaInOdprtaNarocilaZadostujeZaMesecev => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::TrenutnaZalogaInOdprtaNarocilaZadostujeZaMesecev, "Zaloga Sum SAP").on_hover_text("Seštevek trenutne zaloge v SAP-u in odprtih naročil, ki zadostuje za X mesecev na osnovi povprečne porabe preteklih 3 mesecev, če artikel nima 3M porabe računa na osnovi 24M porabe"); });},
        ViewQueryFields::Dobavitelji => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::Dobavitelji, "Dobavitelji"); });},
        ViewQueryFields::Cena => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::Cena, "Cena"); });},
        ViewQueryFields::Valuta => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::Valuta, "Valuta"); });},
        ViewQueryFields::RazpolozljivaZaloga => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::RazpolozljivaZaloga, "Zaloga 100"); });},
        ViewQueryFields::MinimalnaZaloga => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::MinimalnaZaloga, "Min zaloga"); });},
        ViewQueryFields::MaximalnaZaloga => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::MaximalnaZaloga, "Max zaloga"); });},
        ViewQueryFields::BlagovnaSkupina => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::BlagovnaSkupina, "Blagovna skupina"); });},
        ViewQueryFields::Pakiranje => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::Pakiranje, "Pakiranje"); });},
        ViewQueryFields::Lokacija => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::Lokacija, "Lokacija"); });},
        ViewQueryFields::Opomba => {header.col(|ui| {ui.radio_value(sort, ViewQueryFields::Opomba, "Opomba"); });},
    }
}


pub fn construct_body(field: ViewQueryFields,
                      table_row: &mut egui_extras::TableRow,
                      index: usize,
                      row: &ViewQuery,
                      mut row_color: Color32,
                      poraba_nabava_data: &mut PorabaNabavaRows,
                      mut db_manager: &mut DBManager,
                      sort_state: &SortState,
                      row_data: &mut Rows,
                      editing_dobavni_rok_row: &mut Option<usize>,
                      edit_dobavni_rok_input: &mut String,
                      editing_min_zaloga_row: &mut Option<usize>,
                      edit_min_zaloga_row_input: &mut String,
                      editing_max_zaloga_row: &mut Option<usize>,
                      edit_max_zaloga_row_input: &mut String,
                      editing_pakiranje_row: &mut Option<usize>,
                      edit_pakiranje_input: &mut String,
                      editing_blagovna_skupina_row: &mut Option<usize>,
                      edit_blagovna_skupina_input: &mut String,
                      editing_opomba_row: &mut Option<usize>,
                      edit_opomba_input: &mut String,
) {
    match field {
        ViewQueryFields::Material => {
            table_row.col(|ui| {
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                if ui.label(RichText::new(row.material.to_string()).underline().background_color(Color32::TRANSPARENT))
                    .on_hover_cursor(CursorIcon::PointingHand)
                    .clicked() {

                    poraba_nabava_data.query(row.material, row.naziv_materiala.as_ref().unwrap_or(&"".to_string()).as_str(), row.zaloga.unwrap_or(0.),  &mut db_manager);
                }
            });
        }
        ViewQueryFields::NazivMateriala => {
            table_row.col(|ui| {
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                ui.label(row.naziv_materiala.clone().unwrap_or_else(|| "".to_string()));
            });
        }
        ViewQueryFields::OsnovnaMerskaEnota => {
            table_row.col(|ui| {
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                let t = row.osnovna_merska_enota.clone().unwrap_or_else(|| "".to_string());
                ui.label(&t);
            });
        }
        ViewQueryFields::NabavnaSkupina => {
            table_row.col(|ui| {
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                let nabavna_skupina = row.nabavna_skupina.clone().unwrap_or_else(|| "".to_string());

                ui.label(format_nabavnik(nabavna_skupina.as_str()).unwrap_or(nabavna_skupina.as_str()));
            });
        }
        ViewQueryFields::MRP => {
            table_row.col(|ui| {
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                let t = row.mrp_karakteristika.clone().unwrap_or_else(|| "".to_string());
                ui.label(&t);
            });
        }
        ViewQueryFields::Zaloga => {
            table_row.col(|ui| {
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                ui.label(row.zaloga.map_or("".to_string(), |v| format_number_custom(v, 1)));
            });
        }
        ViewQueryFields::Poraba3M => {
            table_row.col(|ui| {
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                let poraba_3m = row.poraba_3m.map_or("".to_string(), |v| format_number_custom(v, 1));
                let poraba_24m = row.poraba_24m.map_or("".to_string(), |v| format_number_custom(v, 1));

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
        }
        ViewQueryFields::Poraba24M => {
            table_row.col(|ui| {
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                ui.label(row.poraba_24m.map_or("".to_string(), |v| format_number_custom(v, 1)));
            });
        }
        ViewQueryFields::OdprtaNarocila => {
            table_row.col(|ui| {
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                ui.label(row.odprta_narocila.map_or("".to_string(), |v| format_number_custom(v, 0)));
            });
        }
        ViewQueryFields::DobavniRok => {
            table_row.col(|ui| {
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);

                if *editing_dobavni_rok_row == Some(index) {
                    let response = ui.text_edit_singleline(edit_dobavni_rok_input);
                    if response.lost_focus() {
                        *editing_dobavni_rok_row = None;

                        let os_resp = MessageDialog::new()
                            .set_title("Potrdi vnos")
                            .set_level(MessageLevel::Info)
                            .set_buttons(MessageButtons::OkCancel)
                            .show();

                        match os_resp {
                            MessageDialogResult::Ok => {
                                let _ = db_manager.store_dobavni_rok_to_server((
                                    row.material,
                                    parse_string_to_optional_f64(edit_dobavni_rok_input.as_str()),
                                ));
                                row_data.query(&mut db_manager, sort_state);
                            },
                            _ => {}
                        }
                    }

                } else {
                    let label_text = row.dobavni_rok.map_or(" ".repeat(18), |v| format_number_custom(v, 1));
                    let resp = ui.label(label_text).on_hover_cursor(CursorIcon::Help);
                    if resp.double_clicked() {
                        *editing_dobavni_rok_row = Some(index);
                        *edit_dobavni_rok_input = row.dobavni_rok.map_or("".to_string(), |v| format!("{}", v));
                    }

                }

            });
        }
        ViewQueryFields::TrenutnaZalogaZadostujeZaMesecev => {
            table_row.col(|ui| {
                let old = row_color;
                if row.odprta_narocila.is_some_and(|o| o != 0.) &&
                    row.trenutna_zaloga_zadostuje_za_mesecev.is_some_and(|val| val < row.dobavni_rok.unwrap_or(0.)) {
                    row_color = RED;
                }
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                ui.label(row.trenutna_zaloga_zadostuje_za_mesecev.map_or("".to_string(), |v| format_number_custom(v, 1)));
                row_color = old;
            });
        }
        ViewQueryFields::TrenutnaZalogaInOdprtaNarocilaZadostujeZaMesecev => {
            table_row.col(|ui| {
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                ui.label(row.trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev.map_or("".to_string(), |v| format_number_custom(v, 1)));
            });
        }
        ViewQueryFields::Dobavitelji => {
            table_row.col(|ui| {
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                let t = row.dobavitelji.clone().unwrap_or_else(|| "".to_string());
                ui.label(&t);
            });
        }
        ViewQueryFields::Cena => {
            table_row.col(|ui| {
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                ui.label(row.cena.map_or("".to_string(), |v| format_number_custom(v, 2)));
            });

        }
        ViewQueryFields::Valuta => {
            table_row.col(|ui| {
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                let t = row.valuta.clone().unwrap_or_else(|| "".to_string());
                ui.label(&t);
            });
        }
        ViewQueryFields::RazpolozljivaZaloga => {
            table_row.col(|ui| {
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                ui.label(row.razpolozljiva_zaloga.map_or("".to_string(), |v| format_number_custom(v, 1)));
            });
        }
        ViewQueryFields::Lokacija => {
            table_row.col(|ui| {
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                let t = row.lokacija.clone().unwrap_or_else(|| "".to_string());
                ui.label(&t);
            });
        }
        ViewQueryFields::MinimalnaZaloga => {
            table_row.col(|ui| {
                let old = row_color;
                if row.minimalna_zaloga.is_some_and(|val| val > (row.zaloga.unwrap_or(0.) + row.odprta_narocila.unwrap_or(0.)))  {
                    // teal
                    row_color = TEAL;
                }
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);

                if *editing_min_zaloga_row == Some(index) {
                    let response = ui.text_edit_singleline(edit_min_zaloga_row_input);
                    if response.lost_focus() {
                        *editing_min_zaloga_row = None;

                        let os_resp = MessageDialog::new()
                            .set_title("Potrdi vnos")
                            .set_level(MessageLevel::Info)
                            .set_buttons(MessageButtons::OkCancel)
                            .show();

                        match os_resp {
                            MessageDialogResult::Ok => {
                                let _ = db_manager.store_min_zaloga_to_server((
                                                                        row.material,
                                                                        edit_min_zaloga_row_input.clone().parse::<f64>().ok()),
                                );
                                row_data.query(db_manager, sort_state);
                            },
                            _ => {}
                        }
                    }

                } else {
                    let label_text = row.minimalna_zaloga.map_or(" ".repeat(28), |v| format_number_custom(v, 0));
                    let resp = ui.label(label_text).on_hover_cursor(CursorIcon::Help);
                    if resp.double_clicked() {
                        *editing_min_zaloga_row = Some(index);
                        *edit_min_zaloga_row_input = row.minimalna_zaloga.map_or("".to_string(), |v| format_number_custom(v, 0));
                    }

                }


                row_color = old;
            });
        }
        ViewQueryFields::MaximalnaZaloga => {
            table_row.col(|ui| {
                let old = row_color;
                if row.maximalna_zaloga.is_some_and(|val| val < row.zaloga.unwrap_or(0.)) {
                    // indigo
                    row_color = INDIGO;
                }

                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);

                if *editing_max_zaloga_row == Some(index) {
                    let response = ui.text_edit_singleline(edit_max_zaloga_row_input);
                    if response.lost_focus() {
                        *editing_max_zaloga_row = None;

                        let os_resp = MessageDialog::new()
                            .set_title("Potrdi vnos")
                            .set_level(MessageLevel::Info)
                            .set_buttons(MessageButtons::OkCancel)
                            .show();

                        match os_resp {
                            MessageDialogResult::Ok => {
                                let _ = db_manager.store_max_zaloga_to_server((
                                                                        row.material,
                                                                        edit_max_zaloga_row_input.clone().parse::<f64>().ok()),
                                );
                                row_data.query(db_manager, sort_state);
                            },
                            _ => {}
                        }
                    }

                } else {
                    let label_text = row.maximalna_zaloga.map_or(" ".repeat(28), |v| format_number_custom(v, 0));
                    let resp = ui.label(label_text).on_hover_cursor(CursorIcon::Help);
                    if resp.double_clicked() {
                        *editing_max_zaloga_row = Some(index);
                        *edit_max_zaloga_row_input = row.maximalna_zaloga.map_or("".to_string(), |v| format_number_custom(v, 0));
                    }

                }


                row_color = old;
            });
        }
        ViewQueryFields::BlagovnaSkupina => {
            table_row.col(|ui| {
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);

                if *editing_blagovna_skupina_row == Some(index) {
                    let response = ui.text_edit_singleline(edit_blagovna_skupina_input);
                    if response.lost_focus() {
                        *editing_blagovna_skupina_row = None;

                        let os_resp = MessageDialog::new()
                            .set_title("Potrdi vnos")
                            .set_level(MessageLevel::Info)
                            .set_buttons(MessageButtons::OkCancel)
                            .show();

                        match os_resp {
                            MessageDialogResult::Ok => {
                                let _ = db_manager.store_blagovna_skupina_to_server((
                                                                              row.material,
                                                                              edit_blagovna_skupina_input.clone()),
                                );
                                row_data.query(db_manager, sort_state);
                            },
                            _ => {}
                        }
                    }

                } else {
                    let mut label_text = row.blagovna_skupina.clone().unwrap_or(" ".repeat(30));
                    if label_text.is_empty() {
                        label_text = " ".repeat(1);
                    }
                    let resp = ui.label(label_text.clone()).on_hover_cursor(CursorIcon::Help);
                    if resp.double_clicked() {
                        *editing_blagovna_skupina_row = Some(index);
                        *edit_blagovna_skupina_input = row.blagovna_skupina.clone().unwrap_or(String::new());
                    }

                }
            });
        }
        ViewQueryFields::Pakiranje => {
            table_row.col(|ui| {
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);

                if *editing_pakiranje_row == Some(index) {
                    let response = ui.text_edit_singleline(edit_pakiranje_input);
                    if response.lost_focus() {
                        *editing_pakiranje_row = None;

                        let os_resp = MessageDialog::new()
                            .set_title("Potrdi vnos")
                            .set_level(MessageLevel::Info)
                            .set_buttons(MessageButtons::OkCancel)
                            .show();

                        match os_resp {
                            MessageDialogResult::Ok => {
                                let _ = db_manager.store_pakiranje_to_server((
                                                                       row.material,
                                                                       edit_pakiranje_input.clone()),
                                );
                                row_data.query(db_manager, sort_state);
                            },
                            _ => {}
                        }
                    }

                } else {
                    let mut label_text = row.pakiranje.clone().unwrap_or_else(|| " ".repeat(20));
                    if label_text.is_empty() {
                        label_text = " ".repeat(20);
                    }
                    let resp = ui.label(label_text.clone()).on_hover_cursor(CursorIcon::Help);
                    if resp.double_clicked() {
                        *editing_pakiranje_row = Some(index);
                        *edit_pakiranje_input = row.pakiranje.clone().unwrap_or(String::new());
                    }

                }
            });
        }
        ViewQueryFields::Opomba => {
            table_row.col(|ui| {
                ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);

                if *editing_opomba_row == Some(index) {
                    let response = ui.text_edit_singleline(edit_opomba_input);
                    if response.lost_focus() {
                        *editing_opomba_row = None;

                        let os_resp = MessageDialog::new()
                            .set_title("Potrdi vnos")
                            .set_level(MessageLevel::Info)
                            .set_buttons(MessageButtons::OkCancel)
                            .show();

                        match os_resp {
                            MessageDialogResult::Ok => {
                                let _ = db_manager.store_opomba_to_server((
                                                                          row.material,
                                                                          edit_opomba_input.clone()),
                                );
                                row_data.query(db_manager, sort_state);
                            },
                            _ => {}
                        }
                    }

                } else {
                    let mut label_text = row.opomba.clone().unwrap_or_else(|| " ".repeat(73));
                    if label_text.is_empty() {
                        label_text = " ".repeat(73);
                    }
                    let resp = ui.label(label_text.clone()).on_hover_cursor(CursorIcon::Help);
                    if resp.double_clicked() {
                        *editing_opomba_row = Some(index);
                        *edit_opomba_input = row.opomba.clone().unwrap_or(String::new());
                    }

                }
            });
        }
    }
}



