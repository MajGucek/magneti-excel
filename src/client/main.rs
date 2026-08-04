#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]
#![allow(deprecated)] // I am not learning a whole new ecosystem

mod db;
mod graph;

mod config;

use std::sync::Arc;
use graph::*;


use std::time::{Duration, Instant};
use eframe::{NativeOptions};
use eframe::egui::*;
use eframe::egui::Ui;
use egui_extras::{Column, TableBuilder};
use env_logger::Env;
use rfd::{MessageDialog, MessageLevel};
use rust_xlsxwriter::{Format, Note, Workbook};
use crate::config::Config;
use crate::db::{DBManager, ViewQueryFields, SortState, ViewQuery};


static YELLOW: Color32 = Color32::YELLOW;
static ORANGE: Color32 = Color32::from_rgb(255, 153, 51);
static RED: Color32 = Color32::from_rgb(255, 135, 135);
static GREEN: Color32 = Color32::LIGHT_GREEN;
static BLUE: Color32 = Color32::LIGHT_BLUE;
static VIOLET: Color32 = Color32::from_rgb(191, 136, 187);
static TEAL: Color32 = Color32::from_rgb(0, 150, 136);
static INDIGO: Color32 = Color32::from_rgb(180, 100, 255);


struct App {
    db_address: String,
    db_manager: DBManager,
    last_query: Instant,

    row_data: Rows,
    config: Config,
    sort_state: SortState,

    poraba_nabava_data: PorabaNabavaRows,

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


    editing_dobavni_rok_row: Option<usize>,
    edit_dobavni_rok_input: String,

    editing_opomba_row: Option<usize>,
    edit_opomba_input: String,

    editing_min_zaloga_row: Option<usize>,
    edit_min_zaloga_row_input: String,

    editing_max_zaloga_row: Option<usize>,
    edit_max_zaloga_row_input: String,

    editing_blagovna_skupina_row: Option<usize>,
    edit_blagovna_skupina_input: String,

    editing_pakiranje_row: Option<usize>,
    edit_pakiranje_input: String,
}

impl App {
    pub fn new<'a>(cc: &'a eframe::CreationContext<'a>) -> Self {
        cc.egui_ctx.send_viewport_cmd(ViewportCommand::Maximized(true));
        cc.egui_ctx.set_visuals(Visuals::light());

        cc.egui_ctx.style_mut(|s| {
           s.override_text_style = Some(TextStyle::Heading);
            s.visuals.override_text_color = Some(Color32::from_rgb(5, 5, 5));
        });

        let config = Config::load();

        let mut row_data = None;
        let db_manager = DBManager::create(config.get_url());
        let sort_state = SortState::default();

        let result = db_manager.get_data(&sort_state);

        match result {
            Err(err) => {
                log::error!("initial_load error: {:?}", err.to_string());
                Some(false)
            },
            Ok(rows) => {
                row_data = Some(rows);
                log::info!("row_data loaded: {}", row_data.as_ref().unwrap().len());
                Some(true)
            }
        };

        let url = config.get_url();
        log::info!("Full url: {:?}", url);
        let db_address = url.trim_start_matches("http://")
            .split(":")
            .next()
            .unwrap_or("").to_string();
        Self {
            db_address,
            db_manager,
            last_query: Instant::now(),

            row_data: Rows {row_data},
            config,
            sort_state,
            poraba_nabava_data: PorabaNabavaRows::default(),

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

            editing_dobavni_rok_row: None,
            edit_dobavni_rok_input: String::new(),

            editing_opomba_row: None,
            edit_opomba_input: String::new(),

            editing_min_zaloga_row: None,
            edit_min_zaloga_row_input: String::new(),

            editing_max_zaloga_row: None,
            edit_max_zaloga_row_input: String::new(),

            editing_blagovna_skupina_row: None,
            edit_blagovna_skupina_input: String::new(),

            editing_pakiranje_row: None,
            edit_pakiranje_input: String::new(),

        }
    }

}




pub fn render_choose_panel(ctx: &Context, row_config: &mut Config) -> bool {
    let mut changed = false;

    let open_id = Id::new("column_choose_open");
    let mut open = ctx.data(|d| d.get_temp::<bool>(open_id)).unwrap_or(false);

    Area::new(Id::new("column_choose_button"))
        .anchor(Align2::RIGHT_TOP, [-8.0, 8.0])
        .show(ctx, |ui| {
            if ui.button("⚙").clicked() {
                open = !open;
            }
        });

    if open {
        Window::new("Stolpci")
            .id(Id::new("column_choose_window"))
            .resizable(true)
            .collapsible(false)
            .default_width(220.0)
            .anchor(Align2::RIGHT_TOP, [-8.0, 40.0])
            .open(&mut open)
            .show(ctx, |ui| {

                let mut move_up = None;
                let mut move_down = None;
                let mut remove_idx = None;
                let display_columns = row_config.get_mut_display_columns();
                let can_remove = display_columns.len() > 1;

                ScrollArea::vertical()
                    .id_salt("active_cols")
                    .max_height(280.0)
                    .show(ui, |ui| {
                        for (i, field) in display_columns.iter().enumerate() {
                            ui.horizontal(|ui| {
                                if ui.small_button("^").clicked() && i > 0 {
                                    move_up = Some(i);
                                }
                                if ui.small_button("v").clicked()
                                    && i + 1 < display_columns.len()
                                {
                                    move_down = Some(i);
                                }
                                ui.add_enabled_ui(can_remove, |ui| {
                                    if ui.small_button("x").clicked() {
                                        remove_idx = Some(i);
                                    }
                                });
                                ui.label(field.to_string());
                            });
                        }
                    });

                if let Some(i) = move_up {
                    display_columns.swap(i, i - 1);
                    changed = true;
                }
                if let Some(i) = move_down {
                    display_columns.swap(i, i + 1);
                    changed = true;
                }
                if let Some(i) = remove_idx {
                    if display_columns.len() > 1 {
                        display_columns.remove(i);
                        changed = true;
                    }
                }

                ui.separator();

                let mut to_add = None;
                ScrollArea::vertical()
                    .id_salt("avail_cols")
                    .max_height(280.0)
                    .show(ui, |ui| {
                        //let display_columns = row_config.get_mut_display_columns();
                        for field in ViewQueryFields::ALL {
                            if !display_columns.contains(&field) {
                                ui.horizontal(|ui| {
                                    if ui.small_button("+").clicked() {
                                        to_add = Some(field);
                                    }
                                    ui.label(field.to_string());
                                });
                            }
                        }
                    });

                if let Some(f) = to_add {
                    display_columns.insert(0, f);
                    changed = true;
                }

                ui.separator();
                if ui.button("Ponastavi").clicked() {
                    *row_config = Config::default();
                    changed = true;
                }
            });
    }

    ctx.data_mut(|d| d.insert_temp(open_id, open));

    changed
}

pub struct Rows {
    row_data: Option<Vec<ViewQuery>>
}

impl Rows {
    fn query(&mut self, db_manager: &DBManager, sort_state: &SortState) -> Option<bool> {
        log::info!("querying");
        match db_manager.get_data(&sort_state) {
            Ok(rows) => {
                self.row_data = Some(rows);
                log::info!("Loaded row_data");
                Some(true)
            }
            Err(err) => {
                log::error!("query_load error: {:?}", err);
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
                    let has_open_orders = row.odprta_narocila.is_some_and(|v| v != 0.);
                    let no_3m_no_24m = !(row.poraba_3m.is_some_and(|v| v == 0.) && row.poraba_24m.is_some_and(|v| v == 0.));

                    color_matches |= row.trenutna_zaloga_zadostuje_za_mesecev.is_some() && self.filter_rumena && row.dobavni_rok.is_some() && months_left >= 1.0 && months_left < 1.5 && no_open_orders && no_3m_no_24m;
                    color_matches |= row.trenutna_zaloga_zadostuje_za_mesecev.is_some() && self.filter_oranzna && row.dobavni_rok.is_some() && months_left >= 0.3 && months_left < 1.0 && no_open_orders && no_3m_no_24m;
                    color_matches |= row.trenutna_zaloga_zadostuje_za_mesecev.is_some() && self.filter_rdeca && row.dobavni_rok.is_some() && months_left < 0.3 && no_open_orders && row.dobavni_rok.unwrap_or(0.) < 90. && no_3m_no_24m;

                    color_matches |= self.filter_viola && row.dobavni_rok.is_some() && row.dobavni_rok.unwrap_or(0.) >= 90. &&
                        no_open_orders;

                    color_matches |= self.filter_modra && row.dobavni_rok.is_some() && has_open_orders &&
                        row.trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev.unwrap_or(0.) <= row.dobavni_rok.unwrap_or(0.);



                    color_matches |= self.filter_zelena && row.dobavni_rok.is_some() && has_open_orders &&
                        !(row.trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev.unwrap_or(0.) <= row.dobavni_rok.unwrap_or(0.));

                    color_matches |= self.filter_teal && row.minimalna_zaloga.is_some_and(|val| val > (row.zaloga.unwrap_or(0.) + row.odprta_narocila.unwrap_or(0.)));
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


    pub fn render_table(&mut self, ui: &mut Ui, data: &Vec<ViewQuery>) {
        ScrollArea::both().show(ui, |ui| {
            ui.style_mut().visuals.faint_bg_color = Color32::from_rgb(200, 200, 200);
            let display_columns = self.config.get_mut_display_columns();
            TableBuilder::new(ui)
                .striped(true)
                .cell_layout(Layout::left_to_right(Align::Center))
                .resizable(true)
                .columns(Column::auto().resizable(true), display_columns.len())
                .header(50.0, |mut header| {
                    display_columns.iter().for_each(|&column| {
                        db::construct_headers(column, &mut header, &mut self.sort_state.sort_column);
                    });
                })
                .body(|body| {
                    body.rows(25., data.len(), |mut table_row| {
                        let index = table_row.index().clone();
                        let row = &data[index];
                        let colors = calculate_colors(row);
                        let row_color = colors.last().cloned().unwrap_or(Color32::TRANSPARENT);

                        display_columns.iter().for_each(|&column| {
                            db::construct_body(
                                column,
                                &mut table_row,
                                index,
                                row,
                                row_color,
                                &mut self.poraba_nabava_data,
                                &self.db_manager,
                                &mut self.sort_state,
                                &mut self.row_data,
                                &mut self.editing_dobavni_rok_row,
                                &mut self.edit_dobavni_rok_input,
                                &mut self.editing_min_zaloga_row,
                                &mut self.edit_min_zaloga_row_input,
                                &mut self.editing_max_zaloga_row,
                                &mut self.edit_max_zaloga_row_input,
                                &mut self.editing_pakiranje_row,
                                &mut self.edit_pakiranje_input,
                                &mut self.editing_blagovna_skupina_row,
                                &mut self.edit_blagovna_skupina_input,
                                &mut self.editing_opomba_row,
                                &mut self.edit_opomba_input,
                            );
                        });

                        if table_row.response().hovered() {
                            table_row.response().ctx.layer_painter(LayerId::new(
                                Order::Foreground,
                                Id::new("hover"),
                            ))
                                .rect_stroke(
                                    table_row.response().rect,
                                    CornerRadius::same(0),
                                    Stroke::new(3.5, Color32::BLACK),
                                    StrokeKind::Outside
                                );
                        }

                    });
                });
        });
    }
}

fn calculate_colors(row: &ViewQuery) -> Vec<Color32> {
    let mut colors = Vec::with_capacity(6);
    colors.push(Color32::TRANSPARENT);


    if row.dobavni_rok.is_some() {
        if row.odprta_narocila.is_some_and(|v| v != 0.) {
            if row.trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev.unwrap_or(0.) <= row.dobavni_rok.unwrap_or(0.) {
                colors.push(BLUE);
            } else {
                colors.push(GREEN);
            }
        }

        if row.odprta_narocila.is_some_and(|v| v == 0.) && row.trenutna_zaloga_zadostuje_za_mesecev.is_some() {
            if row.trenutna_zaloga_zadostuje_za_mesecev.unwrap_or(0.) - row.dobavni_rok.unwrap_or(0.) < 1.5 {
                colors.push(YELLOW);
            }

            if row.trenutna_zaloga_zadostuje_za_mesecev.unwrap_or(0.) - row.dobavni_rok.unwrap_or(0.) < 1.0 {
                colors.push(ORANGE);
            }

            if row.trenutna_zaloga_zadostuje_za_mesecev.unwrap_or(0.) - row.dobavni_rok.unwrap_or(0.) < 0.3 {
                colors.push(RED);
            }
        }

        if row.odprta_narocila.is_some_and(|v| v == 0.) {
            if row.dobavni_rok.unwrap_or(0.) >= 90. {
                colors.push(VIOLET);
            }
        }
    }


    if row.poraba_3m.is_some_and(|v| v == 0.) && row.poraba_24m.is_some_and(|v| v == 0.) {
        colors.retain(|&color| {
            !(color == YELLOW || color == ORANGE || color == RED)
        });
    }


    colors
}




pub fn format_nabavnik<'a>(nabavna_skupina: &str) -> Option<&'a str> {
    match nabavna_skupina {
            "002" => Some("Neli"),
            "008" => Some("Viktoriia"),
            "010" => Some("Boštjan"),
            _ => None,
        }
}


pub fn format_number_custom(value: f64, precision: usize) -> String {
    let factor = 10_f64.powi(precision as i32);
    let scaled = (value * factor).round() as i64;

    let int_part = scaled / factor as i64;
    let mut dec_part = (scaled % factor as i64).abs().to_string();

    // pad decimals (important!)
    while dec_part.len() < precision {
        dec_part = format!("0{}", dec_part);
    }

    let int_str = int_part.abs().to_string();

    let int_part_with_sep = int_str
        .chars()
        .rev()
        .collect::<Vec<_>>()
        .chunks(3)
        .map(|c| c.iter().rev().collect::<String>())
        .collect::<Vec<_>>()
        .into_iter()
        .rev()
        .collect::<Vec<_>>()
        .join(".");

    let sign = if value < 0.0 { "-" } else { "" };

    if precision == 0 {
        format!("{}{}", sign, int_part_with_sep)
    } else {
        format!("{}{},{}", sign, int_part_with_sep, dec_part)
    }
}

impl eframe::App for App {

    fn update(&mut self, ctx: &Context, _frame: &mut eframe::Frame) {
        if self.last_query.elapsed() >= Duration::from_mins(5) {
            self.last_query = Instant::now();

            self.row_data.query(&self.db_manager, &self.sort_state);
            log::info!("Refreshed data");
        }

        ctx.request_repaint_after(Duration::from_millis(100));


        let mut has_row_data = false;
        let data = match &self.row_data.row_data {
            Some(d) => {
                has_row_data = true;
                self.apply_filters(d)
            },
            None => Vec::new(),
        };

        if data.len() == 0 && has_row_data {
            log::info!("After filters, there are no rows left, check if you have full server.sqlite3 with hand inputted data!");
        }


        let old_column = self.sort_state.sort_column;
        let old_sort = self.sort_state.descending;




        TopBottomPanel::top("Search").show(ctx, |ui| {
            menu::bar(ui, |ui| {
                ui.add(
                    TextEdit::singleline(&mut self.filter_sifra_materiala)
                        .hint_text("Iskanje po šifri materiala...")
                );

                ui.separator();
                ui.add(
                    TextEdit::singleline(&mut self.filter_naziv_materiala)
                        .hint_text("Iskanje po nazivu materiala...")
                );
                ui.separator();
                ui.add(
                    TextEdit::singleline(&mut self.filter_nabavnik)
                        .hint_text("Iskanje po nabavniku...")
                );
                ui.separator();
                ui.add(
                    TextEdit::singleline(&mut self.filter_dobavitelj)
                        .hint_text("Iskanje po dobavitelju...")
                );
                ui.separator();
                ui.add(
                    TextEdit::singleline(&mut self.filter_blagovna_skupina)
                        .hint_text("Iskanje po blagovni skupini...")
                );
            });
        });

        TopBottomPanel::top("checkboxi").show(ctx, |ui| {
           menu::bar(ui, |ui| {
               ui.horizontal(|ui| {
                   ui.vertical(|ui| {
                       ui.add_space(5.);

                       ui.horizontal(|ui| {
                           ui.checkbox(&mut self.filter_zaloga_vecja, "Zaloga večja kot");
                           ui.add(
                               TextEdit::singleline(&mut self.filter_zaloga_gt)
                                   .desired_width(25.)
                           );
                       });

                   });

                   ui.separator();

                   ui.horizontal(|ui| {
                       ui.checkbox(&mut self.filter_poraba_vecja, "Poraba večja kot");
                       ui.add(
                           TextEdit::singleline(&mut self.filter_poraba_gt)
                               .desired_width(25.)
                       );
                   });

                   ui.separator();

                   ui.horizontal(|ui| {
                       ui.checkbox(&mut self.filter_odprta_narocila, "Odprta naročila večja kot");
                       ui.add(
                           TextEdit::singleline(&mut self.filter_odprta_narocila_gt)
                               .desired_width(25.)
                       );
                   });

                   ui.separator();

                   ui.horizontal(|ui| {
                       ui.checkbox(&mut self.filter_dobavni_rok, "Dobavni rok večji kot");
                       ui.add(
                           TextEdit::singleline(&mut self.filter_dobavni_rok_gt)
                               .desired_width(25.)
                       );
                   });

                   ui.separator();

                   ui.horizontal(|ui| {
                       ui.checkbox(&mut self.filter_min_zaloga, "Min zaloga večja kot");
                       ui.add(
                           TextEdit::singleline(&mut self.filter_min_zaloga_gt)
                               .desired_width(25.)
                       );
                   });

                   ui.separator();

                   ui.horizontal(|ui| {
                       ui.checkbox(&mut self.filter_max_zaloga, "Max zaloga večja kot");
                       ui.add(
                           TextEdit::singleline(&mut self.filter_max_zaloga_gt)
                               .desired_width(25.)
                       );
                   });

                   ui.separator();

                   ui.checkbox(&mut self.filter_opomba, "Artikel ima opombo");

                   ui.separator();

                   ui.checkbox(&mut self.filter_no_dobavni_rok, "Nima dobavnega roka");
               });


           });
        });

        TopBottomPanel::top("barve").show(ctx, |ui| {
           menu::bar(ui, |ui| {
               ui.horizontal(|ui| {
                   ui.checkbox(&mut self.filter_rumena, RichText::new("Rumeni").color(YELLOW));
                   ui.separator();
                   ui.checkbox(&mut self.filter_oranzna, RichText::new("Oranžni").color(ORANGE));
                   ui.separator();
                   ui.checkbox(&mut self.filter_rdeca, RichText::new("Rdeči").color(RED));
                   ui.separator();
                   ui.checkbox(&mut self.filter_modra, RichText::new("Modri").color(BLUE));
                   ui.separator();
                   ui.checkbox(&mut self.filter_zelena, RichText::new("Zeleni").color(GREEN));
                   ui.separator();
                   ui.checkbox(&mut self.filter_viola, RichText::new("Viola").color(VIOLET));
                   ui.separator();
                   ui.checkbox(&mut self.filter_teal, RichText::new("Smaragdna").color(TEAL));
                   ui.separator();
                   ui.checkbox(&mut self.filter_indigo, RichText::new("Indigo").color(INDIGO));
               });
           });
        });

        TopBottomPanel::top("main").show(ctx, |ui| {
           menu::bar(ui, |ui| {
               if ui.add(
                   TextEdit::singleline(&mut self.db_address)
                       .hint_text("Naslov strežnika").desired_width(100.)
               ).changed() {
                   self.db_manager.update_url(self.db_address.as_str());
                   self.config.update_url(self.db_address.as_str());
               }

               if ui.button("Osveži").clicked() {
                   self.row_data.query(&self.db_manager, &self.sort_state);
                   log::info!("Manually refreshed data");
               }

               ui.separator();

               ui.horizontal(|ui| {
                   if ui.button("Izvozi").clicked() {
                       let mut resp: Result<(), Box<dyn std::error::Error>> = Ok(());

                       let _ = self.row_data.row_data.as_ref().map(|d| { resp = export_filtered_to_excel(&self.apply_filters(&d)); });
                       match resp {
                           Err(err) => {
                               log::error!("export error: {:?}", err);
                               MessageDialog::new().set_title("Napaka").set_description("Napaka pri izvozu").set_level(MessageLevel::Error).show();
                           },
                           Ok(_) => {
                               MessageDialog::new().set_title("Uspeh").set_description("Uspešno izvozil").set_level(MessageLevel::Info).show();
                           }
                       }
                   }
               });

               ui.separator();

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
                   log::info!("Refreshed data after filter reset");
               }


               ui.separator();

               ui.label(format!("Število zadetkov: {}", &data.len()));


               ui.separator();
               ui.radio_value(&mut self.sort_state.descending, false, "Naraščujoče");
               ui.radio_value(&mut self.sort_state.descending, true, "Padajoče");
           });
        });

        let changed = render_choose_panel(ctx, &mut self.config);
        if changed {
            self.config.save()
        }

        CentralPanel::default().show(ctx, |ui| {
            ScrollArea::vertical().show(ui, |ui| {
                self.render_table(ui, &data);
            });
        });


        let mut chart_clicked = false;
        Area::new(Id::from("chart"))
            .anchor(Align2::RIGHT_BOTTOM, [-25., -25.])
            .show(ctx, |ui| {
                chart_clicked = self.poraba_nabava_data.render(ui);
                if ui.input(|i| i.pointer.any_pressed()) && !chart_clicked {
                    self.poraba_nabava_data.clear();
                }
            });


        let new_column = self.sort_state.sort_column;
        let new_sort = self.sort_state.descending;

        if old_sort != new_sort || old_column != new_column {
            self.row_data.query(&self.db_manager, &self.sort_state);
            log::info!("Refreshed data due to old_sort != new_sort");
        }
    }


}


pub fn export_filtered_to_excel(
    data: &[ViewQuery],
) -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_default_note_author("Magneti Excel");

    worksheet.write_string(0, 0, "Material")?;
    worksheet.write_string(0, 1, "Naziv")?;
    worksheet.write_string(0, 2, "Zaloga 100")?;
    worksheet.write_string(0, 3, "Zaloga Sum")?;
    worksheet.write_string(0, 4, "Poraba 3M")?;
    worksheet.insert_note(0, 4, &Note::new("Povprečna mesečna poraba za zadnje 3 mesece"))?;

    worksheet.write_string(0, 5, "Poraba 24M")?;
    worksheet.insert_note(0, 5, &Note::new("Povprečna mesečna poraba za zadnjih 24 mesecev"))?;

    worksheet.write_string(0, 6, "Odprto")?;
    worksheet.insert_note(0, 6, &Note::new("Odprta naročila dobaviteljem"))?;

    worksheet.write_string(0, 7, "Dobava")?;
    worksheet.insert_note(0, 7, &Note::new("Predviden dobavni rok v mesecih"))?;

    worksheet.write_string(0, 8, "Zaloga SAP")?;
    worksheet.insert_note(0, 8, &Note::new("Trenutna zaloga v SAP-u, ki zadostuje za X mesecev na osnovi povprečne porabe preteklih 3 mesecev, če artikel nima 3M porabe računa na osnovi 24M porabe"))?;

    worksheet.write_string(0, 9, "Zaloga Sum SAP")?;
    worksheet.insert_note(0, 9, &Note::new("Seštevek trenutne zaloge v SAP-u in odprtih naročil, ki zadostuje za X mesecev na osnovi povprečne porabe preteklih 3 mesecev, če artikel nima 3M porabe računa na osnovi 24M porabe"))?;
    worksheet.write_string(0, 10, "Cena")?;
    worksheet.write_string(0, 11, "Valuta")?;
    worksheet.write_string(0, 12, "Enota")?;
    worksheet.write_string(0, 13, "Minimalna Zaloga")?;
    worksheet.write_string(0, 14, "Maximalna Zaloga")?;
    worksheet.write_string(0, 15, "Pakiranje")?;
    worksheet.write_string(0, 16, "Lokacija")?;
    worksheet.write_string(0, 17, "MRP")?;
    worksheet.write_string(0, 18, "Blagovna Skupina")?;
    worksheet.write_string(0, 19, "Opomba")?;
    worksheet.write_string(0, 20, "Nabavnik")?;
    worksheet.write_string(0, 21, "Dobavitelji")?;

    fn round_f64(value: f64, precision: u32) -> f64 {
        let factor = 10_f64.powi(precision as i32);
        (value * factor).round() / factor
    }

    for (row_idx, item) in data.iter().enumerate() {
        let row = (row_idx + 1) as u32;
        /*
        let [r, g, b, _a] = calculate_colors(item).last().map(|&c| if c == Color32::TRANSPARENT {Color32::WHITE} else {c}).unwrap_or(Color32::WHITE).to_array();
        let format = Format::new().set_background_color(Color::RGB(
            ((r as u32) << 16) | ((g as u32) << 8) | (b as u32)
        ));
         */

        let format = Format::default();



        let empty = String::new();
        worksheet.write_number_with_format(row, 0, item.material as f64, &format)?;
        worksheet.write_string_with_format(row, 1, item.naziv_materiala.as_ref().unwrap_or(&empty), &format)?;
        worksheet.write_number_with_format(row, 2, round_f64(item.razpolozljiva_zaloga.unwrap_or(0.), 1), &format)?;
        worksheet.write_number_with_format(row, 3, round_f64(item.zaloga.unwrap_or(0.), 1), &format)?;
        worksheet.write_number_with_format(row, 4, round_f64(item.poraba_3m.unwrap_or(0.), 1), &format)?;
        worksheet.write_number_with_format(row, 5, round_f64(item.poraba_24m.unwrap_or(0.), 1), &format)?;
        worksheet.write_number_with_format(row, 6, round_f64(item.odprta_narocila.unwrap_or(0.),0), &format)?;
        worksheet.write_number_with_format(row, 7, round_f64(item.dobavni_rok.unwrap_or(0.), 1), &format)?;
        worksheet.write_number_with_format(row, 8, round_f64(item.trenutna_zaloga_zadostuje_za_mesecev.unwrap_or(0.), 1), &format)?;
        worksheet.write_number_with_format(row, 9, round_f64(item.trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev.unwrap_or(0.), 1), &format)?;
        worksheet.write_number_with_format(row, 10, round_f64(item.cena.unwrap_or(0.), 1), &format)?;
        worksheet.write_string_with_format(row, 11, item.valuta.as_ref().unwrap_or(&empty), &format)?;
        worksheet.write_string_with_format(row, 12, item.osnovna_merska_enota.as_ref().unwrap_or(&empty), &format)?;
        worksheet.write_number_with_format(row, 13, round_f64(item.minimalna_zaloga.unwrap_or(0.), 0), &format)?;
        worksheet.write_number_with_format(row, 14, round_f64(item.maximalna_zaloga.unwrap_or(0.), 0), &format)?;
        worksheet.write_string_with_format(row, 15, item.pakiranje.as_ref().unwrap_or(&empty), &format)?;
        worksheet.write_string_with_format(row, 16, item.lokacija.as_ref().unwrap_or(&empty), &format)?;
        worksheet.write_string_with_format(row, 17, item.mrp_karakteristika.as_ref().unwrap_or(&empty), &format)?;
        worksheet.write_string_with_format(row, 18, item.blagovna_skupina.as_ref().unwrap_or(&empty), &format)?;
        worksheet.write_string_with_format(row, 19, item.opomba.as_ref().unwrap_or(&empty), &format)?;
        worksheet.write_string_with_format(row, 20, item.nabavna_skupina.as_ref().unwrap_or(&empty), &format)?;
        worksheet.write_string_with_format(row, 21, item.dobavitelji.as_ref().unwrap_or(&empty), &format)?;

    }

    workbook.save("../../Analitika.xlsx")?;
    Ok(())
}


fn parse_string_to_optional_f64(s: &str) -> Option<f64> {
    if s.eq("") {
        None
    } else {
        Some(s.replace(",", ".").parse().unwrap_or(0.))
    }
}


fn load_taskbar_icon() -> Arc<IconData> {
    let icon_bytes = include_bytes!("icon.png");

    let image = image::load_from_memory(icon_bytes)
        .expect("Failed to open icon path")
        .into_rgba8();

    let (width, height) = image.dimensions();

    Arc::new(IconData {
        rgba: image.into_raw(),
        width,
        height,
    })
}

fn main() {
    let debug = true;

    let level = if debug { "info" } else { "warn" };

    env_logger::Builder::from_env(
        Env::default().default_filter_or(level)
    )
        .init();

    log::info!("Client started");

    let icon = load_taskbar_icon();
    let options = NativeOptions {
        viewport: ViewportBuilder::default()
            .with_icon(icon)
            .with_title("Uporabnik"),
        ..Default::default()
    };

    eframe::run_native(
        "Uporabnik",
        options,
        Box::new(|cc| Ok(Box::new(App::new(cc))))
    ).unwrap();
}
