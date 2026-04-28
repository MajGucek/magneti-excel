#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]
#![allow(deprecated)] // I am not learning a whole new ecosystem

mod parse;
mod db;

use std::collections::HashMap;
use std::time::{Duration, Instant};
use chrono::{Datelike, Utc};
use eframe::{NativeOptions};
use eframe::egui::*;
use eframe::egui::Ui;
use egui_extras::{Column, TableBuilder};
use env_logger::Env;
use rfd::{MessageButtons, MessageDialog, MessageDialogResult, MessageLevel};
use rust_xlsxwriter::{Format, Workbook};
use crate::db::{DBManager, SortColumn, SortState, ViewQuery};
use crate::parse::{parse_all_files};


static YELLOW: Color32 = Color32::YELLOW;
static ORANGE: Color32 = Color32::from_rgb(255, 153, 51);
static RED: Color32 = Color32::from_rgb(255, 135, 135);
static GREEN: Color32 = Color32::LIGHT_GREEN;
static BLUE: Color32 = Color32::LIGHT_BLUE;
static VIOLET: Color32 = Color32::from_rgb(191, 136, 187);
static TEAL: Color32 = Color32::from_rgb(0, 150, 136);
static INDIGO: Color32 = Color32::from_rgb(180, 100, 255);


struct App {
    db_manager: DBManager,
    last_query: Instant,

    row_data: Rows,
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

        let mut row_data = None;
        let db_manager = DBManager { db_name: "magneti_db.sqlite3".to_string() };
        let _ = db_manager.try_create_tables();
        let sort_state = SortState::default();
        let _ = db_manager.try_drop_view();
        let _ = db_manager.try_create_view().unwrap();

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

        Self {
            db_manager,
            last_query: Instant::now(),

            row_data: Rows {row_data},
            sort_state,
            poraba_nabava_data: PorabaNabavaRows {material: 0, naziv: String::new(),  months: Vec::new(), poraba_nabava: Vec::new()},

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


struct PorabaNabavaRows {
    material: i64,
    naziv: String,
    months: Vec<String>,
    poraba_nabava: Vec<(f64, f64)>,
}
impl PorabaNabavaRows {
    fn clear(&mut self) {
        self.material = 0;
        self.naziv = String::new();
        self.months = Vec::new();
        self.poraba_nabava = Vec::new();
    }
    fn render(&self, ui: &mut Ui) -> bool {
        if self.months.is_empty() {
            return false;
        }

        ui.set_min_size(vec2(1500.0, 800.0));

        let title_rect = {
            let title_height = 100.0;
            Rect::from_min_max(
                pos2(ui.min_rect().left(), ui.min_rect().top()),
                pos2(ui.min_rect().right(), ui.min_rect().top() + title_height),
            )
        };

        ui.painter().rect_filled(
            title_rect,
            Rounding::same(0),
            Color32::WHITE,
        );

        ui.painter().text(
            pos2(title_rect.center().x, title_rect.center().y - 30.),
            Align2::CENTER_CENTER,
            &format!("Poraba - {}, {}", self.material, self.naziv),
            FontId::proportional(20.0),
            Color32::BLACK,
        );

        ui.painter().text(
            pos2(title_rect.center().x - 50., title_rect.center().y - 0.),
            Align2::CENTER_CENTER,
            "Poraba",
            FontId::proportional(20.0),
            Color32::GREEN,
        );

        ui.painter().text(
            pos2(title_rect.center().x + 50., title_rect.center().y - 0.),
            Align2::CENTER_CENTER,
            "Nabava",
            FontId::proportional(20.0),
            Color32::LIGHT_RED,
        );


        let rect = ui.max_rect();  // Chart area (below title)
        let padding = vec2(60.0, 40.0);

        let plot_rect = Rect::from_min_max(
            pos2(rect.left(), rect.top() + 80.),  // Skip title space
            pos2(rect.right(), rect.bottom()),
        );

        let slot_width = (plot_rect.width() - padding.x * 2.0) / self.months.len() as f32;
        let bar_width = slot_width * 0.7;

        let max_value = self.poraba_nabava.iter().fold(f64::NEG_INFINITY, |a, &b| a.max(
            if b.0 > b.1 {b.0} else {b.1}
        )).max(1.0);

        // Background
        ui.painter().rect_filled(plot_rect, Rounding::same(0), Color32::from_gray(240));

        // Grid lines
        let y_step = max_value / 10.0;
        for i in 0..=10 {
            let y_value = (i as f64) * y_step;
            let y_pos = plot_rect.bottom() - padding.y -
                (y_value / max_value * (plot_rect.height() - padding.y * 2.0) as f64) as f32;

            ui.painter().text(
                pos2(plot_rect.left() + 40.0, y_pos),
                Align2::RIGHT_CENTER,
                &format!("{:.0}", y_value),
                FontId::proportional(11.0),
                Color32::BLACK,
            );

            ui.painter().line_segment(
                [
                    pos2(plot_rect.left() + padding.x, y_pos),
                    pos2(plot_rect.right() - padding.x, y_pos)
                ],
                Stroke::new(0.5, Color32::from_gray(180)),
            );
        }

        // Vertical grid
        for i in (0..self.months.len()).step_by(3) {
            let x = plot_rect.left() + padding.x + (i as f32) * slot_width;
            ui.painter().line_segment(
                [
                    pos2(x, plot_rect.top() + padding.y),
                    pos2(x, plot_rect.bottom() - padding.y)
                ],
                Stroke::new(0.3, Color32::from_gray(190)),
            );
        }

        for (i, ((poraba, nabava), month)) in self.poraba_nabava.iter().zip(&self.months).enumerate() {
            let x = plot_rect.left() + padding.x + (i as f32) * slot_width + (slot_width - bar_width) / 2.0;


            let poraba_bar_height = (poraba / max_value * (plot_rect.height() - padding.y * 2.0) as f64) as f32;
            let poraba_bar_rect = Rect::from_min_max(
                pos2(x, plot_rect.bottom() - padding.y - poraba_bar_height),
                pos2(x + bar_width, plot_rect.bottom() - padding.y),
            );

            let nabava_bar_height = (nabava / max_value * (plot_rect.height() - padding.y * 2.0) as f64) as f32;
            let nabava_bar_rect = Rect::from_min_max(
                pos2(x, plot_rect.bottom() - padding.y - nabava_bar_height),
                pos2(x + bar_width, plot_rect.bottom() - padding.y),
            );

            let green = Color32::from_rgba_unmultiplied(0, 255, 0, 120);
            let dark_green = Color32::from_rgba_unmultiplied(0, 100, 0, 180);

            let red = Color32::from_rgba_unmultiplied(255, 0, 0, 120);
            let light_red = Color32::from_rgba_unmultiplied(255, 100, 100, 120);

            let (first, second) = if poraba_bar_height > nabava_bar_height {
                (
                    (poraba_bar_rect, green, dark_green),
                    (nabava_bar_rect, light_red, red),
                )
            } else {
                (
                    (nabava_bar_rect, light_red, red),
                    (poraba_bar_rect, green, dark_green),
                )
            };

            for (rect, fill, stroke_color) in [first, second] {
                ui.painter().rect(
                    rect,
                    0.0,
                    fill,
                    Stroke::new(1.0, stroke_color),
                    StrokeKind::Inside,
                );
            }



            ui.painter().text(
                pos2(x + bar_width / 2.0, plot_rect.bottom() - 20.0),
                Align2::CENTER_CENTER,
                month.split("-").collect::<Vec<&str>>().join("\n  "),
                FontId::proportional(12.0),
                Color32::BLACK,
            );
        }

        ui.interact(rect, ui.id(), Sense::click()).clicked()
    }

    fn query(&mut self, material: i64, naziv: &str, db_manager: &DBManager) {
        let raw_poraba_data: Vec<(String, f64)> = match db_manager.get_poraba(material) {
            Ok(rows) => {
                let mut data = Vec::new();
                for row in rows {
                    data.push((row.month.clone(), row.poraba));
                }
                data
            }
            Err(err) => {
                log::error!("DB error: {:?}", err);
                Vec::new()
            }
        };

        let raw_nabava_data: Vec<(String, f64)> = match db_manager.get_nabava(material) {
            Ok(rows) => {
                let mut data = Vec::new();
                for row in rows {
                    data.push((row.month.clone(), row.nabava));
                }
                data
            }
            Err(err) => {
                log::error!("DB error: {:?}", err);
                Vec::new()
            }
        };

        let mut map: HashMap<String, (f64, f64)> = HashMap::new();
        for (month, poraba) in raw_poraba_data {
            map.insert(month, (poraba, 0.));
        }



        for (month, nabava) in raw_nabava_data {
            map.entry(month)
                .and_modify(|e| e.1 = nabava)
                .or_insert((0., nabava));
        }
        let raw_data: Vec<(String, f64, f64)> = map.into_iter().map(|(m, (p, n))| (m, p, n))
            .collect();


        if raw_data.is_empty() {
            self.clear();
        }

        let month_data: Vec<((i32, u32), String, (f64, f64))> = raw_data
            .iter()
            .filter_map(|(month_str, poraba, nabava)| {
                let parts: Vec<&str> = month_str.split('-').collect();
                if parts.len() == 2 {
                    if let (Ok(year), Ok(month)) = (
                        parts[0].parse::<i32>(),
                        parts[1].parse::<u32>()
                    ) {
                        Some(((year, month), month_str.clone(), (*poraba, *nabava)))
                    } else {
                        None
                    }
                } else {
                    None
                }
            })
            .collect();

        let mut month_data = month_data;
        month_data.sort_by_key(|(m, _, _)| *m);

        let today_year = Utc::now().year();
        let today_month = Utc::now().month();

        let mut months = Vec::new();
        let mut poraba_nabava = Vec::new();

        let first = month_data[0].clone();
        months.push(first.1.clone());
        poraba_nabava.push(first.2);
        let mut prev_year = first.0 .0;
        let mut prev_month = first.0 .1;

        for ((year, month), month_str, value) in month_data.iter().skip(1) {
            let mut current_month = prev_month + 1;
            let mut current_year = prev_year;
            if current_month > 12 {
                current_month = 1;
                current_year += 1;
            }

            while current_year < *year || (current_year == *year && current_month < *month) {
                months.push(format!("{:04}-{:02}", current_year, current_month));
                poraba_nabava.push((0.0, 0.0));
                current_month += 1;
                if current_month > 12 {
                    current_month = 1;
                    current_year += 1;
                }
            }

            months.push(month_str.clone());
            poraba_nabava.push(*value);
            prev_year = *year;
            prev_month = *month;
        }


        let mut current_month = prev_month + 1;
        let mut current_year = prev_year;
        while current_year < today_year || (current_year == today_year && current_month <= today_month) {
            months.push(format!("{:04}-{:02}", current_year, current_month));
            poraba_nabava.push((0.0, 0.0));
            current_month += 1;
            if current_month > 12 {
                current_month = 1;
                current_year += 1;
            }
        }


        self.material = material;
        self.naziv = naziv.to_string();
        self.months = months;
        self.poraba_nabava = poraba_nabava;

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

                    color_matches |= self.filter_rumena && row.dobavni_rok.is_some() && months_left >= 1.0 && months_left < 1.5 && no_open_orders && no_3m_no_24m;
                    color_matches |= self.filter_oranzna && row.dobavni_rok.is_some() && months_left >= 0.3 && months_left < 1.0 && no_open_orders && no_3m_no_24m;
                    color_matches |= self.filter_rdeca && row.dobavni_rok.is_some() && months_left < 0.3 && no_open_orders && row.dobavni_rok.unwrap_or(0.) < 90. && no_3m_no_24m;

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
            TableBuilder::new(ui)
                .striped(true)
                .cell_layout(Layout::left_to_right(Align::Center))
                .columns(Column::exact(100.), 1) // Material
                .columns(Column::exact(300.), 1)// Naziv
                .columns(Column::exact(85.), 1) // Zaloga
                .columns(Column::exact(120.5), 1)// Poraba 3M
                .columns(Column::exact(120.5), 1)// Poraba 24M
                .columns(Column::exact(90.), 1)// Odprto
                .columns(Column::exact(90.), 1)// Dobava
                .columns(Column::exact(110.), 1)// Zaloga SAP
                .columns(Column::exact(120.), 1)// Sum Zaloga
                .columns(Column::exact(120.), 1)// Enota
                .columns(Column::exact(120.), 1)// Minimalna zaloga
                .columns(Column::exact(120.), 1)// Maximalna zaloga
                .columns(Column::exact(120.), 1)// Pakiranje
                .columns(Column::exact(160.), 1)// Blagovna Skupina
                .columns(Column::exact(300.), 1)// Opomba
                .columns(Column::exact(110.), 1)// Nabavnik
                .columns(Column::remainder(), 1)//Dobavitelji

                .header(50.0, |mut header| {
                    header.col(|ui| {ui.radio_value(&mut self.sort_state.sort_column, SortColumn::Material, "Material"); });
                    header.col(|ui| {ui.radio_value(&mut self.sort_state.sort_column, SortColumn::NazivMateriala, "Naziv"); });


                    header.col(|ui| {ui.radio_value(&mut self.sort_state.sort_column, SortColumn::Zaloga, "Zaloga").on_hover_text("Trenutna zaloga v SAP-u"); });
                    header.col(|ui| {ui.radio_value(&mut self.sort_state.sort_column, SortColumn::Poraba3M, "Poraba 3M").on_hover_text("Povprečna mesečna poraba za zadnje 3 mesece"); });
                    header.col(|ui| {ui.radio_value(&mut self.sort_state.sort_column, SortColumn::Poraba24M, "Poraba 24M").on_hover_text("Povprečna mesečna poraba za zadnjih 12 mesecev"); });
                    header.col(|ui| {ui.radio_value(&mut self.sort_state.sort_column, SortColumn::OdprtaNarocila, "Odprto").on_hover_text("Odprta naročila dobaviteljem"); });
                    header.col(|ui| {ui.radio_value(&mut self.sort_state.sort_column, SortColumn::DobavniRok, "Dobava").on_hover_text("Predviden dobavni rok v mesecih"); });

                    header.col(|ui| {ui.radio_value(&mut self.sort_state.sort_column, SortColumn::TrenutnaZalogaZadostujeZaMesecev, "Zaloga SAP").on_hover_text("Trenutna zaloga v SAP-u, ki zadostuje za X mesecev na osnovi povprečne porabe preteklih 3 mesecev, če artikel nima 3M porabe računa na osnovi 24M porabe"); });

                    header.col(|ui| {ui.radio_value(&mut self.sort_state.sort_column, SortColumn::TrenutnaZalogaInOdprtaNarocilaZadostujeZaMesecev, "Sum Zaloga").on_hover_text("Seštevek trenutne zaloge v SAP-u in odprtih naročil, ki zadostuje za X mesecev na osnovi povprečne porabe preteklih 3 mesecev, če artikel nima 3M porabe računa na osnovi 24M porabe"); });
                    header.col(|ui| {ui.radio_value(&mut self.sort_state.sort_column, SortColumn::OsnovnaMerskaEnota, "Enota"); });
                    header.col(|ui| {ui.radio_value(&mut self.sort_state.sort_column, SortColumn::MinimalnaZaloga, "Min zaloga"); });
                    header.col(|ui| {ui.radio_value(&mut self.sort_state.sort_column, SortColumn::MaximalnaZaloga, "Max zaloga"); });
                    header.col(|ui| {ui.radio_value(&mut self.sort_state.sort_column, SortColumn::Pakiranje, "Pakiranje"); });
                    header.col(|ui| {ui.radio_value(&mut self.sort_state.sort_column, SortColumn::BlagovnaSkupina, "Blagovna Skupina"); });
                    header.col(|ui| {ui.radio_value(&mut self.sort_state.sort_column, SortColumn::Opomba, "Opomba"); });
                    header.col(|ui| {ui.radio_value(&mut self.sort_state.sort_column, SortColumn::NabavnaSkupina, "Nabavnik").on_hover_text("002 Neli\n008 Viktoriia\n010 Boštjan"); });
                    header.col(|ui| {ui.radio_value(&mut self.sort_state.sort_column, SortColumn::Dobavitelji, "Dobavitelji"); });
                    //header.col(|ui| {ui.heading("MRP"); });
                })
                .body(|body| {
                    body.rows(25., data.len(), |mut table_row| {
                        let index = table_row.index().clone();

                        let row = &data[index];
                        let colors = calculate_colors(row);
                        let mut row_color = colors.last().cloned().unwrap_or(Color32::TRANSPARENT);


                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            if ui.label(RichText::new(row.material.to_string()).underline().background_color(Color32::TRANSPARENT))
                                .on_hover_cursor(CursorIcon::PointingHand)
                                .clicked() {

                                self.poraba_nabava_data.query(row.material, row.naziv_materiala.as_ref().unwrap_or(&"".to_string()).as_str(),  &self.db_manager);
                            }
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.naziv_materiala.clone().unwrap_or_else(|| "".to_string()));
                        });





                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.zaloga.map_or("".to_string(), |v| format_number_custom(v, 1)));
                        });

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

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.poraba_24m.map_or("".to_string(), |v| format_number_custom(v, 1)));
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.odprta_narocila.map_or("".to_string(), |v| format_number_custom(v, 0)));
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);

                            if self.editing_dobavni_rok_row == Some(index) {
                                let response = ui.text_edit_singleline(&mut self.edit_dobavni_rok_input);
                                if response.lost_focus() {
                                    self.editing_dobavni_rok_row = None;

                                    let os_resp = MessageDialog::new()
                                        .set_title("Potrdi vnos")
                                        .set_level(MessageLevel::Info)
                                        .set_buttons(MessageButtons::OkCancel)
                                        .show();

                                    match os_resp {
                                        MessageDialogResult::Ok => {
                                            let _ = self.db_manager.store_dobavni_rok((
                                                row.material,
                                                parse_string_to_optional_f64(self.edit_dobavni_rok_input.as_str()),
                                            ));
                                            self.row_data.query(&self.db_manager, &self.sort_state);
                                        },
                                        _ => {}
                                    }
                                }

                            } else {
                                let label_text = row.dobavni_rok.map_or(" ".repeat(18), |v| format_number_custom(v, 1));
                                let resp = ui.label(label_text).on_hover_cursor(CursorIcon::Help);
                                if resp.double_clicked() {
                                    self.editing_dobavni_rok_row = Some(index);
                                    self.edit_dobavni_rok_input = row.dobavni_rok.map_or("".to_string(), |v| format!("{}", v));
                                }

                            }

                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.trenutna_zaloga_zadostuje_za_mesecev.map_or("".to_string(), |v| format_number_custom(v, 1)));
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            ui.label(row.trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev.map_or("".to_string(), |v| format_number_custom(v, 1)));
                        });


                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                            let t = row.osnovna_merska_enota.clone().unwrap_or_else(|| "".to_string());
                            ui.label(&t);
                        });

                        table_row.col(|ui| {
                            let old = row_color;
                            if row.minimalna_zaloga.is_some_and(|val| val > (row.zaloga.unwrap_or(0.) + row.odprta_narocila.unwrap_or(0.)))  {
                                // teal
                                row_color = TEAL;
                            }
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);

                            if self.editing_min_zaloga_row == Some(index) {
                                let response = ui.text_edit_singleline(&mut self.edit_min_zaloga_row_input);
                                if response.lost_focus() {
                                    self.editing_min_zaloga_row = None;

                                    let os_resp = MessageDialog::new()
                                        .set_title("Potrdi vnos")
                                        .set_level(MessageLevel::Info)
                                        .set_buttons(MessageButtons::OkCancel)
                                        .show();

                                    match os_resp {
                                        MessageDialogResult::Ok => {
                                            let _ = self.db_manager.store_min_zaloga((
                                                row.material,
                                                self.edit_min_zaloga_row_input.clone().parse::<f64>().ok()),
                                            );
                                            self.row_data.query(&self.db_manager, &self.sort_state);
                                        },
                                        _ => {}
                                    }
                                }

                            } else {
                                let label_text = row.minimalna_zaloga.map_or(" ".repeat(28), |v| format_number_custom(v, 0));
                                let resp = ui.label(label_text).on_hover_cursor(CursorIcon::Help);
                                if resp.double_clicked() {
                                    self.editing_min_zaloga_row = Some(index);
                                    self.edit_min_zaloga_row_input = row.minimalna_zaloga.map_or("".to_string(), |v| format_number_custom(v, 0));
                                }

                            }


                            row_color = old;
                        });
                        table_row.col(|ui| {
                            let old = row_color;
                            if row.maximalna_zaloga.is_some_and(|val| val < row.zaloga.unwrap_or(0.)) {
                                // indigo
                                row_color = INDIGO;
                            }

                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);

                            if self.editing_max_zaloga_row == Some(index) {
                                let response = ui.text_edit_singleline(&mut self.edit_max_zaloga_row_input);
                                if response.lost_focus() {
                                    self.editing_max_zaloga_row = None;

                                    let os_resp = MessageDialog::new()
                                        .set_title("Potrdi vnos")
                                        .set_level(MessageLevel::Info)
                                        .set_buttons(MessageButtons::OkCancel)
                                        .show();

                                    match os_resp {
                                        MessageDialogResult::Ok => {
                                            let _ = self.db_manager.store_max_zaloga((
                                                                                         row.material,
                                                                                         self.edit_max_zaloga_row_input.clone().parse::<f64>().ok()),
                                            );
                                            self.row_data.query(&self.db_manager, &self.sort_state);
                                        },
                                        _ => {}
                                    }
                                }

                            } else {
                                let label_text = row.maximalna_zaloga.map_or(" ".repeat(28), |v| format_number_custom(v, 0));
                                let resp = ui.label(label_text).on_hover_cursor(CursorIcon::Help);
                                if resp.double_clicked() {
                                    self.editing_max_zaloga_row = Some(index);
                                    self.edit_max_zaloga_row_input = row.maximalna_zaloga.map_or("".to_string(), |v| format_number_custom(v, 0));
                                }

                            }


                            row_color = old;
                        });


                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);

                            if self.editing_pakiranje_row == Some(index) {
                                let response = ui.text_edit_singleline(&mut self.edit_pakiranje_input);
                                if response.lost_focus() {
                                    self.editing_pakiranje_row = None;

                                    let os_resp = MessageDialog::new()
                                        .set_title("Potrdi vnos")
                                        .set_level(MessageLevel::Info)
                                        .set_buttons(MessageButtons::OkCancel)
                                        .show();

                                    match os_resp {
                                        MessageDialogResult::Ok => {
                                            let _ = self.db_manager.store_pakiranje((
                                                                                               row.material,
                                                                                               self.edit_pakiranje_input.clone()),
                                            );
                                            self.row_data.query(&self.db_manager, &self.sort_state);
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
                                    self.editing_pakiranje_row = Some(index);
                                    self.edit_pakiranje_input = row.pakiranje.clone().unwrap_or(String::new());
                                }

                            }
                        });

                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);

                            if self.editing_blagovna_skupina_row == Some(index) {
                                let response = ui.text_edit_singleline(&mut self.edit_blagovna_skupina_input);
                                if response.lost_focus() {
                                    self.editing_blagovna_skupina_row = None;

                                    let os_resp = MessageDialog::new()
                                        .set_title("Potrdi vnos")
                                        .set_level(MessageLevel::Info)
                                        .set_buttons(MessageButtons::OkCancel)
                                        .show();

                                    match os_resp {
                                        MessageDialogResult::Ok => {
                                            let _ = self.db_manager.store_blagovna_skupina((
                                                                                           row.material,
                                                                                           self.edit_blagovna_skupina_input.clone()),
                                            );
                                            self.row_data.query(&self.db_manager, &self.sort_state);
                                        },
                                        _ => {}
                                    }
                                }

                            } else {
                                let mut label_text = row.blagovna_skupina.clone().unwrap_or(" ".repeat(73));
                                if label_text.is_empty() {
                                    label_text = " ".repeat(73);
                                }
                                let resp = ui.label(label_text.clone()).on_hover_cursor(CursorIcon::Help);
                                if resp.double_clicked() {
                                    self.editing_blagovna_skupina_row = Some(index);
                                    self.edit_blagovna_skupina_input = row.blagovna_skupina.clone().unwrap_or(String::new());
                                }

                            }
                        });


                        table_row.col(|ui| {
                            ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);

                            if self.editing_opomba_row == Some(index) {
                                let response = ui.text_edit_singleline(&mut self.edit_opomba_input);
                                if response.lost_focus() {
                                    self.editing_opomba_row = None;

                                    let os_resp = MessageDialog::new()
                                        .set_title("Potrdi vnos")
                                        .set_level(MessageLevel::Info)
                                        .set_buttons(MessageButtons::OkCancel)
                                        .show();

                                    match os_resp {
                                        MessageDialogResult::Ok => {
                                            let _ = self.db_manager.store_opomba_to_db((
                                                row.material,
                                                self.edit_opomba_input.clone()),
                                            );
                                            self.row_data.query(&self.db_manager, &self.sort_state);
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
                                    self.editing_opomba_row = Some(index);
                                    self.edit_opomba_input = row.opomba.clone().unwrap_or(String::new());
                                }

                            }
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

        if row.odprta_narocila.is_some_and(|v| v == 0.) {
            if row.trenutna_zaloga_zadostuje_za_mesecev.unwrap_or(0.) - row.dobavni_rok.unwrap_or(0.) < 1.5 {
                colors.push(YELLOW);
            }

            if row.trenutna_zaloga_zadostuje_za_mesecev.unwrap_or(0.) - row.dobavni_rok.unwrap_or(0.) < 1.0 {
                colors.push(ORANGE);
            }

            if row.trenutna_zaloga_zadostuje_za_mesecev.unwrap_or(0.) - row.dobavni_rok.unwrap_or(0.) < 0.3 {
                colors.push(RED);
            }

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




fn format_nabavnik<'a>(nabavna_skupina: &str) -> Option<&'a str> {
    match nabavna_skupina {
            "002" => Some("Neli"),
            "008" => Some("Viktoriia"),
            "010" => Some("Boštjan"),
            _ => None,
        }
}


fn format_number_custom(value: f64, precision: usize) -> String {
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
        }

        ctx.request_repaint_after(Duration::from_millis(100));

        let data = match &self.row_data.row_data {
            Some(d) => self.apply_filters(d),
            None => Vec::new(),
        };


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
               if ui.button("Osveži").clicked() {
                   self.row_data.query(&self.db_manager, &self.sort_state);
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
               }


               ui.separator();

               if ui.button("Vnos Excel-ov").clicked() {
                   let dir = dirs_next::document_dir().unwrap_or_else(|| std::path::PathBuf::from("C:\\"));
                   let files = rfd::FileDialog::new()
                       .set_title("Izberi vse datoteke")
                       .set_directory(dir)
                       .pick_files();
                   if files.is_some() {
                       match self.db_manager.drop_non_permanent() {
                           Err(e) => log::error!("dropping error: {}", e),
                           Ok(_) => log::info!("Successfully dropped data")
                       }
                       let res = parse_all_files(files.unwrap(), &self.db_manager);
                       ui.ctx().request_repaint();
                       match res {
                           Ok(_) => {
                               MessageDialog::new()
                                   .set_title("Uspeh")
                                   .set_description("Shranil Excel-e")
                                   .set_level(MessageLevel::Info)
                                   .show();
                           },
                           Err(e) => {
                               MessageDialog::new()
                                   .set_title("Napaka")
                                   .set_description(format!("Napaka pri obdelavi Excel-ov\n {:?}", e.to_string()))
                                   .set_level(MessageLevel::Error)
                                   .show();
                           },
                       }
                   }

               }

               ui.separator();

               ui.label(format!("Število zadetkov: {}", &data.len()));


               ui.separator();
               ui.radio_value(&mut self.sort_state.descending, false, "Naraščujoče");
               ui.radio_value(&mut self.sort_state.descending, true, "Padajoče");
           });
        });

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
        }
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
    let debug = true;

    let level = if debug { "info" } else { "warn" };

    env_logger::Builder::from_env(
        Env::default().default_filter_or(level)
    )
        .init();

    log::info!("App started");


    eframe::run_native(
        "Magneti Excel",
        NativeOptions::default(),
        Box::new(|cc| Ok(Box::new(App::new(cc))))
    ).unwrap();
}
