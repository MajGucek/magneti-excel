use std::collections::HashMap;
use chrono::{Datelike, Utc};
use eframe::egui::{pos2, vec2, Align2, Color32, FontId, Rect, Rounding, Sense, Stroke, StrokeKind, Ui};
use crate::db::DBManager;

#[derive(Default)]
pub struct PorabaNabavaRows {
    material: i64,
    naziv: String,
    months: Vec<String>,
    poraba_nabava: Vec<(f64, f64, f64)>,
}
impl PorabaNabavaRows {
    pub fn clear(&mut self) {
        self.material = 0;
        self.naziv = String::new();
        self.months = Vec::new();
        self.poraba_nabava = Vec::new();
    }
    pub fn render(&self, ui: &mut Ui) -> bool {
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
            &format!("Poraba/Nabava/Zaloga - {}, {}", self.material, self.naziv),
            FontId::proportional(20.0),
            Color32::BLACK,
        );

        ui.painter().text(
            pos2(title_rect.center().x - 125., title_rect.center().y - 0.),
            Align2::CENTER_CENTER,
            "Poraba",
            FontId::proportional(20.0),
            Color32::LIGHT_RED,
        );

        ui.painter().text(
            pos2(title_rect.center().x + 0., title_rect.center().y - 0.),
            Align2::CENTER_CENTER,
            "Nabava",
            FontId::proportional(20.0),
            Color32::GREEN,
        );

        ui.painter().text(
            pos2(title_rect.center().x + 125., title_rect.center().y - 0.),
            Align2::CENTER_CENTER,
            "Zaloga",
            FontId::proportional(20.0),
            Color32::BLUE,
        );


        let rect = ui.max_rect();
        let padding = vec2(60.0, 40.0);

        let plot_rect = Rect::from_min_max(
            pos2(rect.left(), rect.top() + 80.),
            pos2(rect.right(), rect.bottom()),
        );

        let slot_width = (plot_rect.width() - padding.x * 2.0) / self.months.len() as f32;
        let bar_width = slot_width * 0.75;

        let max_value = self.poraba_nabava
            .iter()
            .fold(f64::NEG_INFINITY, |a, &(x, y, z)| {
                a.max(x.max(y).max(z))
            })
            .max(1.0);

        ui.painter().rect_filled(plot_rect, Rounding::same(0), Color32::from_gray(240));

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

        for (i, ((poraba, nabava, zaloga), month)) in self.poraba_nabava.iter().zip(&self.months).enumerate() {
            let x = plot_rect.left() + padding.x + (i as f32) * slot_width + (slot_width - bar_width) / 2.0;


            let poraba_bar_height = (poraba / max_value * (plot_rect.height() - padding.y * 2.0) as f64) as f32;
            let poraba_thickness = 1./5.;
            let bar_center = x + bar_width / 3.0;
            let poraba_width = bar_width * poraba_thickness;
            let poraba_bar_rect = Rect::from_min_max(
                pos2(bar_center - poraba_width / 2.0, plot_rect.bottom() - padding.y - poraba_bar_height, ),
                pos2(bar_center + poraba_width / 2.0, plot_rect.bottom() - padding.y, ),
            );


            let nabava_bar_height = (nabava / max_value * (plot_rect.height() - padding.y * 2.0) as f64) as f32;
            let nabava_thickness = 1./5.;
            let bar_center = x + bar_width * (2.0 / 3.0);
            let nabava_width = bar_width * nabava_thickness;
            let nabava_bar_rect = Rect::from_min_max(
                pos2(bar_center - nabava_width / 2.0, plot_rect.bottom() - padding.y - nabava_bar_height, ),
                pos2(bar_center + nabava_width / 2.0, plot_rect.bottom() - padding.y, ),
            );

            let zaloga_bar_height = (zaloga / max_value * (plot_rect.height() - padding.y * 2.0) as f64) as f32;
            let zaloga_thickness = 1./5.;
            let bar_center = x + bar_width * 1.;
            let zaloga_width = bar_width * zaloga_thickness;
            let zaloga_bar_rect = Rect::from_min_max(
                pos2(bar_center - zaloga_width / 2.0, plot_rect.bottom() - padding.y - zaloga_bar_height, ),
                pos2(bar_center + zaloga_width / 2.0, plot_rect.bottom() - padding.y, ),
            );

            let green = Color32::from_rgb(0, 255, 0);
            let dark_green = Color32::from_rgb(0, 100, 0);

            let red = Color32::from_rgb(255, 0, 0);
            let light_red = Color32::from_rgb(255, 100, 100);

            let blue = Color32::from_rgb(0, 0, 255);
            let light_blue = Color32::from_rgb(100, 100, 255);

            let mut bars = [
                (poraba_bar_rect, light_red, red),
                (nabava_bar_rect, green, dark_green),
                (zaloga_bar_rect, light_blue, blue),
            ];
            bars.sort_by_key(|(rect, _, _)| {
                -(rect.height() as i32)
            });

            for (rect, fill, stroke_color) in bars {
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

    pub fn query(&mut self, material: i64, naziv: &str, zaloga_sum: f64,  db_manager: &DBManager) {
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
        let mut poraba_nabava: Vec<(f64, f64, f64)> = Vec::new();

        let first = month_data[0].clone();
        months.push(first.1.clone());
        poraba_nabava.push((first.2.0, first.2.1, 0.));
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
                poraba_nabava.push((0.0, 0.0, 0.));
                current_month += 1;
                if current_month > 12 {
                    current_month = 1;
                    current_year += 1;
                }
            }

            months.push(month_str.clone());
            poraba_nabava.push((value.0, value.1, 0.));
            prev_year = *year;
            prev_month = *month;
        }


        let mut current_month = prev_month + 1;
        let mut current_year = prev_year;
        while current_year < today_year || (current_year == today_year && current_month <= today_month) {
            months.push(format!("{:04}-{:02}", current_year, current_month));
            poraba_nabava.push((0.0, 0.0, 0.));
            current_month += 1;
            if current_month > 12 {
                current_month = 1;
                current_year += 1;
            }
        }

        let n = poraba_nabava.len();
        assert!(n > 0);
        poraba_nabava[n - 1].2 = zaloga_sum;

        for i in 1..n {
            poraba_nabava[n - 1 - i].2 = poraba_nabava[n - i].2 + poraba_nabava[n - i].0 - poraba_nabava[n - i].1;
        }



        self.material = material;
        self.naziv = naziv.to_string();
        self.months = months;
        self.poraba_nabava = poraba_nabava;
    }
}