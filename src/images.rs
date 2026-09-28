use std::path::{Path, PathBuf};
use eframe::egui::{pos2, vec2, Align2, Area, Button, Color32, Context, CornerRadius, CursorIcon, FontId, Id, Image, Key, Order, Rect, Sense, Ui};

static IMAGES_DIR: &str = "images";
static EXTENSIONS: [&str; 5] = ["jpg", "jpeg", "png", "bmp", "webp"];

// every image gets fit into this box, rest stays gray
static IMAGE_WIDTH: f32 = 500.0;
static IMAGE_HEIGHT: f32 = 400.0;

#[derive(Default)]
pub struct ImageCarousel {
    material: i64,
    naziv: String,
    images: Vec<PathBuf>,
    index: usize,
    open: bool,
    just_opened: bool,
    fullscreen: bool,
}

impl ImageCarousel {
    pub fn clear(&mut self, ctx: &Context) {
        self.material = 0;
        self.naziv = String::new();
        self.images = Vec::new();
        self.index = 0;
        self.open = false;
        self.just_opened = false;
        self.fullscreen = false;

        // so new/changed images in the folder get reloaded next time
        ctx.forget_all_images();
    }

    pub fn is_open(&self) -> bool {
        self.open
    }

    fn next(&mut self) {
        self.index = (self.index + 1) % self.images.len();
    }

    fn previous(&mut self) {
        self.index = (self.index + self.images.len() - 1) % self.images.len();
    }

    // returns true if popup should stay open (pointer over it, it was opened this frame or is fullscreen)
    pub fn render(&mut self, ui: &mut Ui) -> bool {
        if !self.open {
            return false;
        }

        // press and release of the naziv click can land in the same frame
        let just_opened = std::mem::take(&mut self.just_opened);
        let was_fullscreen = self.fullscreen;

        if ui.input(|i| i.key_pressed(Key::Escape)) {
            if self.fullscreen {
                self.fullscreen = false;
            } else {
                self.clear(ui.ctx());
                return false;
            }
        }

        let (rect, _) = ui.allocate_exact_size(vec2(IMAGE_WIDTH + 40.0, IMAGE_HEIGHT + 40.0 + 60.0 + 50.0), Sense::click());

        let title_rect = Rect::from_min_max(
            rect.min,
            pos2(rect.right(), rect.top() + 60.0),
        );

        let image_rect = Rect::from_min_max(
            pos2(rect.left(), title_rect.bottom()),
            pos2(rect.right(), rect.bottom() - 50.0),
        );

        let bottom_rect = Rect::from_min_max(
            pos2(rect.left(), image_rect.bottom()),
            rect.max,
        );

        ui.painter().rect_filled(title_rect, CornerRadius::same(0), Color32::WHITE);
        ui.painter().rect_filled(image_rect, CornerRadius::same(0), Color32::from_gray(240));
        ui.painter().rect_filled(bottom_rect, CornerRadius::same(0), Color32::WHITE);

        ui.painter().text(
            title_rect.center(),
            Align2::CENTER_CENTER,
            &format!("Slike - {}, {}", self.material, self.naziv),
            FontId::proportional(20.0),
            Color32::BLACK,
        );

        if self.images.is_empty() {
            ui.painter().text(
                image_rect.center(),
                Align2::CENTER_CENTER,
                &format!("Ni slik za material {}", self.material),
                FontId::proportional(20.0),
                Color32::BLACK,
            );

            return ui.rect_contains_pointer(rect) || just_opened;
        }

        if ui.input(|i| i.key_pressed(Key::ArrowLeft)) {
            self.previous();
        }
        if ui.input(|i| i.key_pressed(Key::ArrowRight)) {
            self.next();
        }

        let path = self.images[self.index].clone();
        let uri = format!("file://{}", path.display());

        let image_box = Rect::from_center_size(image_rect.center(), vec2(IMAGE_WIDTH, IMAGE_HEIGHT));
        let image_response = ui.put(
            image_box,
            Image::new(&uri)
                .fit_to_exact_size(image_box.size())
                .maintain_aspect_ratio(true)
                .sense(Sense::click()),
        ).on_hover_cursor(CursorIcon::ZoomIn);

        if image_response.clicked() {
            self.fullscreen = true;
        }

        ui.painter().text(
            pos2(bottom_rect.left() + 20.0, bottom_rect.center().y),
            Align2::LEFT_CENTER,
            path.file_name().and_then(|n| n.to_str()).unwrap_or(""),
            FontId::proportional(12.0),
            Color32::DARK_GRAY,
        );

        self.render_counter(ui, bottom_rect, Color32::BLACK);

        if self.fullscreen {
            self.render_fullscreen(ui.ctx(), &uri);
        }

        ui.rect_contains_pointer(rect) || just_opened || was_fullscreen || self.fullscreen
    }

    fn render_counter(&mut self, ui: &mut Ui, rect: Rect, color: Color32) {
        ui.painter().text(
            rect.center(),
            Align2::CENTER_CENTER,
            &format!("{} / {}", self.index + 1, self.images.len()),
            FontId::proportional(20.0),
            color,
        );

        let previous_rect = Rect::from_center_size(pos2(rect.center().x - 90.0, rect.center().y), vec2(50.0, 30.0));
        let next_rect = Rect::from_center_size(pos2(rect.center().x + 90.0, rect.center().y), vec2(50.0, 30.0));

        if ui.put(previous_rect, Button::new("<")).clicked() {
            self.previous();
        }
        if ui.put(next_rect, Button::new(">")).clicked() {
            self.next();
        }
    }

    fn render_fullscreen(&mut self, ctx: &Context, uri: &str) {
        let screen = ctx.content_rect();

        Area::new(Id::new("images_fullscreen"))
            .order(Order::Foreground)
            .fixed_pos(screen.min)
            .show(ctx, |ui| {
                let (rect, response) = ui.allocate_exact_size(screen.size(), Sense::click());
                ui.painter().rect_filled(rect, CornerRadius::same(0), Color32::from_black_alpha(230));

                let image_rect = Rect::from_min_max(
                    rect.min + vec2(20.0, 20.0),
                    pos2(rect.right() - 20.0, rect.bottom() - 60.0),
                );

                let bottom_rect = Rect::from_min_max(
                    pos2(rect.left(), image_rect.bottom()),
                    rect.max,
                );

                // image doesn't sense clicks, so click on it goes to the background and closes fullscreen
                ui.put(
                    image_rect,
                    Image::new(uri)
                        .fit_to_exact_size(image_rect.size())
                        .maintain_aspect_ratio(true),
                );

                self.render_counter(ui, bottom_rect, Color32::WHITE);

                if response.on_hover_cursor(CursorIcon::ZoomOut).clicked() {
                    self.fullscreen = false;
                }
            });
    }

    pub fn query(&mut self, material: i64, naziv: &str) {
        let mut images: Vec<PathBuf> = match std::fs::read_dir(IMAGES_DIR) {
            Ok(entries) => {
                entries
                    .filter_map(|entry| entry.ok())
                    .map(|entry| entry.path())
                    .filter(|path| is_image_of_material(path, material))
                    .collect()
            }
            Err(err) => {
                log::error!("images dir error: {:?}", err);
                Vec::new()
            }
        };

        // 343, 343 (1), 343 (2), ... 343 (10)
        images.sort_by_key(|path| {
            let stem = path.file_stem().and_then(|s| s.to_str()).unwrap_or("").to_string();
            (stem.len(), stem)
        });

        log::info!("images for {}: {}", material, images.len());

        self.material = material;
        self.naziv = naziv.to_string();
        self.images = images;
        self.index = 0;
        self.open = true;
        self.just_opened = true;
        self.fullscreen = false;
    }
}

fn is_image_of_material(path: &Path, material: i64) -> bool {
    let extension = path.extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase())
        .unwrap_or(String::new());

    if !EXTENSIONS.contains(&extension.as_str()) {
        return false;
    }

    // "343", "343 (1)", "343(2)", "343 - Copy" -> 343, but "3430" -> 3430
    let stem = path.file_stem().and_then(|s| s.to_str()).unwrap_or("");
    let code: String = stem.chars().take_while(|c| c.is_ascii_digit()).collect();

    code.parse::<i64>().is_ok_and(|code| code == material)
}
