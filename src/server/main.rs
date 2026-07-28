#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]
#![allow(deprecated)]

mod db;
mod parse;

use std::path::PathBuf;
use std::sync::{Arc, Mutex};
use std::sync::mpsc::channel;
use std::thread;
use std::thread::JoinHandle;
use std::time::Duration;
use eframe::egui::{CentralPanel, Context, ScrollArea, Ui};
use eframe::{egui, Frame, NativeOptions};
use env_logger::Env;
use notify::{Config, EventKind, RecommendedWatcher, RecursiveMode, Watcher};
use rouille::{router, Response};
use tray_icon::{MouseButton, MouseButtonState, TrayIcon, TrayIconBuilder, TrayIconEvent};
use magneti_excel::{HandInput, SortState};
use crate::db::DBManager;
use crate::parse::{get_existing_files, parse_and_upload_all_files};

pub const FILE_NAMES: [&str; 7] = [
    "ŠIFRANT.XLSX",
    "DOBAVITELJI.XLSX",
    "ZALOGA100.XLSX",
    "PORABA.XLSX",
    "ODPRTA NAROČILA.XLSX",
    "ZALOGA.XLSX",
    "NABAVA.XLSX"
];


struct FolderWatcher {
    handle: Option<JoinHandle<()>>,
    db_manager: Arc<Mutex<DBManager>>,
}

impl FolderWatcher {
    pub fn start(&mut self, folder: PathBuf) {
        let db_manager_thread = Arc::clone(&self.db_manager);

        self.handle = Some(thread::spawn(move || {
            log::info!("starting watcher on: {:?}", folder.to_str());
            let (tx, rx) = channel();

            let mut watcher = match RecommendedWatcher::new(tx, Config::default()) {
                Ok(w) => w,
                Err(e) => {
                    log::error!("Failed to create a watcher: {:?}", e);
                    return;
                }
            };

            let _ = watcher.watch(&folder, RecursiveMode::NonRecursive).inspect_err(|err| {
                log::error!("{:?}", err);
            });

            for res in rx {
                match res {
                    Ok(event ) => {
                        let refresh = event.paths.iter().any(|p| {
                            FILE_NAMES.iter().any(|file| {
                                file.eq_ignore_ascii_case(p.file_name().and_then(|name| name.to_str()).unwrap_or(""))
                            })
                        });
                        if refresh {
                            match event.kind {
                                EventKind::Create(_)
                                //| EventKind::Modify(_)
                                => {
                                    let files = get_existing_files(folder.clone());
                                    let res = parse_and_upload_all_files(files, &db_manager_thread);
                                    match res {
                                        Ok(_) => { log::info!("Parse all files")},
                                        Err(e) => {log::error!("Failed to parse all files: {}", e)}
                                    }
                                },
                                _ => {},
                            }
                        }

                    },
                    Err(e) => {log::error!("Watcher error: {:?}", e)}
                }
            }
        }));
    }

    pub fn stop(&mut self) {
        self.handle.take().map(|h| h.join());
    }
}

struct NetworkController {
    handle: Option<JoinHandle<()>>,
}
impl NetworkController {
    pub fn handle(&mut self, db_manager: Arc<Mutex<DBManager>>) {
        self.handle = Some(thread::spawn(move || {
            let port = "0.0.0.0:8080";
            log::info!("Starting server on: http://{}", port);
            rouille::start_server(port, move |request| {
                router!(request,
                    (POST) (/upload) => {
                        log::info!("Handling /upload");
                        let hand_input: HandInput = match rouille::input::json_input(request) {
                            Ok(v) => v,
                            Err(e) => {
                                log::error!("Failed to parse HandInput: {:?}", e);
                                return Response::text("Invalid HandInput JSON").with_status_code(400);
                            }
                        };

                        match hand_input {
                            HandInput::DobavniRok(material, b) => {
                                if let Ok(db) = db_manager.lock() {
                                    return match db.store_dobavni_rok((material, b)) {
                                        Ok(_) => {
                                            log::info!("stored dobavni_rok");
                                            Response::text("Stored dobavni_rok")
                                        }
                                        Err(e) => {
                                            log::error!("Didnt store dobavni_rok {}", e);
                                            Response::text("Didn't store dobavni_rok")
                                        }
                                    };
                                }
                                Response::text("Failed to aquire db manager lock")
                            },
                            HandInput::Opomba(material, b) => {
                                if let Ok(db) = db_manager.lock() {
                                    return match db.store_opomba((material, b)) {
                                        Ok(_) => {
                                            log::info!("stored opomba");
                                            Response::text("Stored opomba")
                                        }
                                        Err(e) => {
                                            log::error!("Didnt store opomba {}", e);
                                            Response::text("Didn't store opomba")
                                        }
                                    };
                                }
                                Response::text("Failed to aquire db manager lock")
                            },
                            HandInput::MinZaloga(material, b) => {
                                if let Ok(db) = db_manager.lock() {
                                    return match db.store_min_zaloga((material, b)) {
                                        Ok(_) => {
                                            log::info!("stored min_zaloga");
                                            Response::text("Stored min_zaloga")
                                        }
                                        Err(e) => {
                                            log::error!("Didnt store min_zaloga {}", e);
                                            Response::text("Didn't store min_zaloga")
                                        }
                                    };
                                }
                                Response::text("Failed to aquire db manager lock")
                            },
                            HandInput::MaxZaloga(material, b) => {
                                if let Ok(db) = db_manager.lock() {
                                    return match db.store_max_zaloga((material, b)) {
                                        Ok(_) => {
                                            log::info!("stored max_zaloga");
                                            Response::text("Stored max_zaloga")
                                        }
                                        Err(e) => {
                                            log::error!("Didnt store max_zaloga {}", e);
                                            Response::text("Didn't store max_zaloga")
                                        }
                                    };
                                }
                                Response::text("Failed to aquire db manager lock")
                            },
                            HandInput::BlagovnaSkupina(material, b) => {
                                if let Ok(db) = db_manager.lock() {
                                    return match db.store_blagovna_skupina((material, b)) {
                                        Ok(_) => {
                                            log::info!("stored blagovna_skupina");
                                            Response::text("Stored blagovna_skupina")
                                        }
                                        Err(e) => {
                                            log::error!("Didnt store blagovna_skupina {}", e);
                                            Response::text("Didn't store blagovna_skupina")
                                        }
                                    };
                                }
                                Response::text("Failed to aquire db manager lock")
                            },
                            HandInput::Pakiranje(material, b) => {
                                if let Ok(db) = db_manager.lock() {
                                    return match db.store_pakiranje((material, b)) {
                                        Ok(_) => {
                                            log::info!("stored pakiranje");
                                            Response::text("Stored pakiranje")
                                        }
                                        Err(e) => {
                                            log::error!("Didnt store pakiranje {}", e);
                                            Response::text("Didn't store pakiranje")
                                        }
                                    };
                                }
                                Response::text("Failed to aquire db manager lock")
                            }
                        }
                    },
                (GET) (/data/{_start: usize}/{_stop: usize}/{sort_state: SortState}) => {
                        log::info!("Handling /data");
                    if let Ok(db) = db_manager.lock() {
                        let res = db.get_data(&sort_state);
                        match res {
                            Ok(data) => {
                                Response::json(&data)
                            },
                            Err(e) => {
                                log::error!("{:?}", e);
                                Response::text("Bad db.get_data()")
                            },
                        }
                    } else {
                        Response::text("Couldn't lock db_manager, try again later!")
                    }
                },
                    (GET) (/poraba/{material: i64}) => {
                        log::info!("Handling /poraba");
                        if let Ok(db) = db_manager.lock() {
                            let res = db.get_poraba(material);
                            match res {
                                Ok(data) => {
                                    Response::json(&data)
                                },
                                Err(e) => {
                                    log::error!("{:?}", e);
                                    Response::text("Bad db.get_poraba()")
                                },
                            }
                        } else {
                            Response::text("Couldn't lock db_manager, try again later!")
                        }
                    },

                    (GET) (/nabava/{material: i64}) => {
                        log::info!("handling /nabava");
                        if let Ok(db) = db_manager.lock() {
                            let res = db.get_nabava(material);
                            match res {
                                Ok(data) => {
                                    Response::json(&data)
                                },
                                Err(e) => {
                                    log::error!("{:?}", e);
                                    Response::text("Bad db.get_nabava()")
                                },
                            }
                        } else {
                            Response::text("Couldn't lock db_manager, try again later!")
                        }
                    },
                _ => {
                        log::info!("Handling bad link");
                        Response::text("bad link")
                    },
            )
            });
        }));
    }
}


struct App {
    _tray_icon: TrayIcon,
    folder_watcher: FolderWatcher,
}

impl App {
    pub fn new<'a>(_cc: &'a eframe::CreationContext<'a>) -> Self {
        let db_manager = Arc::new(Mutex::new(DBManager::default()));
        let mut network_controller = NetworkController {
            handle: None,
        };
        network_controller.handle(Arc::clone(&db_manager));


        let icon_data = vec![255, 0, 0, 255].repeat(32 * 32);
        let icon = tray_icon::Icon::from_rgba(icon_data, 32, 32).unwrap();

        let _tray_icon = TrayIconBuilder::new()
            .with_tooltip("Magneti Excel")
            .with_icon(icon)
            .build()
            .unwrap();

        Self {
            _tray_icon,
            folder_watcher: FolderWatcher {
                handle: None,
                db_manager: Arc::clone(&db_manager),
            },
        }
    }


    pub fn render(&mut self, ui: &mut Ui) {
        if ui.button("Opazuj folder").clicked() {
            if let Some(path) = rfd::FileDialog::new().pick_folder() {
                self.folder_watcher.stop();
                log::info!("{}", format!("{:?}", path));
                self.folder_watcher.start(path);

            }
        }
    }
}

impl eframe::App for App {
    fn update(&mut self, ctx: &Context, _frame: &mut Frame) {
        ctx.request_repaint_after(Duration::from_millis(100));

        ctx.input(|i| {
            if i.raw.viewport().minimized == Some(true) {
                ctx.send_viewport_cmd(egui::ViewportCommand::Minimized(false));
            }
        });

        if let Ok(event) = TrayIconEvent::receiver().try_recv() {
            if let TrayIconEvent::Click {
                button: MouseButton::Left,
                button_state: MouseButtonState::Up,
                ..
            } = event
            {
                ctx.send_viewport_cmd(egui::ViewportCommand::Maximized(true));
                ctx.send_viewport_cmd(egui::ViewportCommand::Focus);
                ctx.request_repaint();
            }
        }

        CentralPanel::default().show(ctx, |ui| {
            ScrollArea::vertical().show(ui, |ui| {
                self.render(ui);
            });
        });
    }
}

fn main() {
    let debug = true;

    let level = if debug { "info" } else { "warn" };

    env_logger::Builder::from_env(
        Env::default().default_filter_or(level)
    )
        .init();

    log::info!("Server started");


    eframe::run_native(
        "Magneti Strežnik",
        NativeOptions::default(),
        Box::new(|cc| Ok(Box::new(App::new(cc))))
    ).unwrap();
}