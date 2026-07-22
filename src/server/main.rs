//#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]
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
use eframe::{Frame, NativeOptions};
use env_logger::Env;
use notify::{Config, EventKind, RecommendedWatcher, RecursiveMode, Watcher};
use rouille::{router, Response};
use magneti_excel::SortState;
use crate::db::DBManager;
use crate::parse::{get_existing_files, parse_all_files};

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
                                    let res = parse_all_files(files, &db_manager_thread);
                                    match res {
                                        Ok(_) => { log::info!("Parse all files")},
                                        Err(e) => {log::error!("Failed to parse all files")}
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
    db_manager: Arc<Mutex<DBManager>>,
}
impl NetworkController {
    pub fn handle(&mut self) {
        let port = "0.0.0.0:8080";
        log::info!("Starting server on: http://{}", port);
        /*
        rouille::start_server(port, move |request| {

            router!(request,
                (GET) (/data/{start: usize}/{stop: usize}/{sort_state: SortState}) => {
                    if let Ok(db) = self.db_manager.lock() {
                        db.
                    }

                    Response::json(&)
                },
                _ => Response::text("Endpoint not found, try /data/0/10/...").with_status_code(404)
            )
        });
        
         */
    }
}


struct App {
    folder_watcher: FolderWatcher,
    network_controller: NetworkController,
}

impl App {
    pub fn new<'a>(cc: &'a eframe::CreationContext<'a>) -> Self {
        let db_manager = Arc::new(Mutex::new(DBManager::default()));
        let network_controller = NetworkController {
            db_manager: Arc::clone(&db_manager),
        };
        Self {
            folder_watcher: FolderWatcher {
                handle: None,
                db_manager: Arc::clone(&db_manager),
            },
            network_controller,
        }
    }


    pub fn render(&mut self, ui: &mut Ui) {
        if ui.button("Opazuj folder").clicked() {
            if let Some(path) = rfd::FileDialog::new().pick_folder() {
                self.folder_watcher.stop();
                log::info!("{}", format!("{:?}", path));
                self.folder_watcher.start(path);

                self.network_controller.handle(); // FOREVER BLOCKS THREAD
            }
        }
    }
}

impl eframe::App for App {
    fn update(&mut self, ctx: &Context, frame: &mut Frame) {
        ctx.request_repaint_after(Duration::from_millis(100));

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
        "Magneti Excel",
        NativeOptions::default(),
        Box::new(|cc| Ok(Box::new(App::new(cc))))
    ).unwrap();
}