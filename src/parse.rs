use std::collections::{HashMap, HashSet};
use std::ffi::{OsString};
use std::path::PathBuf;
use calamine::{open_workbook_auto, DataType, Reader};
use chrono::{Local, Months, NaiveDate};
use crate::db::DBManager;

pub fn parse_all_files(files: Vec<PathBuf>, db_manager: &DBManager) -> Result<(), Box<dyn std::error::Error>> {
    
    let sifrant_file = files.iter().filter(|file| {
        match file.file_name().unwrap().to_ascii_uppercase().to_str().unwrap() {
            "ŠIFRANT.XLSX" => true,
            _ => false,
        }
    }).cloned().take(1).collect::<PathBuf>();
    log::info!("Started parsing sifrant file");
    let _ = parse_sifrant_file(sifrant_file).and_then(|rows| {
        log::info!("Finished parsing sifrant file");
        db_manager.store_sifrant_to_db(rows)
    }).inspect_err(|e| log::error!("{}", e.to_string()));

    let dobavitelji_file = files.iter().filter(|file| {
        match file.file_name().unwrap().to_ascii_uppercase().to_str().unwrap() {
            "DOBAVITELJI.XLSX" => true,
            _ => false,
        }
    }).cloned().take(1).collect::<PathBuf>();
    log::info!("Started parsing dobavitelji file");
    let _ = parse_dobavitelji_file(dobavitelji_file).and_then(|rows| {
        log::info!("Finished parsing dobavitelji file");
        db_manager.store_dobavitelji_to_db(rows)
    }).inspect_err(|e| log::error!("{}", e.to_string()));

    let zaloga100_file = files.iter().filter(|file| {
        match file.file_name().unwrap().to_ascii_uppercase().to_str().unwrap() {
            "ZALOGA100.XLSX" => true,
            _ => false,
        }
    }).cloned().take(1).collect::<PathBuf>();
    log::info!("Started parsing zaloga100 file");
    let _ = parse_razpolozljiva_zaloga_file(zaloga100_file).and_then(|rows| {
        log::info!("Finished parsing zaloga100 file");
        db_manager.store_razpolozljive_zaloge_to_db(rows)
    }).inspect_err(|e| log::error!("{}", e.to_string()));


    let import_files = files.iter().filter(|file| {
        match file.file_name().unwrap().to_ascii_uppercase().to_str().unwrap() {
            "PORABA.XLSX" | "ODPRTA NAROČILA.XLSX" | "ZALOGA.XLSX" => true,
            _ => false,
        }
    }).cloned().collect::<Vec<PathBuf>>();
    log::info!("Started parsing 3 files");
    let _ = parse_import_files(import_files).and_then(|row_data| {
        log::info!("Finished parsing 3 files");
        db_manager.store_to_data(row_data)
    }).inspect_err(|e| log::error!("{}", e.to_string()));



    let poraba_file = files.iter().filter(|file| {
        match file.file_name().unwrap().to_ascii_uppercase().to_str().unwrap() {
            "PORABA.XLSX" => true,
            _ => false,
        }
    }).cloned().take(1).collect::<PathBuf>();
    log::info!("Started parsing poraba file");
    let _ = parse_poraba_file(poraba_file).and_then(|row_data| {
        log::info!("Finished parsing poraba file");
        db_manager.store_poraba_to_db(row_data)
    }).inspect_err(|e| log::error!("{}", e.to_string()));

    let prevzemi_file = files.iter().filter(|file| {
        match file.file_name().unwrap().to_ascii_uppercase().to_str().unwrap() {
            "NABAVA.XLSX" => true,
            _ => false,
        }
    }).cloned().take(1).collect::<PathBuf>();
    log::info!("Started parsing prevzemi file");
    let _ = parse_nabava_file(prevzemi_file).and_then(|row_data| {
        log::info!("Finished parsing prevzemi file");
        db_manager.store_nabava_to_db(row_data)
    }).inspect_err(|e| log::error!("{}", e.to_string()));
    
    db_manager.try_create_view()?;


    log::info!("Finished parsing ALL files");
    Ok(())
}


pub struct NabavaData {
    pub material: i64,
    pub nabava: f64,
    pub date: NaiveDate,
}

pub fn parse_nabava_file(file: PathBuf) -> Result<Vec<NabavaData>, Box<dyn std::error::Error>> {
    if !file.file_name().unwrap_or(OsString::default().as_os_str()).eq("NABAVA.XLSX") {
        Err("Bad filename!")?;
    }
    let mut workbook = open_workbook_auto(file)?;
    let range= workbook.worksheet_range(workbook.sheet_names().get(0).ok_or("Workbook has no sheets")?).unwrap();
    let mut nabava_rows: Vec<NabavaData> = Vec::with_capacity(1000);
    log::info!("Started parsing nabava");
    for row in range.rows().skip(1) {
        let material = row.get(1)
            .and_then(DataType::get_string)
            .map(|f| f.parse::<i64>().unwrap_or(0))
            .unwrap_or(0);
        if material == 0 { continue; }


        let date = row.get(8)
            .and_then(DataType::get_datetime)
            .map(|date| {
                let  (y, m, d, _, _, _, _) = date.to_ymd_hms_milli();
                NaiveDate::from_ymd_opt(y as i32, m as u32, d as u32).unwrap_or(NaiveDate::default())
            })
            .unwrap_or(NaiveDate::default());

        let nabava = row.get(9)
            .and_then(DataType::get_float)
            .map(|f| f)
            .unwrap_or(0.);

        nabava_rows.push(NabavaData {
            material,
            nabava,
            date,
        })
    }
    log::info!("Parsed nabava: {}", range.rows().len());

    Ok(nabava_rows)
}

pub struct PorabaData {
    pub material: i64,
    pub poraba: f64,
    pub date: NaiveDate,
}

pub fn parse_poraba_file(file: PathBuf) -> Result<Vec<PorabaData>, Box<dyn std::error::Error>> {
    if !file.file_name().unwrap_or(OsString::default().as_os_str()).eq("PORABA.XLSX") {
        Err("Bad filename!")?;
    }
    let mut workbook = open_workbook_auto(file)?;
    let range= workbook.worksheet_range(workbook.sheet_names().get(0).ok_or("Workbook has no sheets")?).unwrap();
    let mut poraba_rows: Vec<PorabaData> = Vec::with_capacity(1000);
    log::info!("Started parsing poraba");
    for row in range.rows().skip(1) {
        let material = row.get(1)
            .and_then(DataType::get_string)
            .map(|f| f.parse::<i64>().unwrap_or(0))
            .unwrap_or(0);
        if material == 0 { continue; }


        let date = row.get(8)
            .and_then(DataType::get_datetime)
            .map(|date| {
                let  (y, m, d, _, _, _, _) = date.to_ymd_hms_milli();
                NaiveDate::from_ymd_opt(y as i32, m as u32, d as u32).unwrap_or(NaiveDate::default())
            })
            .unwrap_or(NaiveDate::default());
        
        let poraba = row.get(9)
            .and_then(DataType::get_float)
            .map(|f| f.abs())
            .unwrap_or(0.);

        poraba_rows.push(PorabaData {
            material,
            poraba,
            date
        })
    }
    log::info!("Parsed poraba: {}", range.rows().len());


    Ok(poraba_rows)
}


#[derive(Default, Debug)]
pub struct RowData {
    pub material: i64,
    pub zaloga: f64,
    pub poraba_3m: f64,
    pub poraba_24m: f64,
    pub odprta_narocila: f64,
}
pub fn parse_import_files(files: Vec<PathBuf>) -> Result<Vec<RowData>, Box<dyn std::error::Error>> {
    let poraba_file = files.iter().find(|&path_buf| {
        path_buf.file_name().unwrap() == "PORABA.XLSX"
    }).ok_or("File PORABA.XLSX not found")?;
    let mut workbook = open_workbook_auto(poraba_file)?;
    let range= workbook.worksheet_range(workbook.sheet_names().get(0).ok_or("Workbook has no sheets")?).unwrap();
    let mut poraba_map: HashMap<i64, Vec<(f64, NaiveDate)>> = HashMap::new();
    log::info!("Started parsing poraba");
    for row in range.rows().skip(1) {
        let material = row.get(1)
            .and_then(DataType::get_string)
            .map(|f| f.parse::<i64>().unwrap_or(0))
            .unwrap_or(0);
        if material == 0 { continue; }


        let datum = row.get(8)
            .and_then(DataType::get_datetime)
            .map(|date| {
                let  (y, m, d, _, _, _, _) = date.to_ymd_hms_milli();
                NaiveDate::from_ymd_opt(y as i32, m as u32, d as u32).unwrap_or(NaiveDate::default())
            })
            .unwrap_or(NaiveDate::default());



        let klc_v_em_vnosa = row.get(9)
            .and_then(DataType::get_float)
            .map(|f| f)
            .unwrap_or(0.);

        let entry = poraba_map
            .entry(material)
            .or_default();
        entry.push((klc_v_em_vnosa, datum));
    }
    log::info!("Parsed poraba: {}", range.rows().len());

    let odprta_narocila_file = files.iter().find(|&path_buf| {
        path_buf.file_name().unwrap() == "ODPRTA NAROČILA.XLSX"
    }).ok_or("File ODPRTA_NAROČILA.XLSX not found")?;
    let mut workbook = open_workbook_auto(odprta_narocila_file)?;
    let range= workbook.worksheet_range(workbook.sheet_names().get(0).ok_or("Workbook has no sheets")?).unwrap();
    let mut dobava_map: HashMap<i64, f64> = HashMap::new();
    log::info!("Started parsing odprta naročila");
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
    log::info!("Parsed odprta naročila: {}", range.rows().len());


    let zaloga_file = files.iter().find(|&path_buf| {
        path_buf.file_name().unwrap() == "ZALOGA.XLSX"
    }).ok_or("File ZALOGA.XLSX not found")?;
    let mut workbook = open_workbook_auto(zaloga_file)?;
    let range= workbook.worksheet_range(workbook.sheet_names().get(0).ok_or("Workbook has no sheets")?).unwrap();
    let mut row_data: Vec<RowData> = Vec::with_capacity(100);
    let mut zaloga_map: HashMap<i64, f64> = HashMap::new();
    log::info!("Started parsing zaloga");
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

        let entry = zaloga_map
            .entry(material)
            .or_insert(0.);
        *entry += zaloga;
    }
    log::info!("Parsed zaloga: {}", range.rows().len());


    log::info!("Parsing for data table");
    let key_set: HashSet<i64> = poraba_map.keys()
        .chain(dobava_map.keys())
        .chain(zaloga_map.keys())
        .copied().collect();

    key_set.iter().for_each(|&material| {
        let poraba_3m = poraba_map.get(&material)
            .map(|vec| {
                vec.iter()
                    .filter(|(_, date)| is_within_last_months(&date, 3))
                    .map(|(val, _)| val.abs())
                    .sum::<f64>() / 3.
            }).unwrap_or(0.);

        let poraba_24m = poraba_map.get(&material)
            .map(|vec| {
                vec.iter()
                    .filter(|(_, date)| is_within_last_months(&date, 24))
                    .map(|(val, _)| val.abs())
                    .sum::<f64>() / 24.
            }).unwrap_or(0.);

        
        
        let odprta_narocila = *match dobava_map.get(&material) {
            Some(val) => {
                val
            }
            None => &0.
        };

        let zaloga = *match zaloga_map.get(&material) {
            Some(val) => {
                val
            }
            None => &0.
        };


        row_data.push(RowData {
            material,
            zaloga,
            poraba_3m,
            poraba_24m,
            odprta_narocila
        });
    });

    log::info!("Finished parsing import files");

    Ok(row_data)
}


fn is_within_last_months(date: &NaiveDate, months: u32) -> bool {
    let today = Local::now().date_naive();
    let cutoff = today.checked_sub_months(Months::new(months)).unwrap();

    date >= &cutoff && date <= &today
}

#[derive(Default)]
pub struct SifrantRow {
    pub material: i64,
    pub naziv_materiala: String,
    pub osnovna_merska_enota: String,
    pub nabavna_skupina: String,
    pub mrp_karakteristika: String,
}
pub fn parse_sifrant_file(path: PathBuf) -> Result<Vec<SifrantRow>, Box<dyn std::error::Error>> {
    if !path.file_name().unwrap_or(OsString::default().as_os_str()).eq("ŠIFRANT.XLSX") {
        Err("Bad filename!")?;
    }
    let mut row_data = Vec::new();
    let mut workbook = open_workbook_auto(path)?;
    let range= workbook.worksheet_range(workbook.sheet_names().get(0).ok_or("Workbook has no sheets")?).unwrap();
    log::info!("Started parsing sifrant");
    for row in range.rows().skip(1) {
        let material = row.get(0)
            .and_then(DataType::get_string)
            .map(|f| f.parse::<i64>().unwrap_or(0))
            .unwrap_or(0);
        if material == 0 { continue; }



        let naziv_materiala = row.get(3)
            .and_then(DataType::get_string)
            .unwrap_or("").to_string();

        let osnovna_merska_enota = row.get(7)
            .and_then(DataType::get_string)
            .unwrap_or("").to_string();

        let nabavna_skupina = row.get(8)
            .and_then(DataType::get_string)
            .unwrap_or("").to_string();

        let mrp_karakteristika = row.get(10)
            .and_then(DataType::get_string)
            .unwrap_or("").to_string();

        row_data.push(SifrantRow {
            material,
            naziv_materiala,
            osnovna_merska_enota,
            nabavna_skupina,
            mrp_karakteristika,
        });
    }
    log::info!("Parsed sifrant: {}", range.rows().len());

    Ok(row_data)
}


#[derive(Default)]
pub struct DobaviteljRow {
    pub material: i64,
    pub dobavitelj: String,
}
pub fn parse_dobavitelji_file(path: PathBuf) -> Result<Vec<DobaviteljRow>, Box<dyn std::error::Error>> {
    if !path.file_name().unwrap_or(OsString::default().as_os_str()).eq("DOBAVITELJI.XLSX") {
        Err("Bad filename!")?;
    }
    let mut row_data = Vec::new();
    let mut workbook = open_workbook_auto(path)?;
    let range= workbook.worksheet_range(workbook.sheet_names().get(0).ok_or("Workbook has no sheets")?).unwrap();
    let mut dobavitelji_map: HashMap<i64, Vec<String>> = HashMap::new();
    log::info!("Started parsing dobavitelji");
    for row in range.rows().skip(1) {
        let material = row.get(0)
            .and_then(DataType::get_string)
            .map(|f| f.parse::<i64>().unwrap_or(0))
            .unwrap_or(0);
        if material == 0 { continue; }

        let dobavitelj = row.get(7)
            .and_then(DataType::get_string)
            .unwrap_or("").to_string();

        let mut dont_add = false;

        if dobavitelji_map.contains_key(&material) {
           dont_add = dobavitelji_map.get(&material).as_ref().unwrap().iter().any(|dob| {
              dob.to_lowercase().eq(dobavitelj.to_lowercase().as_str())
           });
        }

        if !dont_add {
            dobavitelji_map.entry(material).or_insert_with(|| Vec::new()).push(dobavitelj);
        }
    }

    dobavitelji_map.into_iter().for_each(|(material, dobavitelji)| {
        row_data.push(DobaviteljRow {
            material,
            dobavitelj: dobavitelji.join(", "),
        });
    });

    log::info!("Parsed dobavitelji: {}", range.rows().len());
    Ok(row_data)
}



#[derive(Default)]
pub struct RazpolozljivaZalogaRow {
    pub material: i64,
    pub razpolozljiva_zaloga: f64,
}
pub fn parse_razpolozljiva_zaloga_file(path: PathBuf) -> Result<Vec<RazpolozljivaZalogaRow>, Box<dyn std::error::Error>> {
    if !path.file_name().unwrap_or(OsString::default().as_os_str()).eq("ZALOGA100.XLSX") {
        Err("Bad filename!")?;
    }
    let mut row_data = Vec::new();
    let mut workbook = open_workbook_auto(path)?;
    let range= workbook.worksheet_range(workbook.sheet_names().get(0).ok_or("Workbook has no sheets")?).unwrap();
    let mut zaloga_100_map: HashMap<i64, f64> = HashMap::new();
    log::info!("Started parsing zaloga100");
    for row in range.rows().skip(1) {
        let material = row.get(0)
            .and_then(DataType::get_string)
            .map(|f| f.parse::<i64>().unwrap_or(0))
            .unwrap_or(0);
        if material == 0 { continue; }

        let razpolozljiva_zaloga = row.get(10)
            .and_then(DataType::get_float)
            .unwrap_or(0.);


        let entry = zaloga_100_map
            .entry(material)
            .or_insert(0.);
        *entry += razpolozljiva_zaloga;

    }

    zaloga_100_map.into_iter().for_each(|(material, razpolozljiva_zaloga)| {
       row_data.push(RazpolozljivaZalogaRow {
           material,
           razpolozljiva_zaloga,
       })
    });

    log::info!("Parsed zaloga100: {}", range.rows().len());

    Ok(row_data)
}