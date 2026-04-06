use std::collections::HashMap;
use std::path::PathBuf;
use calamine::{open_workbook_auto, DataType, Reader};
use rfd::{MessageDialog, MessageLevel};


#[derive(Default, Debug)]
pub struct RowData {
    pub material: i64,
    pub zaloga: f64,
    pub poraba_3m: f64,
    pub poraba_12m: f64,
    pub odprta_narocila: f64,
}
pub fn parse_import_files(files: Vec<PathBuf>) -> Result<Vec<RowData>, Box<dyn std::error::Error>> {
    if files.len() != 4 {
        MessageDialog::new()
            .set_title("Napaka, nepravilno število datotek != 4")
            .set_level(MessageLevel::Error)
            .show();
    }



    let poraba_file = files.iter().find(|&path_buf| {
        path_buf.file_name().unwrap() == "PORABA 3M.XLSX"
    }).ok_or("File PORABA 3M.XLSX not found")?;
    let mut workbook = open_workbook_auto(poraba_file)?;
    let range= workbook.worksheet_range(workbook.sheet_names().get(0).ok_or("Workbook has no sheets")?).unwrap();
    let mut poraba_3m_map: HashMap<i64, f64> = HashMap::new();
    for row in range.rows().skip(1) {
        let material = row.get(1)
            .and_then(DataType::get_string)
            .map(|f| f.parse::<i64>().unwrap_or(0))
            .unwrap_or(0);
        if material == 0 { continue; }

        let klc_v_em_vnosa = row.get(9)
            .and_then(DataType::get_float)
            .map(|f| f)
            .unwrap_or(0.);

        let entry = poraba_3m_map
            .entry(material)
            .or_insert(0.0);

        *entry += klc_v_em_vnosa;
    }



    let poraba_file = files.iter().find(|&path_buf| {
        path_buf.file_name().unwrap() == "PORABA 12M.XLSX"
    }).ok_or("File PORABA 12M.XLSX not found")?;
    let mut workbook = open_workbook_auto(poraba_file)?;
    let range= workbook.worksheet_range(workbook.sheet_names().get(0).ok_or("Workbook has no sheets")?).unwrap();
    let mut poraba_12m_map: HashMap<i64, f64> = HashMap::new();
    for row in range.rows().skip(1) {
        let material = row.get(1)
            .and_then(DataType::get_string)
            .map(|f| f.parse::<i64>().unwrap_or(0))
            .unwrap_or(0);
        if material == 0 { continue; }

        let klc_v_em_vnosa = row.get(9)
            .and_then(DataType::get_float)
            .map(|f| f)
            .unwrap_or(0.);

        let entry = poraba_12m_map
            .entry(material)
            .or_insert(0.0);

        *entry += klc_v_em_vnosa;
    }

    let odprta_narocila_file = files.iter().find(|&path_buf| {
        path_buf.file_name().unwrap() == "ODPRTA NAROČILA.XLSX"
    }).ok_or("File ODPRTA_NAROČILA.XLSX not found")?;
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
    }).ok_or("File ZALOGA.XLSX not found")?;
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


        let poraba_3m = poraba_3m_map.get(&material).unwrap_or(&0.).abs() / 3.;
        let poraba_12m = poraba_12m_map.get(&material).unwrap_or(&0.).abs() / 12.;
        let odprta_narocila = *dobava_map.get(&material).unwrap_or(&0.);

        row_data.push(RowData {
            material,
            zaloga,
            poraba_3m,
            poraba_12m,
            odprta_narocila
        });
    }


    Ok(row_data)
}




#[derive(Default)]
pub struct SifrantRow {
    pub material: i64,
    pub naziv_materiala: String,
    pub nabavna_skupina: String,
    pub mrp_karakteristika: String,
}
pub fn parse_sifrant_file(path: PathBuf) -> Result<Vec<SifrantRow>, Box<dyn std::error::Error>> {
    if !path.file_name().unwrap().eq("ŠIFRANT.XLSX") {
        Err("Bad filename!")?;
    }
    let mut row_data = Vec::new();
    let mut workbook = open_workbook_auto(path)?;
    let range= workbook.worksheet_range(workbook.sheet_names().get(0).ok_or("Workbook has no sheets")?).unwrap();
    for row in range.rows().skip(1) {
        let material = row.get(0)
            .and_then(DataType::get_string)
            .map(|f| f.parse::<i64>().unwrap_or(0))
            .unwrap_or(0);
        if material == 0 { continue; }

        let naziv_materiala = row.get(3)
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
            nabavna_skupina,
            mrp_karakteristika,
        });
    }

    Ok(row_data)
}


#[derive(Default)]
pub struct DobaviteljRow {
    pub material: i64,
    pub dobavitelj: String,
}
pub fn parse_dobavitelji_file(path: PathBuf) -> Result<Vec<DobaviteljRow>, Box<dyn std::error::Error>> {
    if !path.file_name().unwrap().eq("DOBAVITELJI.XLSX") {
        Err("Bad filename!")?;
    }
    let mut row_data = Vec::new();
    let mut workbook = open_workbook_auto(path)?;
    let range= workbook.worksheet_range(workbook.sheet_names().get(0).ok_or("Workbook has no sheets")?).unwrap();
    let mut dobavitelji_map: HashMap<i64, Vec<String>> = HashMap::new();

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


    Ok(row_data)
}