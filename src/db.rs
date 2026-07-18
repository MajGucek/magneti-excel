use chrono::NaiveDate;
use eframe::egui::{Color32, CornerRadius, CursorIcon, RichText};
use rfd::{MessageButtons, MessageDialog, MessageDialogResult, MessageLevel};
use sqlite::{Connection, State};
use serde::{Deserialize, Serialize};
use crate::{format_nabavnik, format_number_custom, parse_string_to_optional_f64, Rows, INDIGO, RED, TEAL};
use crate::graph::PorabaNabavaRows;
use crate::parse::{DobaviteljRow, NabavaData, PorabaData, RowData, SifrantRow, RazpolozljivaZalogaRow};

pub struct DBManager {
    pub db_name: String
}

impl DBManager {
    pub fn try_create_tables(&self) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        self.create_sifrant_table(&connection)?;
        self.create_data_table(&connection)?;
        self.create_poraba_table(&connection)?;
        self.create_nabava_table(&connection)?;
        self.create_dobavitelji_table(&connection)?;
        self.create_razpolozljive_zaloge_table(&connection)?;

        self.create_dobavni_rok_table(&connection)?;
        self.create_opomba_table(&connection)?;
        self.create_min_zaloga_table(&connection)?;
        self.create_max_zaloga_table(&connection)?;
        self.create_blagovna_skupina_table(&connection)?;
        self.create_pakiranje_table(&connection)?;

        Ok(())
    }


    pub fn get_data(&self, sort: &SortState) -> Result<Vec<ViewQuery>, Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        ViewQuery::query(&connection, &sort)
    }

    pub fn get_poraba(&self, material: i64) -> Result<Vec<PorabaQuery>, Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        PorabaQuery::query(material, &connection)
    }

    pub fn get_nabava(&self, material: i64) -> Result<Vec<NabavaQuery>, Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        NabavaQuery::query(material, &connection)
    }

    fn create_data_table(&self, connection: &Connection) -> Result<(), Box<dyn std::error::Error>> {
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("
            CREATE TABLE IF NOT EXISTS data (
                id INTEGER PRIMARY KEY,
                material INTEGER,
                zaloga REAL,
                poraba_3m REAL,
                poraba_24m REAL,
                odprta_narocila REAL,
                trenutna_zaloga_zadostuje_za_mesecev REAL
                    GENERATED ALWAYS AS (
                        CASE
                            WHEN poraba_3m <= 0 OR poraba_3m IS NULL THEN
                            CASE
                                WHEN poraba_24m <= 0 OR poraba_24m IS NULL THEN NULL
                                ELSE zaloga / poraba_24m
                            END
                            ELSE zaloga / poraba_3m
                        END
                    ) VIRTUAL,
                trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev REAL
                    GENERATED ALWAYS AS (
                        CASE
                            WHEN poraba_3m <= 0 OR poraba_3m IS NULL THEN
                            CASE
                                WHEN poraba_24m <= 0 OR poraba_24m IS NULL THEN NULL
                                ELSE (zaloga + odprta_narocila) / poraba_24m
                            END
                            ELSE (zaloga + odprta_narocila) / poraba_3m
                        END
                    ) VIRTUAL
            );
        ")?;
        connection.execute("COMMIT")?;
        log::info!("Created Data Table");
        Ok(())
    }


    pub fn drop_sifrant(&self) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("DROP TABLE sifrant;")?;
        connection.execute("COMMIT")?;
        log::info!("Dropped sifrant");

        Ok(())
    }

    pub fn drop_data(&self) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("DROP TABLE data;")?;
        connection.execute("COMMIT")?;
        log::info!("Dropped data");

        Ok(())
    }

    pub fn drop_dobavitelji(&self) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("DROP TABLE dobavitelji;")?;
        connection.execute("COMMIT")?;
        log::info!("Dropped dobavitelji");

        Ok(())
    }

    pub fn drop_razpolozljive_zaloge(&self) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("DROP TABLE razpolozljive_zaloge;")?;
        connection.execute("COMMIT")?;
        log::info!("Dropped razpolozljive_zaloge");

        Ok(())
    }

    pub fn drop_porabe(&self) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("DROP TABLE porabe;")?;
        connection.execute("COMMIT")?;
        log::info!("Dropped porabe");

        Ok(())
    }

    pub fn drop_nabave(&self) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("DROP TABLE nabave;")?;
        connection.execute("COMMIT")?;
        log::info!("Dropped nabave");

        Ok(())
    }



    pub fn store_to_data(&self, row_data: Vec<RowData>) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        let _ = self.drop_data();
        self.create_data_table(&connection)?;

        let mut statement = connection.prepare("
            INSERT INTO data (material, zaloga, poraba_3m, poraba_24m, odprta_narocila) VALUES (?, ?, ?, ?, ?)
        ")?;
        connection.execute("BEGIN TRANSACTION")?;
        log::info!("Started inserting into data");
        for row in row_data.iter() {
            statement.bind((1, row.material))?;
            statement.bind((2, row.zaloga))?;
            statement.bind((3, row.poraba_3m))?;
            statement.bind((4, row.poraba_24m))?;
            statement.bind((5, row.odprta_narocila))?;
            statement.next()?;
            statement.reset()?;
        }
        connection.execute("COMMIT")?;
        log::info!("Stored to table data: {}", row_data.len());
        Ok(())
    }


    fn create_poraba_table(&self, connection: &Connection) -> Result<(), Box<dyn std::error::Error>> {
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("
            CREATE TABLE IF NOT EXISTS porabe (
                id INTEGER PRIMARY KEY,
                material INTEGER NOT NULL,
                poraba REAL,
                date DATE
            );
        ")?;
        connection.execute("COMMIT")?;
        log::info!("Created porabe table");
        Ok(())
    }

    pub fn store_poraba_to_db(&self, rows: Vec<PorabaData>) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        let _ = self.drop_porabe();
        self.create_poraba_table(&connection)?;

        let mut statement = connection.prepare("
            INSERT INTO porabe (material, poraba, date) VALUES (?, ?, ?)
        ")?;

        connection.execute("BEGIN TRANSACTION")?;
        log::info!("Started inserting into poraba");
        for poraba_row in rows.iter() {
            statement.bind((1, poraba_row.material))?;
            statement.bind((2, poraba_row.poraba))?;
            statement.bind(&[(3, convert_to_sql(poraba_row.date).as_str())][..])?;
            statement.next()?;
            statement.reset()?;
        }
        connection.execute("COMMIT")?;
        log::info!("Stored to table Porabe: {}", rows.len());
        Ok(())
    }

    fn create_nabava_table(&self, connection: &Connection) -> Result<(), Box<dyn std::error::Error>> {
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("
            CREATE TABLE IF NOT EXISTS nabave (
                id INTEGER PRIMARY KEY,
                material INTEGER NOT NULL,
                nabava REAL,
                date DATE
            );
        ")?;
        connection.execute("COMMIT")?;
        log::info!("Created nabave table");
        Ok(())
    }

    pub fn store_nabava_to_db(&self, rows: Vec<NabavaData>) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        let _ = self.drop_nabave();
        self.create_nabava_table(&connection)?;

        let mut statement = connection.prepare("
            INSERT INTO nabave (material, nabava, date) VALUES (?, ?, ?)
        ")?;

        connection.execute("BEGIN TRANSACTION")?;
        log::info!("Started inserting into nabava");
        for nabava_row in rows.iter() {
            statement.bind((1, nabava_row.material))?;
            statement.bind((2, nabava_row.nabava))?;
            statement.bind(&[(3, convert_to_sql(nabava_row.date).as_str())][..])?;
            statement.next()?;
            statement.reset()?;
        }
        connection.execute("COMMIT")?;
        log::info!("Stored to table Nabave: {}", rows.len());
        Ok(())
    }

    fn create_razpolozljive_zaloge_table(&self, connection: &Connection) -> Result<(), Box<dyn std::error::Error>> {
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("
            CREATE TABLE IF NOT EXISTS razpolozljive_zaloge (
                id INTEGER PRIMARY KEY,
                material INTEGER NOT NULL,
                razpolozljiva_zaloga REAL,
                lokacija TEXT
            );
        ")?;
        connection.execute("COMMIT")?;
        log::info!("Created razpolozljive_zaloge table");
        Ok(())
    }

    pub fn store_razpolozljive_zaloge_to_db(&self, rows: Vec<RazpolozljivaZalogaRow>) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        let _ = self.drop_razpolozljive_zaloge();
        self.create_razpolozljive_zaloge_table(&connection)?;

        let mut statement = connection.prepare("
            INSERT INTO razpolozljive_zaloge (material, razpolozljiva_zaloga, lokacija) VALUES (?, ?, ?)
        ")?;

        connection.execute("BEGIN TRANSACTION")?;
        log::info!("Started inserting into razpolozljive_zaloge");
        for razpolozljiva_zaloga_row in rows.iter() {
            statement.bind((1, razpolozljiva_zaloga_row.material))?;
            statement.bind((2, razpolozljiva_zaloga_row.razpolozljiva_zaloga))?;
            statement.bind((3, razpolozljiva_zaloga_row.lokacija.as_str()))?;
            statement.next()?;
            statement.reset()?;
        }
        connection.execute("COMMIT")?;
        log::info!("Stored to razpolozljive_zaloge table: {}", rows.len());
        Ok(())
    }

    fn create_dobavitelji_table(&self, connection: &Connection) -> Result<(), Box<dyn std::error::Error>> {
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("
            CREATE TABLE IF NOT EXISTS dobavitelji (
                id INTEGER PRIMARY KEY,
                material INTEGER NOT NULL,
                dobavitelj TEXT,
                cena REAL,
                valuta TEXT
            );
        ")?;
        connection.execute("COMMIT")?;
        log::info!("Created dobavitelji table");
        Ok(())
    }

    pub fn store_dobavitelji_to_db(&self, rows: Vec<DobaviteljRow>) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        let _ = self.drop_dobavitelji();
        self.create_dobavitelji_table(&connection)?;

        let mut statement = connection.prepare("
            INSERT INTO dobavitelji (material, dobavitelj, cena, valuta) VALUES (?, ?, ?, ?)
        ")?;

        connection.execute("BEGIN TRANSACTION")?;
        log::info!("Started inserting into dobavitelji");
        for dobavitelj_row in rows.iter() {
            statement.bind((1, dobavitelj_row.material))?;
            statement.bind(&[(2, dobavitelj_row.dobavitelj.as_str())][..])?;
            statement.bind((3, dobavitelj_row.cena))?;
            statement.bind(&[(4, dobavitelj_row.valuta.as_str())][..])?;
            statement.next()?;
            statement.reset()?;
        }
        connection.execute("COMMIT")?;
        log::info!("Stored to dobavitelji table: {}", rows.len());
        Ok(())
    }


    fn create_sifrant_table(&self, connection: &Connection) -> Result<(), Box<dyn std::error::Error>> {
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("
            CREATE TABLE IF NOT EXISTS sifrant (
                material INTEGER PRIMARY KEY ,
                naziv_materiala TEXT,
                osnovna_merska_enota TEXT,
                nabavna_skupina TEXT,
                mrp_karakteristika TEXT
            );
        ")?;
        connection.execute("COMMIT")?;
        log::info!("Created sifrant table");
        Ok(())
    }

    pub fn store_sifrant_to_db(&self, rows: Vec<SifrantRow>) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        let _ = self.drop_sifrant();
        self.create_sifrant_table(&connection)?;

        let mut statement = connection.prepare("
            INSERT INTO sifrant (material, naziv_materiala, osnovna_merska_enota, nabavna_skupina, mrp_karakteristika) VALUES (?, ?, ?, ?, ?)
        ")?;

        connection.execute("BEGIN TRANSACTION")?;
        log::info!("Started inserting into sifrant");
        for sifrant_row in rows.iter() {
            statement.bind((1, sifrant_row.material))?;
            statement.bind(&[(2, sifrant_row.naziv_materiala.as_str())][..])?;
            statement.bind(&[(3, sifrant_row.osnovna_merska_enota.as_str())][..])?;
            statement.bind(&[(4, sifrant_row.nabavna_skupina.as_str())][..])?;
            statement.bind(&[(5, sifrant_row.mrp_karakteristika.as_str())][..])?;
            statement.next()?;
            statement.reset()?;
        }

        connection.execute("COMMIT")?;
        log::info!("Stored to sifrant table: {}", rows.len());
        Ok(())
    }


    fn create_pakiranje_table(&self, connection: &Connection) -> Result<(), Box<dyn std::error::Error>> {
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("
            CREATE TABLE IF NOT EXISTS pakiranja (
                material INTEGER PRIMARY KEY ,
                pakiranje TEXT NOT NULL
            );
        ")?;

        connection.execute("COMMIT")?;
        log::info!("Created pakiranja table");
        Ok(())
    }

    pub fn store_pakiranje(&self, pakiranje: (i64, String)) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        self.create_pakiranje_table(&connection)?;

        let mut statement = connection.prepare("
            INSERT INTO pakiranja (material, pakiranje) VALUES (?, ?) ON CONFLICT(material) DO UPDATE SET pakiranje = excluded.pakiranje
        ")?;
        connection.execute("BEGIN TRANSACTION")?;
        statement.bind((1, pakiranje.0))?;
        statement.bind((2, pakiranje.1.as_str()))?;
        statement.next()?;
        connection.execute("COMMIT")?;
        log::info!("Stored to pakiranja table");
        Ok(())
    }

    fn create_opomba_table(&self, connection: &Connection) -> Result<(), Box<dyn std::error::Error>> {
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("
            CREATE TABLE IF NOT EXISTS opombe (
                material INTEGER PRIMARY KEY ,
                opomba TEXT NOT NULL
            );
        ")?;
        connection.execute("CREATE INDEX IF NOT EXISTS idx_opombe_material ON opombe(material);")?;

        connection.execute("COMMIT")?;
        log::info!("Created opombe table");
        Ok(())
    }

    pub fn store_opomba_to_db(&self, opomba: (i64, String)) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        self.create_opomba_table(&connection)?;

        let mut statement = connection.prepare("
            INSERT INTO opombe (material, opomba) VALUES (?, ?) ON CONFLICT(material) DO UPDATE SET opomba = excluded.opomba
        ")?;
        connection.execute("BEGIN TRANSACTION")?;
        statement.bind((1, opomba.0))?;
        statement.bind((2, opomba.1.as_str()))?;
        statement.next()?;
        connection.execute("COMMIT")?;
        log::info!("Stored to opombe table");
        
        Ok(())
    }


    fn create_blagovna_skupina_table(&self, connection: &Connection) -> Result<(), Box<dyn std::error::Error>> {
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("
            CREATE TABLE IF NOT EXISTS blagovne_skupine (
                material INTEGER PRIMARY KEY ,
                blagovna_skupina TEXT NOT NULL
            );
        ")?;

        connection.execute("COMMIT")?;
        log::info!("Created blagovne_skupine table");
        Ok(())
    }

    pub fn store_blagovna_skupina(&self, blagovna_skupina: (i64, String)) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        self.create_blagovna_skupina_table(&connection)?;

        let mut statement = connection.prepare("
            INSERT INTO blagovne_skupine (material, blagovna_skupina) VALUES (?, ?) ON CONFLICT(material) DO UPDATE SET blagovna_skupina = excluded.blagovna_skupina
        ")?;
        connection.execute("BEGIN TRANSACTION")?;
        statement.bind((1, blagovna_skupina.0))?;
        statement.bind((2, blagovna_skupina.1.as_str()))?;
        statement.next()?;
        connection.execute("COMMIT")?;
        log::info!("Stored to blagovne_skupine table");
        Ok(())
    }


    fn create_min_zaloga_table(&self, connection: &Connection) -> Result<(), Box<dyn std::error::Error>> {
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("
            CREATE TABLE IF NOT EXISTS minimalne_zaloge (
                material INTEGER PRIMARY KEY,
                minimalna_zaloga REAL
            );
        ")?;

        connection.execute("COMMIT")?;
        log::info!("Created min. zaloge table");
        Ok(())
    }

    pub fn store_min_zaloga(&self, min_zaloga_row: (i64, Option<f64>)) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        self.create_min_zaloga_table(&connection)?;

        let mut statement = connection.prepare("
        INSERT INTO minimalne_zaloge (material, minimalna_zaloga) VALUES (?, ?) ON CONFLICT(material) DO UPDATE SET minimalna_zaloga = excluded.minimalna_zaloga
    ")?;
        connection.execute("BEGIN TRANSACTION")?;
        statement.bind((1, min_zaloga_row.0))?;
        statement.bind((2, min_zaloga_row.1))?;
        statement.next()?;
        connection.execute("COMMIT")?;
        log::info!("Stored to min. zaloge table");
        Ok(())
    }

    fn create_max_zaloga_table(&self, connection: &Connection) -> Result<(), Box<dyn std::error::Error>> {
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("
            CREATE TABLE IF NOT EXISTS maximalne_zaloge (
                material INTEGER PRIMARY KEY,
                maximalna_zaloga REAL
            );
        ")?;

        connection.execute("COMMIT")?;
        log::info!("Created Max. zaloge table");
        Ok(())
    }

    pub fn store_max_zaloga(&self, max_zaloga_row: (i64, Option<f64>)) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        self.create_max_zaloga_table(&connection)?;

        let mut statement = connection.prepare("
        INSERT INTO maximalne_zaloge (material, maximalna_zaloga) VALUES (?, ?) ON CONFLICT(material) DO UPDATE SET maximalna_zaloga = excluded.maximalna_zaloga
    ")?;
        connection.execute("BEGIN TRANSACTION")?;
        statement.bind((1, max_zaloga_row.0))?;
        statement.bind((2, max_zaloga_row.1))?;
        statement.next()?;
        connection.execute("COMMIT")?;
        log::info!("Stored to max. zaloge table");
        Ok(())
    }


    fn create_dobavni_rok_table(&self, connection: &Connection) -> Result<(), Box<dyn std::error::Error>> {
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("
            CREATE TABLE IF NOT EXISTS dobavni_roki (
                material INTEGER PRIMARY KEY,
                dobavni_rok REAL
            );
        ")?;
        connection.execute("CREATE INDEX IF NOT EXISTS idx_dobavni_roki_material ON dobavni_roki(material);")?;

        connection.execute("COMMIT")?;
        log::info!("Created dobavni_roki table");
        Ok(())
    }

    pub fn store_dobavni_rok(&self, dobavni_rok_row: (i64, Option<f64>)) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        self.create_dobavni_rok_table(&connection)?;

        let mut statement = connection.prepare("
        INSERT INTO dobavni_roki (material, dobavni_rok) VALUES (?, ?) ON CONFLICT(material) DO UPDATE SET dobavni_rok = excluded.dobavni_rok
    ")?;
        connection.execute("BEGIN TRANSACTION")?;
        statement.bind((1, dobavni_rok_row.0))?;
        statement.bind((2, dobavni_rok_row.1))?;
        statement.next()?;
        connection.execute("COMMIT")?;
        log::info!("Stored to dobavni_roki table");
        Ok(())
    }



    pub fn try_drop_view(&self) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        connection.execute("
            DROP VIEW view_podatki;
        ")?;
        Ok(())
    }

    pub fn try_create_view(&self) -> Result<(), Box<dyn std::error::Error>> {
        log::info!("trying to create view");
        let connection = sqlite::open(self.db_name.as_str())?;
        connection.execute("
            CREATE VIEW IF NOT EXISTS view_podatki AS
            SELECT
                s.material,
                s.naziv_materiala,
                s.osnovna_merska_enota,
                s.nabavna_skupina,
                s.mrp_karakteristika,
                d.zaloga,
                d.poraba_3m,
                d.poraba_24m,
                d.odprta_narocila,
                c.dobavni_rok,
                d.trenutna_zaloga_zadostuje_za_mesecev,
                d.trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev,
                COALESCE(dob.dobavitelji_list, ' ') AS dobavitelji,
                dob.cena,
                dob.valuta,
                razp_zal.razpolozljiva_zaloga,
                razp_zal.lokacija,
                min_z.minimalna_zaloga,
                max_z.maximalna_zaloga,
                blag_s.blagovna_skupina,
                pak.pakiranje,
                o.opomba
            FROM sifrant s
            LEFT JOIN data d ON s.material = d.material
            LEFT JOIN dobavni_roki c ON s.material = c.material
            LEFT JOIN opombe o ON s.material = o.material
            LEFT JOIN (
                SELECT material,
                LTRIM(GROUP_CONCAT(dobavitelj, ', '), ', ') AS dobavitelji_list,
                cena,
                valuta
                FROM dobavitelji GROUP BY material
            ) dob ON s.material = dob.material
            LEFT JOIN razpolozljive_zaloge razp_zal ON s.material = razp_zal.material
            LEFT JOIN minimalne_zaloge min_z ON s.material = min_z.material
            LEFT JOIN maximalne_zaloge max_z ON s.material = max_z.material
            LEFT JOIN blagovne_skupine blag_s ON s.material = blag_s.material
            LEFT JOIN pakiranja pak ON s.material = pak.material
            ;
        ")?;
        log::info!("Created View");
        Ok(())
    }
}

pub fn convert_to_sql(date: NaiveDate) -> String {
    date.format("%Y-%m-%d").to_string()
}

#[derive(Default, Clone, Debug)]
pub struct PorabaQuery {
    pub poraba: f64,
    pub month: String
}

impl PorabaQuery {
    fn query(material: i64, connection: &Connection) -> Result<Vec<Self>, Box<dyn std::error::Error>> {
        let mut rows: Vec<PorabaQuery> = Vec::with_capacity(2500);

        let sql = "
            SELECT
                SUM(poraba) as total,
                COALESCE(strftime('%Y-%m', date), '') as month
            FROM porabe
            WHERE material = ?
            GROUP BY month
            ORDER BY month;
        ".to_string();
        let mut statement = connection.prepare(sql)?;
        statement.bind((1, material))?;

        while let State::Row = statement.next()? {
            let mut row = PorabaQuery::default();
            row.poraba = statement.read::<f64, _>(0)?;
            row.month = statement.read::<String, _>(1)?;
            rows.push(row);
        }

        Ok(rows)
    }
}


#[derive(Default, Clone, Debug)]
pub struct NabavaQuery {
    pub nabava: f64,
    pub month: String
}

impl NabavaQuery {
    fn query(material: i64, connection: &Connection) -> Result<Vec<Self>, Box<dyn std::error::Error>> {
        let mut rows: Vec<NabavaQuery> = Vec::with_capacity(2500);

        let sql = "
            SELECT
                SUM(nabava) as total,
                COALESCE(strftime('%Y-%m', date), '') as month
            FROM nabave
            WHERE material = ?
            GROUP BY month
            ORDER BY month;
        ".to_string();
        let mut statement = connection.prepare(sql)?;
        statement.bind((1, material))?;

        while let State::Row = statement.next()? {
            let mut row = NabavaQuery::default();
            row.nabava = statement.read::<f64, _>(0)?;
            row.month = statement.read::<String, _>(1)?;
            rows.push(row);
        }

        Ok(rows)
    }
}




#[derive(Default, Clone)]
pub struct ViewQuery {
    pub material: i64,
    pub naziv_materiala: Option<String>,
    pub osnovna_merska_enota: Option<String>,
    pub nabavna_skupina: Option<String>,
    pub mrp_karakteristika: Option<String>,
    pub zaloga: Option<f64>,
    pub poraba_3m: Option<f64>,
    pub poraba_24m: Option<f64>,
    pub odprta_narocila: Option<f64>,
    pub dobavni_rok: Option<f64>,
    pub trenutna_zaloga_zadostuje_za_mesecev: Option<f64>,
    pub trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev: Option<f64>,
    pub dobavitelji: Option<String>,
    pub cena: Option<f64>,
    pub valuta: Option<String>,
    pub razpolozljiva_zaloga: Option<f64>,
    pub lokacija: Option<String>,
    pub minimalna_zaloga: Option<f64>,
    pub maximalna_zaloga: Option<f64>,
    pub blagovna_skupina: Option<String>,
    pub pakiranje: Option<String>,
    pub opomba: Option<String>,
}


impl ViewQuery {
    fn query(connection: &Connection, sort: &SortState) -> Result<Vec<Self>, Box<dyn std::error::Error>> {
        let mut rows = Vec::with_capacity(2500);

        let order_clause = sort.sql_order();
        let sql = format!("SELECT * FROM view_podatki {}", order_clause);
        let mut statement = connection.prepare(sql)?;

        while let State::Row = statement.next()? {
            let mut row = ViewQuery::default();
            row.material = statement.read(0)?;
            row.naziv_materiala = statement.read(1)?;
            row.osnovna_merska_enota = statement.read(2)?;
            row.nabavna_skupina = statement.read(3)?;
            row.mrp_karakteristika = statement.read(4)?;
            row.zaloga = statement.read(5)?;
            row.poraba_3m = statement.read(6)?;
            row.poraba_24m = statement.read(7)?;
            row.odprta_narocila = statement.read(8)?;
            row.dobavni_rok = statement.read(9)?;
            row.trenutna_zaloga_zadostuje_za_mesecev = statement.read(10)?;
            row.trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev = statement.read(11)?;
            row.dobavitelji = statement.read(12)?;
            row.cena = statement.read(13)?;
            row.valuta = statement.read(14)?;
            row.razpolozljiva_zaloga = statement.read(15)?;
            row.lokacija = statement.read(16)?;
            row.minimalna_zaloga = statement.read(17)?;
            row.maximalna_zaloga = statement.read(18)?;
            row.blagovna_skupina = statement.read(19)?;
            row.pakiranje = statement.read(20)?;
            row.opomba = statement.read(21)?;
            rows.push(row);
        }

        
        Ok(rows)
    }
}


#[derive(Default, Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum ViewQueryFields {
    #[default]
    Material,
    NazivMateriala,
    OsnovnaMerskaEnota,
    NabavnaSkupina,
    MRP,
    Zaloga,
    Poraba3M,
    Poraba24M,
    OdprtaNarocila,
    DobavniRok,
    TrenutnaZalogaZadostujeZaMesecev,
    TrenutnaZalogaInOdprtaNarocilaZadostujeZaMesecev,
    Dobavitelji,
    Cena,
    Valuta,
    RazpolozljivaZaloga,
    Lokacija,
    MinimalnaZaloga,
    MaximalnaZaloga,
    BlagovnaSkupina,
    Pakiranje,
    Opomba,
}

impl std::fmt::Display for ViewQueryFields {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let name = format!("{self:?}");
        let mut out = String::new();
        for (i, c) in name.chars().enumerate() {
            if c.is_uppercase() && i != 0 {
                out.push(' ');
            }
            out.push(c);
        }
        write!(f, "{out}")
    }
}

impl ViewQueryFields {
    pub const ALL: [ViewQueryFields; 22] = [
        ViewQueryFields::Material,
        ViewQueryFields::NazivMateriala,
        ViewQueryFields::OsnovnaMerskaEnota,
        ViewQueryFields::NabavnaSkupina,
        ViewQueryFields::MRP,
        ViewQueryFields::Zaloga,
        ViewQueryFields::Poraba3M,
        ViewQueryFields::Poraba24M,
        ViewQueryFields::OdprtaNarocila,
        ViewQueryFields::DobavniRok,
        ViewQueryFields::TrenutnaZalogaZadostujeZaMesecev,
        ViewQueryFields::TrenutnaZalogaInOdprtaNarocilaZadostujeZaMesecev,
        ViewQueryFields::Dobavitelji,
        ViewQueryFields::Cena,
        ViewQueryFields::Valuta,
        ViewQueryFields::RazpolozljivaZaloga,
        ViewQueryFields::Lokacija,
        ViewQueryFields::MinimalnaZaloga,
        ViewQueryFields::MaximalnaZaloga,
        ViewQueryFields::BlagovnaSkupina,
        ViewQueryFields::Pakiranje,
        ViewQueryFields::Opomba,
    ];
    pub fn construct_headers(&self, header: &mut egui_extras::TableRow, sort: &mut ViewQueryFields) {
        match self {
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
    pub fn construct_body(&self,
                          table_row: &mut egui_extras::TableRow,
                          index: usize,
                          row: &ViewQuery,
                          mut row_color: Color32,
                          poraba_nabava_data: &mut PorabaNabavaRows,
                          db_manager: &DBManager,
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
        match self {
            ViewQueryFields::Material => {
                table_row.col(|ui| {
                    ui.painter().rect_filled(ui.max_rect(), CornerRadius::same(0), row_color);
                    if ui.label(RichText::new(row.material.to_string()).underline().background_color(Color32::TRANSPARENT))
                        .on_hover_cursor(CursorIcon::PointingHand)
                        .clicked() {

                        poraba_nabava_data.query(row.material, row.naziv_materiala.as_ref().unwrap_or(&"".to_string()).as_str(), row.zaloga.unwrap_or(0.),  &db_manager);
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
                                    let _ = db_manager.store_dobavni_rok((
                                        row.material,
                                        parse_string_to_optional_f64(edit_dobavni_rok_input.as_str()),
                                    ));
                                    row_data.query(db_manager, sort_state);
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
                    ui.label(row.cena.map_or("".to_string(), |v| format_number_custom(v, 1)));
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
                                    let _ = db_manager.store_min_zaloga((
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
                                    let _ = db_manager.store_max_zaloga((
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
                                    let _ = db_manager.store_blagovna_skupina((
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
                            label_text = " ".repeat(73);
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
                                    let _ = db_manager.store_pakiranje((
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
                                    let _ = db_manager.store_opomba_to_db((
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
    pub(crate) fn as_str(&self) -> &'static str {
        match self {
            ViewQueryFields::Material => "material",
            ViewQueryFields::NazivMateriala => "naziv_materiala",
            ViewQueryFields::OsnovnaMerskaEnota => "osnovna_merska_enota",
            ViewQueryFields::NabavnaSkupina => "nabavna_skupina",
            ViewQueryFields::MRP => "mrp_karakteristika",
            ViewQueryFields::Zaloga => "zaloga",
            ViewQueryFields::Poraba3M => "poraba_3m",
            ViewQueryFields::Poraba24M => "poraba_24m",
            ViewQueryFields::OdprtaNarocila => "odprta_narocila",
            ViewQueryFields::DobavniRok => "dobavni_rok",
            ViewQueryFields::TrenutnaZalogaZadostujeZaMesecev => "trenutna_zaloga_zadostuje_za_mesecev",
            ViewQueryFields::TrenutnaZalogaInOdprtaNarocilaZadostujeZaMesecev => "trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev",
            ViewQueryFields::Dobavitelji => "dobavitelji",
            ViewQueryFields::Cena => "cena",
            ViewQueryFields::Valuta => "valuta",
            ViewQueryFields::RazpolozljivaZaloga => "razpolozljiva_zaloga",
            ViewQueryFields::MinimalnaZaloga => "minimalna_zaloga",
            ViewQueryFields::MaximalnaZaloga => "maximalna_zaloga",
            ViewQueryFields::BlagovnaSkupina => "blagovna_skupina",
            ViewQueryFields::Pakiranje => "pakiranje",
            ViewQueryFields::Lokacija => "lokacija",
            ViewQueryFields::Opomba => "opomba",
        }
    }
}

pub struct SortState {
    pub sort_column: ViewQueryFields,
    pub descending: bool,
}

impl Default for SortState {
    fn default() -> Self {
        SortState {
            sort_column: ViewQueryFields::default(),
            descending: false,
        }
    }
}

impl SortState {
    fn sql_order(&self) -> String {
        format!("ORDER BY {} {}",
                self.sort_column.as_str(),
                if self.descending { "DESC" } else { "ASC" }
        )
    }
}

