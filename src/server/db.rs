use chrono::NaiveDate;
use eframe::egui::{Color32, CornerRadius, CursorIcon, RichText};
use rfd::{MessageButtons, MessageDialog, MessageDialogResult, MessageLevel};
use sqlite::{Connection, State};
use serde::{Deserialize, Serialize};
use crate::parse::{DobaviteljRow, NabavaData, PorabaData, RazpolozljivaZalogaRow, RowData, SifrantRow};

pub struct DBManager {
    pub db_name: String
}

impl DBManager {
    pub fn default() -> Self {
        let db_name = "server.sqlite3".to_string();

        let db_manager = Self {
            db_name,
        };
        let _ = db_manager.try_create_tables();
        let _ = db_manager.try_drop_view();
        let _ = db_manager.try_create_view();
        db_manager
    }
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









