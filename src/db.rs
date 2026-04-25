use sqlite::{Connection, State};
use crate::parse::{DobaviteljRow, RowData, SifrantRow};

pub struct DBManager {
    pub db_name: String
}

impl DBManager {
    pub fn try_create_tables(&self) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        self.create_sifrant_table(&connection)?;
        self.create_dobavni_rok_table(&connection)?;
        self.create_data_table(&connection)?;
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


    fn create_data_table(&self, connection: &Connection) -> Result<(), Box<dyn std::error::Error>> {
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("
            CREATE TABLE IF NOT EXISTS data (
                material INTEGER PRIMARY KEY,
                zaloga REAL,
                poraba_3m REAL,
                poraba_24m REAL,
                odprta_narocila REAL,
                trenutna_zaloga_zadostuje_za_mesecev REAL
                    GENERATED ALWAYS AS (
                        CASE
                            WHEN poraba_3m = 0 OR poraba_3m IS NULL THEN
                            CASE
                                WHEN poraba_24m = 0 OR poraba_24m IS NULL THEN NULL
                                ELSE zaloga / poraba_24m
                            END
                            ELSE zaloga / poraba_3m
                        END
                    ) VIRTUAL,
                trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev REAL
                    GENERATED ALWAYS AS (
                        CASE
                            WHEN poraba_3m = 0 OR poraba_3m IS NULL THEN
                            CASE
                                WHEN poraba_24m = 0 OR poraba_24m IS NULL THEN NULL
                                ELSE (zaloga + odprta_narocila) / poraba_24m
                            END
                            ELSE (zaloga + odprta_narocila) / poraba_3m
                        END
                    ) VIRTUAL
            );
        ")?;
        connection.execute("CREATE INDEX IF NOT EXISTS idx_data_material ON data(material);")?;


        connection.execute("COMMIT")?;
        Ok(())
    }

    pub fn drop_non_permanent(&self) -> Result<(), Box<dyn std::error::Error>> {

        let connection = sqlite::open(self.db_name.as_str())?;
        connection.execute("BEGIN TRANSACTION")?;
        self.try_drop_view()?;
        connection.execute("DROP TABLE sifrant;")?;
        connection.execute("DROP TABLE data;")?;
        connection.execute("DROP TABLE dobavitelji;")?;

        connection.execute("COMMIT")?;
        Ok(())
    }


    pub fn store_to_data(&self, row_data: Vec<RowData>) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        self.create_data_table(&connection)?;

        let mut statement = connection.prepare("
            INSERT INTO data (material, zaloga, poraba_3m, poraba_24m, odprta_narocila) VALUES (?, ?, ?, ?, ?) ON CONFLICT(material) DO UPDATE SET
                zaloga = excluded.zaloga,
                poraba_3m = excluded.poraba_3m,
                poraba_24m = excluded.poraba_24m,
                odprta_narocila = excluded.odprta_narocila
        ")?;
        connection.execute("BEGIN TRANSACTION")?;
        for row in row_data {
            //println!("{}", index);
            statement.bind((1, row.material))?;
            statement.bind((2, row.zaloga))?;
            statement.bind((3, row.poraba_3m))?;
            statement.bind((4, row.poraba_24m))?;
            statement.bind((5, row.odprta_narocila))?;
            statement.next()?;
            statement.reset()?;
        }
        connection.execute("COMMIT")?;
        //println!("commited!");

        self.try_create_view()?;
        //println!("after try_create_view");
        Ok(())
    }



    fn create_dobavitelji_table(&self, connection: &Connection) -> Result<(), Box<dyn std::error::Error>> {
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("
            CREATE TABLE IF NOT EXISTS dobavitelji (
                id INTEGER PRIMARY KEY,
                material INTEGER NOT NULL,
                dobavitelj TEXT
            );
        ")?;
        connection.execute("COMMIT")?;
        Ok(())
    }

    pub fn store_dobavitelji_to_db(&self, rows: Vec<DobaviteljRow>) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        self.create_dobavitelji_table(&connection)?;

        let mut statement = connection.prepare("
            INSERT INTO dobavitelji (material, dobavitelj) VALUES (?, ?)
        ")?;

        connection.execute("BEGIN TRANSACTION")?;
        for (index, dobavitelj_row) in rows.iter().enumerate() {
            println!("index: {}", index);
            statement.bind((1, dobavitelj_row.material))?;
            statement.bind(&[(2, dobavitelj_row.dobavitelj.as_str())][..])?;
            statement.next()?;
            statement.reset()?;
        }
        println!("finished");
        connection.execute("COMMIT")?;

        self.try_create_view()?;
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
        Ok(())
    }

    pub fn store_sifrant_to_db(&self, rows: Vec<SifrantRow>) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        self.create_sifrant_table(&connection)?;

        let mut statement = connection.prepare("
            INSERT INTO sifrant (material, naziv_materiala, osnovna_merska_enota, nabavna_skupina, mrp_karakteristika) VALUES (?, ?, ?, ?, ?)
        ")?;

        connection.execute("BEGIN TRANSACTION")?;
        for (index, sifrant_row) in rows.iter().enumerate() {
            println!("index: {}", index);
            statement.bind((1, sifrant_row.material))?;
            statement.bind(&[(2, sifrant_row.naziv_materiala.as_str())][..])?;
            statement.bind(&[(3, sifrant_row.osnovna_merska_enota.as_str())][..])?;
            statement.bind(&[(4, sifrant_row.nabavna_skupina.as_str())][..])?;
            statement.bind(&[(5, sifrant_row.mrp_karakteristika.as_str())][..])?;
            statement.next()?;
            statement.reset()?;
        }

        connection.execute("COMMIT")?;

        self.try_create_view()?;
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


        self.try_create_view()?;
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


        self.try_create_view()?;
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


        self.try_create_view()?;
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


        self.try_create_view()?;
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


        self.try_create_view()?;
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


        self.try_create_view()?;
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
                LTRIM(GROUP_CONCAT(dobavitelj, ', '), ', ') AS dobavitelji_list
                FROM dobavitelji GROUP BY material
            ) dob ON s.material = dob.material
            LEFT JOIN minimalne_zaloge min_z ON s.material = min_z.material
            LEFT JOIN maximalne_zaloge max_z ON s.material = max_z.material
            LEFT JOIN blagovne_skupine blag_s ON s.material = blag_s.material
            LEFT JOIN pakiranja pak ON s.material = pak.material
            ;
        ")?;
        Ok(())
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
            row.minimalna_zaloga = statement.read(13)?;
            row.maximalna_zaloga = statement.read(14)?;
            row.blagovna_skupina = statement.read(15)?;
            row.pakiranje = statement.read(16)?;
            row.opomba = statement.read(17)?;
            rows.push(row);
        }

        Ok(rows)
    }
}


#[derive(Default, Clone, Copy, Debug, PartialEq, Eq)]
pub enum SortColumn {
    #[default]
    Material,
    NazivMateriala,
    OsnovnaMerskaEnota,
    NabavnaSkupina,
    _MRP,
    Zaloga,
    Poraba3M,
    Poraba24M,
    OdprtaNarocila,
    DobavniRok,
    TrenutnaZalogaZadostujeZaMesecev,
    TrenutnaZalogaInOdprtaNarocilaZadostujeZaMesecev,
    Dobavitelji,
    MinimalnaZaloga,
    MaximalnaZaloga,
    BlagovnaSkupina,
    Pakiranje,
    Opomba,
}

impl SortColumn {
    pub(crate) fn as_str(&self) -> &'static str {
        match self {
            SortColumn::Material => "material",
            SortColumn::NazivMateriala => "naziv_materiala",
            SortColumn::OsnovnaMerskaEnota => "osnovna_merska_enota",
            SortColumn::NabavnaSkupina => "nabavna_skupina",
            SortColumn::_MRP => "mrp_karakteristika",
            SortColumn::Zaloga => "zaloga",
            SortColumn::Poraba3M => "poraba_3m",
            SortColumn::Poraba24M => "poraba_24m",
            &SortColumn::OdprtaNarocila => "odprta_narocila",
            SortColumn::DobavniRok => "dobavni_rok",
            SortColumn::TrenutnaZalogaZadostujeZaMesecev => "trenutna_zaloga_zadostuje_za_mesecev",
            SortColumn::TrenutnaZalogaInOdprtaNarocilaZadostujeZaMesecev => "trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev",
            SortColumn::Dobavitelji => "dobavitelji",
            SortColumn::MinimalnaZaloga => "minimalna_zaloga",
            SortColumn::MaximalnaZaloga => "maximalna_zaloga",
            SortColumn::BlagovnaSkupina => "blagovna_skupina",
            SortColumn::Pakiranje => "pakiranje",
            SortColumn::Opomba => "opomba",
        }
    }
}

pub struct SortState {
    pub sort_column: SortColumn,
    pub descending: bool,
}

impl Default for SortState {
    fn default() -> Self {
        SortState {
            sort_column: SortColumn::default(),
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

