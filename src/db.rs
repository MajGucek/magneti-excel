use sqlite::{Connection, State};
use crate::parse::{RowData, SifrantRow};

pub struct DBManager {
    pub db_name: String
}

impl DBManager {
    pub fn try_create_tables(&self) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        self.create_data_table(&connection)?;
        self.create_sifrant_table(&connection)?;
        self.create_opomba_table(&connection)?;
        self.create_dobavni_rok_table(&connection)?;

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
                poraba_12m REAL,
                odprta_narocila REAL,
                trenutna_zaloga_zadostuje_za_mesecev REAL
                    GENERATED ALWAYS AS (
                        CASE
                            WHEN poraba_3m = 0 THEN NULL
                            ELSE zaloga / poraba_3m
                        END
                    ) VIRTUAL,
                trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev REAL
                    GENERATED ALWAYS AS (
                        CASE
                            WHEN poraba_3m = 0 THEN NULL
                            ELSE (zaloga + odprta_narocila) / poraba_3m
                        END
                    ) VIRTUAL
            );
        ")?;
        connection.execute("CREATE INDEX IF NOT EXISTS idx_data_material ON data(material);")?;


        connection.execute("COMMIT")?;
        Ok(())
    }

    pub fn drop_data(&self) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("DROP TABLE data;")?;

        connection.execute("COMMIT")?;
        Ok(())
    }


    pub fn store_to_data(&self, row_data: Vec<RowData>) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        self.create_data_table(&connection)?;

        let mut statement = connection.prepare("
            INSERT INTO data (material, zaloga, poraba_3m, poraba_12m, odprta_narocila) VALUES (?, ?, ?, ?, ?) ON CONFLICT(material) DO UPDATE SET
                zaloga = excluded.zaloga,
                poraba_3m = excluded.poraba_3m,
                poraba_12m = excluded.poraba_12m,
                odprta_narocila = excluded.odprta_narocila
        ")?;
        connection.execute("BEGIN TRANSACTION")?;
        for row in row_data {
            statement.bind((1, row.material))?;
            statement.bind((2, row.zaloga))?;
            statement.bind((3, row.poraba_3m))?;
            statement.bind((4, row.poraba_12m))?;
            statement.bind((5, row.odprta_narocila))?;
            statement.next()?;
            statement.reset()?;
        }
        connection.execute("COMMIT")?;

        self.try_create_view(&connection);
        Ok(())
    }


    fn create_sifrant_table(&self, connection: &Connection) -> Result<(), Box<dyn std::error::Error>> {
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("
            CREATE TABLE IF NOT EXISTS sifrant (
                material INTEGER PRIMARY KEY ,
                naziv_materiala TEXT,
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
            INSERT INTO sifrant (material, naziv_materiala, nabavna_skupina, mrp_karakteristika) VALUES (?, ?, ?, ?)
        ")?;

        connection.execute("BEGIN TRANSACTION")?;
        for (index, sifrant_row) in rows.iter().enumerate() {
            println!("index: {}", index);
            statement.bind((1, sifrant_row.material))?;
            statement.bind(&[(2, sifrant_row.naziv_materiala.as_str())][..])?;
            statement.bind(&[(3, sifrant_row.nabavna_skupina.as_str())][..])?;
            statement.bind(&[(4, sifrant_row.mrp_karakteristika.as_str())][..])?;
            statement.next()?;
            statement.reset()?;
        }

        connection.execute("COMMIT")?;

        self.try_create_view(&connection);
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


        self.try_create_view(&connection);
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


        self.try_create_view(&connection);
        Ok(())
    }

    pub fn drop_all_tables(&self) -> Result<(), Box<dyn std::error::Error>> {
        let connection = sqlite::open(self.db_name.as_str())?;
        connection.execute("BEGIN TRANSACTION")?;
        connection.execute("
            DROP TABLE data;
        ")?;
        connection.execute("
            DROP TABLE sifrant;
        ")?;

        connection.execute("COMMIT")?;
        Ok(())
    }


    fn try_create_view(&self, connection: &Connection) {
        let _ = connection.execute("
        CREATE VIEW IF NOT EXISTS view_podatki AS
        SELECT
            s.material,
            s.naziv_materiala,
            s.nabavna_skupina,
            s.mrp_karakteristika,
            d.zaloga,
            d.poraba_3m,
            d.poraba_12m,
            d.odprta_narocila,
            c.dobavni_rok,
            d.trenutna_zaloga_zadostuje_za_mesecev,
            d.trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev,
            o.opomba
        FROM sifrant s
        LEFT JOIN data d ON s.material = d.material
        LEFT JOIN dobavni_roki c ON s.material = c.material
        LEFT JOIN opombe o ON s.material = o.material;
    ");
    }
}








#[derive(Default, Clone)]
pub struct ViewQuery {
    pub material: i64,
    pub naziv_materiala: Option<String>,
    pub nabavna_skupina: Option<String>,
    pub mrp_karakteristika: Option<String>,
    pub zaloga: Option<f64>,
    pub poraba_3m: Option<f64>,
    pub poraba_12m: Option<f64>,
    pub odprta_narocila: Option<f64>,
    pub dobavni_rok: Option<f64>,
    pub trenutna_zaloga_zadostuje_za_mesecev: Option<f64>,
    pub trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev: Option<f64>,
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
            row.nabavna_skupina = statement.read(2)?;
            row.mrp_karakteristika = statement.read(3)?;
            row.zaloga = statement.read(4)?;
            row.poraba_3m = statement.read(5)?;
            row.poraba_12m = statement.read(6)?;
            row.odprta_narocila = statement.read(7)?;
            row.dobavni_rok = statement.read(8)?;
            row.trenutna_zaloga_zadostuje_za_mesecev = statement.read(9)?;
            row.trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev = statement.read(10)?;
            row.opomba = statement.read(11)?;
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
    NabavnaSkupina,
    MRP,
    Zaloga,
    Poraba3M,
    Poraba12M,
    OdprtaNarocila,
    DobavniRok,
    TrenutnaZalogaZadostujeZaMesecev,
    TrenutnaZalogaInOdprtaNarocilaZadostujeZaMesecev,
    Opomba,
}

impl SortColumn {
    pub(crate) fn as_str(&self) -> &'static str {
        match self {
            SortColumn::Material => "material",
            SortColumn::NazivMateriala => "naziv_materiala",
            SortColumn::NabavnaSkupina => "nabavna_skupina",
            SortColumn::MRP => "mrp_karakteristika",
            SortColumn::Zaloga => "zaloga",
            SortColumn::Poraba3M => "poraba_3m",
            SortColumn::Poraba12M => "poraba_12m",
            &SortColumn::OdprtaNarocila => "odprta_narocila",
            SortColumn::DobavniRok => "dobavni_rok",
            SortColumn::TrenutnaZalogaZadostujeZaMesecev => "trenutna_zaloga_zadostuje_za_mesecev",
            SortColumn::TrenutnaZalogaInOdprtaNarocilaZadostujeZaMesecev => "trenutna_zaloga_in_odprta_narocila_zadostuje_za_mesecev",
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

