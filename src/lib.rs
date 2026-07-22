use std::str::FromStr;
use chrono::NaiveDate;
use serde::{Deserialize, Serialize};
use sqlite::{Connection, State};

pub fn convert_to_sql(date: NaiveDate) -> String {
    date.format("%Y-%m-%d").to_string()
}

#[derive(Default, Clone, Debug)]
pub struct PorabaQuery {
    pub poraba: f64,
    pub month: String
}

impl PorabaQuery {
    pub fn query(material: i64, connection: &Connection) -> Result<Vec<Self>, Box<dyn std::error::Error>> {
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
    pub fn query(material: i64, connection: &Connection) -> Result<Vec<Self>, Box<dyn std::error::Error>> {
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
    pub fn query(connection: &Connection, sort: &SortState) -> Result<Vec<Self>, Box<dyn std::error::Error>> {
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


#[derive(Serialize, Deserialize)]
pub struct SortState {
    pub sort_column: ViewQueryFields,
    pub descending: bool,
}
impl FromStr for SortState {
    type Err = serde_json::Error;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        serde_json::from_str(s)
    }
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