use serde::{Deserialize, Serialize};
use magneti_excel::ViewQueryFields;
use magneti_excel::ViewQueryFields::{BlagovnaSkupina, Cena, Dobavitelji, DobavniRok, Lokacija, Material, MaximalnaZaloga, MinimalnaZaloga, NabavnaSkupina, NazivMateriala, OdprtaNarocila, Opomba, OsnovnaMerskaEnota, Pakiranje, Poraba24M, Poraba3M, RazpolozljivaZaloga, TrenutnaZalogaInOdprtaNarocilaZadostujeZaMesecev, TrenutnaZalogaZadostujeZaMesecev, Valuta, Zaloga, MRP};

#[derive(Serialize, Deserialize)]
pub struct Config {
    ip_addr: String,
    display_columns: Vec<ViewQueryFields>,
}

impl Config {
    pub fn get_url(&self) -> &str {
        self.ip_addr.as_str()
    }
    pub fn update_url(&mut self, url: &str) {
        log::info!("Updating server IP");
        self.ip_addr = format!("http://{}:8080", url);
        self.save();
    }

    pub fn get_mut_display_columns(&mut self) -> &mut Vec<ViewQueryFields> {
        &mut self.display_columns
    }

    fn path() -> std::path::PathBuf {
        "config.json".into()
    }

    pub fn load() -> Self {
        std::fs::read_to_string(Self::path())
            .ok()
            .and_then(|s| serde_json::from_str(&s).ok())
            .unwrap_or_else(|| Config::default() )
    }

    pub fn save(&self) {
        if let Ok(json) = serde_json::to_string_pretty(self) {
            let _ = std::fs::write(Self::path(), json);
        }
    }
    pub fn default() -> Self {
        Config {
            ip_addr: "127.0.0.1".to_string(),
            display_columns: vec![
                Material,
                NazivMateriala,
                RazpolozljivaZaloga,
                Zaloga,
                Poraba3M,
                Poraba24M,
                OdprtaNarocila,
                DobavniRok,
                TrenutnaZalogaZadostujeZaMesecev,
                TrenutnaZalogaInOdprtaNarocilaZadostujeZaMesecev,
                Cena,
                Valuta,
                OsnovnaMerskaEnota,
                MinimalnaZaloga,
                MaximalnaZaloga,
                Pakiranje,
                Lokacija,
                MRP,
                BlagovnaSkupina,
                Opomba,
                NabavnaSkupina,
                Dobavitelji,
            ],
        }
    }
}