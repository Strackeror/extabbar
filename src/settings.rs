use std::error::Error;

use crate::get_dll_path;
use serde::Deserialize;

#[derive(Deserialize, Debug)]
pub struct Settings {
    pub dark_mode: bool,
}

impl Default for Settings {
    fn default() -> Self {
        Self { dark_mode: true }
    }
}

pub fn current_settings() -> Settings {
    || -> Result<Settings, Box<dyn Error>> {
        let file = std::fs::File::open(get_dll_path().with_file_name("settings.json"))?;
        Ok(serde_json::from_reader(file)?)
    }()
    .map(|s| {
        log::info!("Read settings {:?}", s);
        s
    })
    .unwrap_or_default()
}
