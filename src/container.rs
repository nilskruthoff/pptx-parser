use std::{
    collections::HashMap,
    io::Read,
    path::Path,
};
use super::{Result, Error};

pub struct PptxContainer {
    files: HashMap<String, Vec<u8>>,
}

impl PptxContainer {
    // Öffnet eine PPTX-Datei und lädt alle Dateien in den Speicher
    pub fn open(path: &Path) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        let mut archive = zip::ZipArchive::new(file)?;

        let mut files = HashMap::new();

        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            let mut content = Vec::new();
            file.read_to_end(&mut content)?;
            files.insert(file.name().to_string(), content);
        }

        Ok(Self { files })
    }

    pub fn read_slide(&self, slide_num: u32) -> Result<&[u8]> {
        let path = format!("ppt/slides/slide{}.xml", slide_num);
        self.files.get(&path)
            .map(|v| v.as_slice())
            .ok_or(Error::SlideNotFound)
    }
}

impl PptxContainer {
    pub fn get_slide_paths(&self) -> Vec<String> {
        let mut slides: Vec<String> = self.files
            .keys()
            .filter(|key| key.starts_with("ppt/slides/slide") && key.ends_with(".xml"))
            .cloned()
            .collect();
        slides.sort();
        slides
    }

    pub fn read_slide_by_path(&self, path: &str) -> Result<&[u8]> {
        self.files
            .get(path)
            .map(|v| v.as_slice())
            .ok_or(Error::SlideNotFound)
    }
}