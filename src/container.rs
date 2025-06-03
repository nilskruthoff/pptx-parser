use super::{Error, Result, SlideElement, Slide};
use std::{
    collections::HashMap,
    io::Read,
    path::Path,
};

pub struct PptxContainer {
    files: HashMap<String, Vec<u8>>,
    slides: Vec<Slide>,
}

impl PptxContainer {
    pub fn open(path: &Path) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        let mut archive = zip::ZipArchive::new(file)?;

        let mut files = HashMap::new();
        let mut slides:  Vec<Slide> = Vec::new();

        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            let mut content = Vec::new();
            file.read_to_end(&mut content)?;
            files.insert(file.name().to_string(), content);
        }

        Ok(Self { files, slides })
    }

    pub fn parse(&self) -> Result<Vec<Slide>> {
        let mut slides: Vec<Slide> = Vec::new();
        let slide_paths = self.get_slide_paths();

        for path in slide_paths {
            let slide_data = self.read_slide_by_path(&path)?;
            let slide = Slide::parse(slide_data, path)?;
            slides.push(slide);
        }
        Ok(slides)
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