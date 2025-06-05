use super::{Error, Result, SlideElement, Slide};
use std::{
    collections::HashMap,
    io::Read,
    path::Path,
};

pub struct PptxContainer<'a> {
    files: HashMap<String, Vec<u8>>,
    rels_files: HashMap<String, Vec<u8>>,
    slides: Vec<Slide<'a>>,
}

impl<'a> PptxContainer<'a> {
    pub fn open(path: &Path) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        let mut archive = zip::ZipArchive::new(file)?;

        let mut files = HashMap::new();
        let mut rels_files = HashMap::new(); // Neu
        let mut slides: Vec<Slide> = Vec::new();

        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            let mut content = Vec::new();
            file.read_to_end(&mut content)?;

            let name = file.name().to_string();
            if name.ends_with(".rels") {
                rels_files.insert(name, content);
            } else {
                files.insert(name, content);
            }
        }

        let container_path = path
            .parent()
            .map(|p| p.to_string_lossy().into_owned())
            .unwrap_or_else(|| ".".to_string());

        Ok(Self { files, slides, rels_files })
    }

    pub fn parse(&self) -> Result<Vec<Slide>> {
        let mut slides: Vec<Slide> = Vec::new();
        let slide_paths = self.get_slide_paths();

        for path in slide_paths {
            let slide_data = self.read_slide_by_path(&path)?;
            let rels_path = self.get_slide_rels_path(&path);
            let rels_data = self.read_rels_by_path(&rels_path).ok();

            let mut slide = Slide::parse(slide_data, path, rels_data, &self.files)?;
            slide.link_images();
            slides.push(slide);
        }
        Ok(slides)
    }
}

impl<'a> PptxContainer<'a> {
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

    fn get_slide_rels_path(&self, slide_path: &str) -> String {
        let mut rels_path = slide_path.to_string();
        if let Some(pos) = rels_path.rfind('/') {
            rels_path.insert_str(pos + 1, "_rels/");
        }
        rels_path.push_str(".rels");
        rels_path
    }

    pub fn read_rels_by_path(&self, path: &str) -> Result<&[u8]> {
        self.rels_files
            .get(path)
            .map(|v| v.as_slice())
            .ok_or(Error::SlideNotFound)
    }
}