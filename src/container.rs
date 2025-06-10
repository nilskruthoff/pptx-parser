
use std::{
    collections::HashMap,
    io::Read,
    path::Path,
};
use super::{Error, Result, Slide};

pub struct PptxContainer {
    archive: zip::ZipArchive<std::fs::File>,
    slide_paths: Vec<String>,
}

impl PptxContainer {
    pub fn open(path: &Path) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        let mut archive = zip::ZipArchive::new(file)?;

        // Extrahiere nur die Pfade der Slides, nicht ihren Inhalt
        let mut slide_paths: Vec<String> = Vec::new();

        for i in 0..archive.len() {
            let file = archive.by_index(i)?;
            let name = file.name().to_string();

            if name.starts_with("ppt/slides/slide") && name.ends_with(".xml") {
                slide_paths.push(name);
            }
        }

        slide_paths.sort();

        Ok(Self { archive, slide_paths })
    }

    // Methode, um alle Slides auf einmal zu extrahieren (nicht-Streaming)
    pub fn parse_all(&mut self) -> Result<Vec<Slide>> {
        let mut slides = Vec::new();
        let count = self.slide_paths.len();

        for i in 0..count {
            let path = &self.slide_paths[i].clone();
            if let Some(slide) = self.load_slide(path)? {
                slides.push(slide);
            }
        }

        Ok(slides)
    }

    // Methode, um einen Iterator für Slides zu erhalten (Streaming)
    pub fn iter_slides(&mut self) -> SlideIterator {
        SlideIterator::new(self)
    }

    // Hilfsmethode zum Laden eines einzelnen Slides
    fn load_slide(&mut self, slide_path: &str) -> Result<Option<Slide>> {
        // Slide XML laden
        let slide_data = self.read_file_from_archive(slide_path)?;

        // Relationship-Datei für diesen Slide laden
        let rels_path = self.get_slide_rels_path(slide_path);
        let rels_data = self.read_file_from_archive(&rels_path).ok();

        // Extrahiere Bilder aus den Beziehungen
        let mut images = Vec::new();
        if let Some(ref rels_bytes) = rels_data {
            images = crate::parse_rels::parse_slide_rels(rels_bytes)?;
        }

        // Slide parsen
        let slide_number = Slide::extract_slide_number(slide_path).unwrap_or(0);
        let elements = crate::parse_xml::parse_slide_xml(&slide_data)?;

        // Bild-Ressourcen vorladen
        let mut image_data = HashMap::new();
        for img_ref in &images {
            let img_path = Self::get_full_image_path(slide_path, &img_ref.target);
            if let Ok(data) = self.read_file_from_archive(&img_path) {
                image_data.insert(img_ref.id.clone(), data);
            }
        }

        // Slide erstellen
        let mut slide = Slide::new(
            slide_path.to_string(),
            slide_number,
            elements,
            images,
            image_data,
        );

        slide.link_images();
        Ok(Some(slide))
    }

    // Hilfsmethoden für den Zugriff auf die Archivdateien
    fn read_file_from_archive(&mut self, path: &str) -> Result<Vec<u8>> {
        let mut file = self.archive.by_name(path)?;
        let mut content = Vec::new();
        file.read_to_end(&mut content)?;
        Ok(content)
    }

    fn get_slide_rels_path(&self, slide_path: &str) -> String {
        let mut rels_path = slide_path.to_string();
        if let Some(pos) = rels_path.rfind('/') {
            rels_path.insert_str(pos + 1, "_rels/");
        }
        rels_path.push_str(".rels");
        rels_path
    }

    fn get_full_image_path(slide_path: &str, target: &str) -> String {
        if target.starts_with("../") {
            let adjusted_target = target.trim_start_matches("../");
            format!("ppt/{}", adjusted_target)
        } else {
            let slide_dir = slide_path.rsplit_once('/').map(|(dir, _)| dir).unwrap_or("");
            format!("{}/{}", slide_dir, target)
        }
    }
}

pub struct SlideIterator<'a> {
    container: &'a mut PptxContainer,
    current_paths: Vec<String>, // Pfade beim Erstellen des Iterators kopieren
    current_index: usize,
}

impl<'a> SlideIterator<'a> {
    fn new(container: &'a mut PptxContainer) -> Self {
        let current_paths = container.slide_paths.clone();
        Self {
            container,
            current_paths,
            current_index: 0,
        }
    }
}

impl<'a> Iterator for SlideIterator<'a> {
    type Item = Result<Slide>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.current_index >= self.current_paths.len() {
            return None;
        }

        let slide_path = &self.current_paths[self.current_index];
        self.current_index += 1;

        match self.container.load_slide(slide_path) {
            Ok(Some(slide)) => Some(Ok(slide)),
            Ok(None) => self.next(), // Skip und weiter zum nächsten
            Err(e) => Some(Err(e)),
        }
    }
}