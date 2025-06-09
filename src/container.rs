use super::{Error, Result, Slide};
use std::{
    collections::HashMap,
    io::Read,
    path::Path,
};

/// Holds the internal representation of a loaded PowerPoint (pptx) container.
///
/// `PptxContainer` provides functionalities for accessing slides and their resources
/// directly from a loaded pptx file. It parses and stores XML slides content,
/// relationships (`rels`) files and associated resources such as images.
pub struct PptxContainer<'a> {
    files: HashMap<String, Vec<u8>>,
    rels_files: HashMap<String, Vec<u8>>,
    _slides: Vec<Slide<'a>>,
}

impl<'a> PptxContainer<'a> {
    /// Opens a PowerPoint pptx file and initializes a `PptxContainer`.
    ///
    /// Processes the given file, extracting its internal files into memory. After initialization, the 
    /// container holds slide XML data, relationships files (*.rels), and associated resources.
    ///
    /// # Arguments
    ///
    /// - `path`: Path to the PPTX file.
    ///
    /// # Returns
    ///
    /// Returns a `Result` containing:
    /// - `Ok(PptxContainer)`: structured container instance upon successful file opening.
    /// - `Err(Error)`: if file access or internal unzip operations fail.
    ///
    /// # Errors
    ///
    /// Errors are returned on file access problems or failures during the unzipping process.
    pub fn open(path: &Path) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        let mut archive = zip::ZipArchive::new(file)?;

        let mut files = HashMap::new();
        let mut rels_files = HashMap::new(); // Neu
        let slides: Vec<Slide> = Vec::new();

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

        Ok(Self { files, _slides: slides, rels_files })
    }

    /// Parses the loaded pptx data within the container to structured slides.
    ///
    /// This method interprets XML slide data (`.xml`) and associated relationship files (`.rels`) into
    /// fully formed `Slide` resources, ready to be consumed.
    ///
    /// # Returns
    ///
    /// Returns a `Result` containing:
    /// - `Ok(Vec<Slide>)`: Vector of parsed slides with associated resources.
    /// - `Err(Error)`: If parsing fails on critical parts of XML data.
    ///
    /// # Errors
    ///
    /// Errors result from missing slide data or unexpected XML structure during parsing.
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
    /// Retrieves and returns sorted paths for all available slides in container.
    ///
    /// The resulting `Vec<String>` includes all slide XML file paths (e.g. "ppt/slides/slide1.xml")
    /// sorted by their slide numbers.
    ///
    /// # Returns
    ///
    /// A sorted vector of slide paths.
    pub fn get_slide_paths(&self) -> Vec<String> {
        let mut slides: Vec<String> = self.files
            .keys()
            .filter(|key| key.starts_with("ppt/slides/slide") && key.ends_with(".xml"))
            .cloned()
            .collect();
        slides.sort();
        slides
    }

    /// Retrieves raw XML data for a specific slide by its path.
    ///
    /// Fetches the XML content corresponding to the provided slide path.
    ///
    /// # Arguments
    ///
    /// - `path`: A `&str` representing a specific slide path.
    ///
    /// # Returns
    ///
    /// Returns a `Result` containing either:
    /// - `Ok(&[u8])`: A reference to raw slide XML data.
    /// - `Err(Error::SlideNotFound)`: If no slide is found at the specified path.
    ///
    /// # Errors
    ///
    /// Raised if the slide path doesn't exist in the pptx data.
    pub fn read_slide_by_path(&self, path: &str) -> Result<&[u8]> {
        self.files
            .get(path)
            .map(|v| v.as_slice())
            .ok_or(Error::SlideNotFound)
    }

    /// Retrieves relationship XML data (*.rels) content by relation-path.
    ///
    /// Fetches XML data from the internal relationship file store (`rels_files`) for
    /// given paths.
    ///
    /// # Arguments
    ///
    /// - `path`: Path string which references an internal `.rels` resource.
    ///
    /// # Returns
    ///
    /// Returns a `Result` containing:
    /// - `Ok(&[u8])`: Raw relationship XML data.
    /// - `Err(Error::SlideNotFound)`: If no relationship file is found at specified path.
    ///
    /// # Errors
    ///
    /// Returned when the requested `.rels` file does not exist in pptx data.
    fn get_slide_rels_path(&self, slide_path: &str) -> String {
        let mut rels_path = slide_path.to_string();
        if let Some(pos) = rels_path.rfind('/') {
            rels_path.insert_str(pos + 1, "_rels/");
        }
        rels_path.push_str(".rels");
        rels_path
    }

    /// Constructs and returns the relationship path (*.rels) for a given slide.
    ///
    /// # Arguments
    ///
    /// - `slide_path`: Slide path (`ppt/slides/slideX.xml`) for which the rels path should be built.
    ///
    /// # Returns
    ///
    /// A String path representing the corresponding `.rels` resource path.
    pub fn read_rels_by_path(&self, path: &str) -> Result<&[u8]> {
        self.rels_files
            .get(path)
            .map(|v| v.as_slice())
            .ok_or(Error::SlideNotFound)
    }
}