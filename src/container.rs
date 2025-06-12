use super::{Result, Slide};
use std::{
    collections::HashMap,
    io::Read,
    path::Path,
};
use crate::parser_config::ParserConfig;

/// Holds the internal representation of a loaded PowerPoint (pptx) container.
///
/// `PptxContainer` provides functionalities for accessing slides and their resources
/// directly from a loaded pptx file. It parses and stores XML slides content,
/// relationships (`rels`) files, and associated resources such as images.
pub struct PptxContainer {
    config: ParserConfig,
    archive: zip::ZipArchive<std::fs::File>,
    slide_paths: Vec<String>,
}

impl PptxContainer {
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
    pub fn open(path: &Path, config: ParserConfig) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        let mut archive = zip::ZipArchive::new(file)?;

        let mut slide_paths: Vec<String> = Vec::new();

        for i in 0..archive.len() {
            let file = archive.by_index(i)?;
            let name = file.name().to_string();

            if name.starts_with("ppt/slides/slide") && name.ends_with(".xml") {
                slide_paths.push(name);
            }
        }

        slide_paths.sort();

        Ok(Self { archive, slide_paths, config })
    }

    /// Parses the data of all slides for each path present in the containers' `slide_path` vector.
    /// 
    /// # Note
    /// Parsing is synchronous and in-memory, image data is extracted
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

    
    pub fn iter_slides(&mut self) -> SlideIterator {
        SlideIterator::new(self)
    }

    /// Loads a slide from the PPTX file by its index.
    ///
    /// # Arguments
    ///
    /// * `index` - The zero-based index of the slide to load.
    ///
    /// # Returns
    ///
    /// * `Ok(Some(Slide))` - The parsed slide if found and successfully processed.
    /// * `Ok(None)` - If the index is out of bounds.
    /// * `Err(_)` - If there was an error loading or parsing the slide.
    ///
    /// # Example
    ///
    /// ```
    /// // let mut streamer = open(Path::new("presentation.pptx"))?;
    /// // if let Ok(Some(slide)) = streamer.load_slide(0) {
    ///     // println!("Loaded first slide: {}", slide.slide_number);
    /// // }
    /// ```
    fn load_slide(&mut self, slide_path: &str) -> Result<Option<Slide>> {
        // load xml data
        let slide_data = self.read_file_from_archive(slide_path)?;

        // load relationship file
        let rels_path = self.get_slide_rels_path(slide_path);
        let rels_data = self.read_file_from_archive(&rels_path).ok();

        // parse slide and preload images
        let slide_number = Slide::extract_slide_number(slide_path).unwrap_or(0);
        let elements = crate::parse_xml::parse_slide_xml(&slide_data)?;
        
        let mut images = Vec::new();
        let mut image_data = HashMap::new();
        
        if self.config.extract_images {
            // extract images from relationships
            if let Some(ref rels_bytes) = rels_data {
                images = crate::parse_rels::parse_slide_rels(rels_bytes)?;
            }

            for img_ref in &images {
                let img_path = Self::get_full_image_path(slide_path, &img_ref.target);
                if let Ok(data) = self.read_file_from_archive(&img_path) {
                    image_data.insert(img_ref.id.clone(), data);
                }
            }
        }
        
        let config = self.config.clone();

        let mut slide = Slide::new(
            slide_path.to_string(),
            slide_number,
            elements,
            images,
            image_data,
            config,
        );

        slide.link_images();
        Ok(Some(slide))
    }

    /// Reads a file from the PPTX archive by its internal path.
    ///
    /// # Arguments
    ///
    /// * `path` - The internal path of the file within the PPTX archive.
    ///
    /// # Returns
    ///
    /// * `Ok(Vec<u8>)` - The content of the file as a byte vector.
    /// * `Err(_)` - If the file could not be found or read.
    ///
    /// # Notes
    ///
    /// This is an internal method used to extract individual files from the
    /// PPTX archive (which is essentially a ZIP file).
    fn read_file_from_archive(&mut self, path: &str) -> Result<Vec<u8>> {
        let mut file = self.archive.by_name(path)?;
        let mut content = Vec::new();
        file.read_to_end(&mut content)?;
        Ok(content)
    }

    /// Constructs the path to the relationships file for a given slide.
    ///
    /// # Arguments
    ///
    /// * `slide_path` - The internal path of the slide file.
    ///
    /// # Returns
    ///
    /// The path to the corresponding relationships (.rels) file.
    ///
    /// # Example
    ///
    /// ```
    /// // For a slide path "ppt/slides/slide1.xml"
    /// // Returns "ppt/slides/_rels/slide1.xml.rels"
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

/// An iterator for streaming slides from a PPTX file.
///
/// This iterator allows processing slides one by one, which is more
/// memory-efficient than loading all slides at once. It iterates through
/// all slides in the presentation in order.
///
/// # Example
///
/// ```
/// // let mut streamer = PptxStreamer::open(Path::new("presentation.pptx"))?;
/// // for slide_result in streamer.iter_slides() {
/// //    match slide_result {
/// //        Ok(slide) => println!("Processing slide {}", slide.slide_number),
/// //        Err(e) => eprintln!("Error: {:?}", e),
/// //    }
/// // }
/// ```
pub struct SlideIterator<'a> {
    container: &'a mut PptxContainer,
    current_paths: Vec<String>, // Pfade beim Erstellen des Iterators kopieren
    current_index: usize,
}

impl<'a> SlideIterator<'a> {
    /// Creates a new SlideIterator from a PptxStreamer.
    ///
    /// # Arguments
    ///
    /// * `container` - A mutable reference to a PptxStreamer that will be used to load slides.
    ///
    /// # Returns
    ///
    /// A new SlideIterator instance that will iterate through all slides in the presentation.
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

    /// Advances the iterator and returns the next slide.
    ///
    /// This method loads and processes the next slide from the PPTX file.
    /// It's automatically called when you use the iterator in a for loop.
    ///
    /// # Returns
    ///
    /// * `Some(Ok(Slide))` - The next slide was successfully loaded and processed.
    /// * `Some(Err(_))` - There was an error loading or processing the next slide.
    /// * `None` - There are no more slides to process.
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