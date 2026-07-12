use super::{Result, Slide};
use crate::constants::{COMMENTS_NAMESPACE, NOTES_SLIDE_NAMESPACE, SLIDE_LAYOUT_NAMESPACE, SLIDE_MASTER_NAMESPACE};
use crate::parse_rels::parse_relationships;
use crate::parse_xml::{extract_inherited_positions, InheritedPositions};
use crate::parser_config::ParserConfig;
use rayon::prelude::*;
use std::{
    collections::HashMap,
    io::Read,
    path::Path,
};
use std::sync::Arc;

/// Holds the internal representation of a loaded PowerPoint (pptx) container.
///
/// `PptxContainer` provides functionalities for accessing slides and their resources
/// directly from a loaded pptx file. It parses and stores XML slides content,
/// relationships (`rels`) files, and associated resources such as images.
pub struct PptxContainer {
    pub config: ParserConfig,
    archive: zip::ZipArchive<std::fs::File>,
    pub slide_paths: Vec<String>,
    pub slide_count: u32,
}

impl PptxContainer {
    /// Opens a PowerPoint pptx file and initializes a `PptxContainer`.
    /// 
    /// _For new code that should support both `.pptx` and `.odp`, prefer
    /// [`PresentationContainer`]. This type remains available for PPTX-only
    /// workflows and backwards compatibility._
    ///
    /// Processes the given file, extracting its internal files into memory. After initialization, the 
    /// container holds slide XML data, relationship files (*.rels), and associated resources.
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
        let mut slide_count = 0;

        for i in 0..archive.len() {
            let file = archive.by_index(i)?;
            let name = file.name().to_string();

            if name.starts_with("ppt/slides/slide") && name.ends_with(".xml") {
                slide_paths.push(name);
                slide_count += 1;
            }
        }

        sort_slide_paths(&mut slide_paths);

        Ok(Self { archive, slide_paths, config, slide_count })
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

    /// Parses all slides in the presentation with optimized multithreaded processing.
    ///
    /// This method uses Rayon for parallel processing by:
    /// 1. Preloading all necessary data sequentially (I/O-bound operations)
    /// 2. Performing CPU-intensive XML parsing in parallel
    /// 3. Using shared references for thread-safe data access
    ///
    /// # Returns
    ///
    /// * `Result<Vec<Slide>>` - List of all parsed slides
    pub fn parse_all_multi_threaded(&mut self) -> Result<Vec<Slide>> {
        // Clone paths upfront to avoid holding reference to self
        let slide_paths = self.slide_paths.clone();
        let config = self.config.clone();
        let mut raw_data = Vec::with_capacity(slide_paths.len());
        let mut all_image_data = HashMap::new();

        for slide_path in &slide_paths {
            // Read slide XML and relationships
            let slide_xml = self.read_file_from_archive(slide_path)?;
            let rels_path = self.get_slide_rels_path(slide_path);
            let rels_data = self.read_file_from_archive(&rels_path).ok();
            let slide_number = Slide::extract_slide_number(slide_path).unwrap_or(0);
            let inherited_positions = self.resolve_inherited_positions(slide_path, rels_data.as_deref())?;
            let speaker_notes = self.resolve_speaker_notes(slide_path, rels_data.as_deref())?;
            let comments = self.resolve_comments(slide_path, rels_data.as_deref())?;

            // Preload images if enabled
            let mut slide_images = Vec::new();
            if config.extract_images {
                if let Some(ref data) = rels_data {
                    slide_images = crate::parse_rels::parse_slide_rels(data)?;
                }

                for img_ref in &slide_images {
                    let path = PptxContainer::resolve_target_path(slide_path, &img_ref.target);
                    let data = self.read_file_from_archive(&path)?;
                    all_image_data.entry(img_ref.target.clone()).or_insert(data);
                }
            }

            raw_data.push((slide_path.clone(), slide_number, slide_xml, slide_images, inherited_positions, speaker_notes, comments));
        }

        // Share image data atomically across threads
        let shared_image_data = Arc::new(all_image_data);

        // Parallel processing starts here (CPU-bound tasks)
        let slides: Result<Vec<_>> = raw_data
            .into_par_iter()
            .map(|(path, number, xml, images, inherited_positions, speaker_notes, comments)| {
                // Parse XML in parallel (CPU-intensive)
                let elements = crate::parse_xml::parse_slide_xml_with_inherited_positions(&xml, &inherited_positions)?;

                // Resolve image data from shared registry
                let mut image_map = HashMap::new();
                if config.extract_images {
                    for img_ref in &images {
                        if let Some(data) = shared_image_data.get(&img_ref.target) {
                            image_map.insert(img_ref.id.clone(), data.clone());
                        }
                    }
                }

                // Build slide
                let mut slide = Slide::new(
                    path,
                    number,
                    elements,
                    speaker_notes,
                    comments,
                    images,
                    image_map,
                    config.clone(),
                );
                slide.link_images();
                Ok(slide)
            })
            .collect();

        slides
    }

    
    pub fn iter_slides(&mut self) -> SlideIterator<'_> {
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
    pub fn load_slide(&mut self, slide_path: &str) -> Result<Option<Slide>> {
        // load xml data
        let slide_data = self.read_file_from_archive(slide_path)?;

        // load relationship file
        let rels_path = self.get_slide_rels_path(slide_path);
        let rels_data = self.read_file_from_archive(&rels_path).ok();

        // parse slide and preload images
        let slide_number = Slide::extract_slide_number(slide_path).unwrap_or(0);
        let inherited_positions = self.resolve_inherited_positions(slide_path, rels_data.as_deref())?;
        let speaker_notes = self.resolve_speaker_notes(slide_path, rels_data.as_deref())?;
        let comments = self.resolve_comments(slide_path, rels_data.as_deref())?;
        let elements = crate::parse_xml::parse_slide_xml_with_inherited_positions(&slide_data, &inherited_positions)?;
        
        let mut images = Vec::new();
        let mut image_data = HashMap::new();
        
        if self.config.extract_images {
            // extract images from relationships
            if let Some(ref rels_bytes) = rels_data {
                images = crate::parse_rels::parse_slide_rels(rels_bytes)?;
            }

            for img_ref in &images {
                let img_path = Self::resolve_target_path(slide_path, &img_ref.target);
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
            speaker_notes,
            comments,
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
    pub fn read_file_from_archive(&mut self, path: &str) -> Result<Vec<u8>> {
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
    pub fn get_slide_rels_path(&self, slide_path: &str) -> String {
        let mut rels_path = slide_path.to_string();
        if let Some(pos) = rels_path.rfind('/') {
            rels_path.insert_str(pos + 1, "_rels/");
        }
        rels_path.push_str(".rels");
        rels_path
    }

    fn resolve_inherited_positions(
        &mut self,
        slide_path: &str,
        slide_rels_data: Option<&[u8]>,
    ) -> Result<InheritedPositions> {
        let Some(slide_rels_data) = slide_rels_data else {
            return Ok(InheritedPositions::default());
        };

        let slide_relationships = parse_relationships(slide_rels_data)?;
        let Some(layout_target) = slide_relationships
            .iter()
            .find(|rel| rel.rel_type == SLIDE_LAYOUT_NAMESPACE)
            .map(|rel| rel.target.as_str()) else {
            return Ok(InheritedPositions::default());
        };

        let layout_path = Self::resolve_target_path(slide_path, layout_target);
        let layout_xml = self.read_file_from_archive(&layout_path)?;
        let layout_rels_path = self.get_slide_rels_path(&layout_path);
        let layout_rels_data = self.read_file_from_archive(&layout_rels_path).ok();

        let master_positions = if let Some(layout_rels_data) = layout_rels_data.as_deref() {
            let layout_relationships = parse_relationships(layout_rels_data)?;
            if let Some(master_target) = layout_relationships
                .iter()
                .find(|rel| rel.rel_type == SLIDE_MASTER_NAMESPACE)
                .map(|rel| rel.target.as_str()) {
                let master_path = Self::resolve_target_path(&layout_path, master_target);
                let master_xml = self.read_file_from_archive(&master_path)?;
                extract_inherited_positions(&master_xml, &InheritedPositions::default())?
            } else {
                InheritedPositions::default()
            }
        } else {
            InheritedPositions::default()
        };

        extract_inherited_positions(&layout_xml, &master_positions)
    }

    fn resolve_speaker_notes(
        &mut self,
        slide_path: &str,
        slide_rels_data: Option<&[u8]>,
    ) -> Result<Vec<crate::TextElement>> {
        let Some(slide_rels_data) = slide_rels_data else {
            return Ok(Vec::new());
        };
        let relationships = parse_relationships(slide_rels_data)?;
        let Some(notes_target) = relationships
            .iter()
            .find(|rel| rel.rel_type == NOTES_SLIDE_NAMESPACE)
            .map(|rel| rel.target.as_str()) else {
                return Ok(Vec::new());
            };
        let notes_path = Self::resolve_target_path(slide_path, notes_target);
        let notes_xml = self.read_file_from_archive(&notes_path)?;
        crate::parse_xml::parse_speaker_notes_xml(&notes_xml)
    }

    fn resolve_comments(
        &mut self,
        slide_path: &str,
        slide_rels_data: Option<&[u8]>,
    ) -> Result<Vec<crate::TextElement>> {
        let Some(slide_rels_data) = slide_rels_data else {
            return Ok(Vec::new());
        };
        let relationships = parse_relationships(slide_rels_data)?;
        let Some(comment_target) = relationships
            .iter()
            .find(|rel| rel.rel_type == COMMENTS_NAMESPACE || rel.rel_type.ends_with("/comments"))
            .map(|rel| rel.target.as_str()) else {
                return Ok(Vec::new());
            };
        let comment_path = Self::resolve_target_path(slide_path, comment_target);
        let comment_xml = self.read_file_from_archive(&comment_path)?;
        crate::parse_xml::parse_comments_xml(&comment_xml)
    }

    pub fn resolve_target_path(base_path: &str, target: &str) -> String {
        let mut parts: Vec<&str> = if target.starts_with('/') {
            Vec::new()
        } else {
            let mut parts: Vec<&str> = base_path.split('/').collect();
            let _ = parts.pop();
            parts
        };

        for part in target.split('/') {
            match part {
                "" | "." => {}
                ".." => {
                    let _ = parts.pop();
                }
                _ => parts.push(part),
            }
        }

        parts.join("/")
    }
}

fn sort_slide_paths(slide_paths: &mut [String]) {
    slide_paths.sort_by(|left, right| {
        Slide::extract_slide_number(left)
            .cmp(&Slide::extract_slide_number(right))
            .then_with(|| left.cmp(right))
    });
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;
    use std::io::Write;
    use zip::write::SimpleFileOptions;

    #[test]
    fn test_sort_slide_paths_numerically() {
        let mut slide_paths = vec![
            "ppt/slides/slide10.xml".to_string(),
            "ppt/slides/slide2.xml".to_string(),
            "ppt/slides/slide1.xml".to_string(),
        ];

        sort_slide_paths(&mut slide_paths);

        assert_eq!(slide_paths, vec![
            "ppt/slides/slide1.xml".to_string(),
            "ppt/slides/slide2.xml".to_string(),
            "ppt/slides/slide10.xml".to_string(),
        ]);
    }

    #[test]
    fn test_resolve_target_path_with_parent_segments() {
        let resolved = PptxContainer::resolve_target_path(
            "ppt/slides/slide3.xml",
            "../slideLayouts/slideLayout3.xml",
        );

        assert_eq!(resolved, "ppt/slideLayouts/slideLayout3.xml");
    }

    #[test]
    fn test_resolve_target_path_with_root_relative_target() {
        let resolved = PptxContainer::resolve_target_path(
            "ppt/slides/slide3.xml",
            "/ppt/slideLayouts/slideLayout3.xml",
        );

        assert_eq!(resolved, "ppt/slideLayouts/slideLayout3.xml");
    }

    #[test]
    fn loads_speaker_notes_from_a_slide_relationship() {
        let path = std::env::temp_dir().join(format!(
            "pptx-to-md-speaker-notes-{}.pptx",
            std::process::id()
        ));
        let file = fs::File::create(&path).expect("create temporary PPTX");
        let mut archive = zip::ZipWriter::new(file);
        let options = SimpleFileOptions::default();

        archive
            .start_file("ppt/slides/slide1.xml", options)
            .expect("start slide entry");
        archive
            .write_all(br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>"#)
            .expect("write slide entry");
        archive
            .start_file("ppt/slides/_rels/slide1.xml.rels", options)
            .expect("start relationship entry");
        archive
            .write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/></Relationships>"#)
            .expect("write relationship entry");
        archive
            .start_file("ppt/notesSlides/notesSlide1.xml", options)
            .expect("start notes entry");
        archive
            .write_all(br#"<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>Presenter detail</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>"#)
            .expect("write notes entry");
        archive.finish().expect("finish temporary PPTX");

        let mut container = PptxContainer::open(&path, ParserConfig::default())
            .expect("open temporary PPTX");
        let slides = container.parse_all().expect("parse temporary PPTX");
        assert_eq!(slides.len(), 1);
        assert_eq!(slides[0].speaker_notes.len(), 1);
        assert_eq!(slides[0].speaker_notes[0].runs[0].text, "Presenter detail\n");

        drop(container);
        fs::remove_file(path).expect("remove temporary PPTX");
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
