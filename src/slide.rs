﻿use crate::parser_config::ImageHandlingMode;
use crate::{ElementPosition, ImageReference, ParserConfig, SlideElement};
use base64::{engine::general_purpose, Engine as _};
use image::ImageOutputFormat;
use std::collections::HashMap;
use std::fs;
use std::io::Cursor;
use std::path::{Path, PathBuf};

/// Encapsulates images for manual extraction of images from slides
#[derive(Debug)]
pub struct ManualImage {
    pub base64_content: String,
    pub img_ref: ImageReference,
}

impl ManualImage {
    pub fn new(base64_content: String, img_ref: ImageReference) -> ManualImage {
        Self {
            base64_content,
            img_ref,
        }
    }
}

/// Represents a single slide extracted from a PowerPoint (pptx) file.
///
/// Contains structured slide data including slide number, parsed content elements
/// (text, tables, images, lists), and associated image references.
///
/// A `Slide` can be converted into other formats, such as Markdown, or its
/// contained images can be extracted in base64 representation.
///
/// Typically, you retrieve instances of `Slide` through [`PptxContainer::parse()`].
#[derive(Debug)]
pub struct Slide {
    pub rel_path: String,
    pub slide_number: u32,
    pub elements: Vec<SlideElement>,
    pub images: Vec<ImageReference>,
    pub image_data: HashMap<String, Vec<u8>>,
    pub config: ParserConfig
}

impl Slide {
    pub fn new(
        rel_path: String,
        slide_number: u32,
        elements: Vec<SlideElement>,
        images: Vec<ImageReference>,
        image_data: HashMap<String, Vec<u8>>,
        config: ParserConfig,
    ) -> Self {
        Self {
            rel_path,
            slide_number,
            elements,
            images,
            image_data,
            config,
        }
    }

    /// Converts slide contents into a Markdown formatted string.
    ///
    /// Translates internal slide elements (text, tables, lists, images) to valid
    /// and readable Markdown. Embedded images will be encoded as base64 inline images.
    ///
    /// # Returns
    ///
    /// Returns an `Option<String>`:
    /// - `Some(String)`: Markdown representation of slide if conversion succeeds.
    /// - `None`: If a conversion error occurs during image encoding.
    pub fn convert_to_md(&self) -> Option<String> {
        let mut slide_txt = String::new();
        if self.config.include_slide_comment { slide_txt.push_str(format!("<!-- Slide {} -->\n\n", self.slide_number).as_str()); }
        let mut image_count = 0;

        let mut sorted_elements = self.elements.clone();
        sorted_elements.sort_by_key(|element| {
            let ElementPosition { y, x } = element.position();
            (y, x)
        });
        
        for element in sorted_elements {
            match element {
                SlideElement::Text(text, _pos) => {
                    for run in &text.runs {
                        slide_txt.push_str(&run.render_as_md());
                    }
                    slide_txt.push('\n');
                },
                SlideElement::Table(table, _pos) => {
                    let mut is_header = true;
                    for row in &table.rows {
                        let mut row_texts = Vec::new();
                        for cell in &row.cells {
                            let mut cell_text = String::new();
                            for run in &cell.runs {
                                cell_text.push_str(&run.extract());
                            }
                            row_texts.push(cell_text);
                        }

                        let row_line = format!("| {} |", row_texts.join(" | "));
                        slide_txt.push_str(&row_line);
                        slide_txt.push('\n');

                        if is_header {
                            let separator_line = format!("|{}|", row_texts.iter().map(|_| " --- ").collect::<Vec<_>>().join("|"));
                            slide_txt.push_str(&separator_line);
                            slide_txt.push('\n');
                            is_header = false;
                        }
                    }
                    slide_txt.push('\n');
                },
                SlideElement::Image(image_ref, _pos) => {
                    match self.config.image_handling_mode {
                        ImageHandlingMode::InMarkdown => {
                            if let Some(image_data) = self.image_data.get(&image_ref.id) {
                                let image_data = self.config.compress_images
                                    .then(|| self.compress_image(image_data))
                                    .unwrap_or_else(|| Option::from(image_data.clone()));

                                let base64_string = general_purpose::STANDARD.encode(image_data?);
                                let image_name = &image_ref.target.split('/').last()?;
                                let file_ext = &image_name.split('.').last()?;

                                slide_txt.push_str(format!("![{}](data:image/{};base64,{})", image_name, file_ext, base64_string).as_str());
                            }
                        }
                        ImageHandlingMode::Save => {
                            if let Some(image_data) = self.image_data.get(&image_ref.id) {
                                let image_data = self.config.compress_images
                                    .then(|| self.compress_image(image_data))
                                    .unwrap_or_else(|| Option::from(image_data.clone()));

                                let ext = self.config.compress_images
                                    .then(|| "jpg".to_string())
                                    .unwrap_or_else(|| self.get_image_extension(&image_ref.target.clone()));

                                let output_dir = self.config
                                    .image_output_path
                                    .clone()
                                    .unwrap_or_else(|| PathBuf::from("."));

                                let _ = fs::create_dir_all(&output_dir);

                                let mut image_path = output_dir.clone();
                                let file_name = format!("slide{}_image{}_{}.{}", self.slide_number, image_count + 1, &image_ref.id, ext);
                                image_path.push(&file_name);

                                let _ = fs::write(&image_path, image_data?);

                                let abs_file_url = self.path_to_file_url(&image_path);
                                let html_link = format!(r#"<a href={:?}>{file_name}</a>"#, abs_file_url?);
                                image_count += 1;
                                slide_txt.push_str(&html_link);
                                slide_txt.push('\n');
                            }
                        }
                        ImageHandlingMode::Manually => { slide_txt.push('\n'); continue; }
                    }
                    slide_txt.push('\n');
                }
                SlideElement::List(list_element, _pos) => {
                    let mut counters: Vec<usize> = Vec::new();
                    let mut previous_level = 0;

                    for item in &list_element.items {
                        let mut item_text = String::new();
                        for run in &item.runs {
                            item_text.push_str(&run.extract());
                        }

                        let level = item.level as usize;
                        if level >= counters.len() {
                            counters.resize(level + 1, 0);
                        }

                        match level.cmp(&previous_level) {
                            std::cmp::Ordering::Greater => counters[level] = 0,
                            std::cmp::Ordering::Less => counters.truncate(level + 1),
                            std::cmp::Ordering::Equal => {}
                        }

                        counters[level] += 1;
                        previous_level = level;

                        let indent = "\t".repeat(level);
                        let marker = if item.is_ordered {
                            format!("{}{}. ", indent, counters[level])
                        } else {
                            format!("{}- ", indent)
                        };

                        slide_txt.push_str(&format!("{}{}\n", marker, item_text));
                    }
                },
                _ => ()
            }
        }
        Some(slide_txt)
    }

    /// Extracts the numeric slide identifier from a slide path.
    ///
    /// Helper method to parse slide numbers from internal pptx
    /// slide paths (e.g., "ppt/slides/slide1.xml" → `1`).
    pub fn extract_slide_number(path: &str) -> Option<u32> {
        path.split('/')
            .next_back()
            .and_then(|filename| {
                filename
                    .strip_prefix("slide")
                    .and_then(|s| s.strip_suffix(".xml"))
            })
            .and_then(|num_str| num_str.parse::<u32>().ok())
    }

    /// Links slide images references with their corresponding targets.
    ///
    /// Ensures that each image referenced by its ID is correctly 
    /// linked to the actual internal resource paths stored in the slide.
    /// This method is typically used internally after parsing a slide
    ///
    /// # Notes
    ///
    /// Internally those are the values image references are holding
    ///
    /// | Parameter | Example value         |
    /// |---------- |---------------------- |
    /// | `id`      | *rId2*                |
    /// | `target`  | *../media/image2.png* |
    ///
    pub fn link_images(&mut self) {
        let id_to_target: HashMap<String, String> = self.images
            .iter()
            .map(|img_ref| (img_ref.id.clone(), img_ref.target.clone()))
            .collect();

        for element in &mut self.elements {
            if let SlideElement::Image(ref mut img_ref, _pos) = element {
                if let Some(target) = id_to_target.get(&img_ref.id) {
                    img_ref.target = target.clone();
                }
            }
        }
    }

    /// Extracts the file extension from image paths
    pub fn get_image_extension(&self, path: &str) -> String {
        Path::new(path)
            .extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or("bin")
            .to_string()
    }

    /// Compresses the image data and returning it as a `jpg` byte slice
    /// 
    /// # Parameter
    /// 
    /// - `image_data`: The raw image data as a byte array
    /// 
    /// # Returns
    /// 
    /// - `Vec<u8>`: Returns the compressed and converted jpg byte array
    ///
    /// # Notes
    ///
    /// All images will be converted to `jpg`
    pub fn compress_image(&self, image_data: &[u8]) -> Option<Vec<u8>> {
        let img = match image::load_from_memory(image_data) {
            Ok(image) => image,
            Err(_) => return None,
        };

        let mut output = Vec::new();
        let quality = self.config.quality;

        if img.write_to(&mut Cursor::new(&mut output), ImageOutputFormat::Jpeg(quality)).is_ok() {
            Some(output)
        } else {
            None
        }
    }
    
    pub fn load_images_manually(&self) -> Option<Vec<ManualImage>> {
        let mut images: Vec<ManualImage> = Vec::new();
        
        let image_refs: Vec<&ImageReference> = self.elements
            .iter()
            .filter_map(|element| match element {
                SlideElement::Image(ref img, _pos) => Some(img),
                _ => None,
            })
            .collect();
        
        for image_ref in image_refs {
            if let Some(image_data) = self.image_data.get(&image_ref.id) {
                let image_data = self.config.compress_images
                    .then( | | self.compress_image(image_data))
                    .unwrap_or_else(|| Option::from(image_data.clone()));

                let base64_str = general_purpose::STANDARD.encode(image_data?);
                
                let image = ManualImage::new(
                    base64_str,
                    image_ref.clone(),
                );
                images.push(image);
            }
        }
        
        Some(images)
    }

    fn path_to_file_url(&self, path: &Path) -> Option<String> {
        let abs_path = path.canonicalize().ok()?;
        let mut path_str = abs_path.to_string_lossy().replace('\\', "/");

        // remove windows unc prefix
        if cfg!(windows) {
            if let Some(stripped) = path_str.strip_prefix("//?/") {
                path_str = stripped.to_string();
            }
            Some(format!("file:///{}", path_str))
        } else {
            Some(format!("file://{}", path_str))
        }
    }
}

#[cfg(test)]
mod tests {
    use std::fs;
    use std::path::PathBuf;
    use crate::ElementPosition;
    use super::*;

    fn mock_slide() -> Slide {
        Slide {
            rel_path: "ppt/slides/slide1.xml".to_string(),
            slide_number: 1,
            elements: vec![],
            images: vec![],
            image_data: HashMap::new(),
            config: ParserConfig::default(),
        }
    }

    fn load_image_data(filename: &str) -> Vec<u8> {
        let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        path.push("tests");
        path.push("test_data");
        path.push(filename);
        fs::read(path).expect("Unable to read test data file")
    }
    
    #[test]
    fn test_extract_slide_number() {
        let input = "ppt/slides/slide5.xml";
        
        let actual = Slide::extract_slide_number(input).unwrap();
        let expected: u32 = 5;
        
        assert_eq!(actual, expected);
    }
    
    #[test]
    fn test_get_image_extension() {
        let slide = mock_slide();
        let input = "../media/image1.png";
        
        let actual = slide.get_image_extension(input);
        let expected = "png";
        
        assert_eq!(actual, expected);
    }

    #[test]
    fn test_link_images() {
        let mut slide = mock_slide();
        let _position = ElementPosition::default();
        
        slide.images.push(ImageReference { id: "rId2".to_string(), target: "../media/image1.png".to_string() });
        slide.elements.push(SlideElement::Image(ImageReference { id: "rId2".to_string(), target: "".to_string() }, _position));

        slide.link_images();

        if let SlideElement::Image(img_ref, _postion) = &slide.elements[0] {
            assert_eq!(img_ref.target, "../media/image1.png");
        }
    }

    #[test]
    fn test_image_compression_reduces_size() {
        let mut slide = mock_slide();
        slide.config.quality = 50;

        let raw_image = load_image_data("example-image.jpg");

        if let Some(compression_result) = slide.compress_image(&raw_image) {
            assert!(compression_result.len() < raw_image.len());
        } else {
            panic!("Compression failed");
        }
    }

    #[test]
    fn test_compressed_image_is_valid_jpg() {
        let slide = mock_slide();
        let raw_image = load_image_data("example-image.jpg");

        if let Some(compression_result) = slide.compress_image(&raw_image) {
            let result = image::load_from_memory(&compression_result);
            assert!(result.is_ok());
        } else {
            panic!("Compression failed");
        }
    }
}