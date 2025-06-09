use super::Error;
use crate::parse_xml;
use crate::{parse_rels, ImageReference, SlideElement};
use base64::{engine::general_purpose, Engine as _};
use std::collections::HashMap;


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
pub struct Slide<'a> {
    pub rel_path: String,
    pub slide_number: u32,
    pub elements: Vec<SlideElement>,
    pub images: Vec<ImageReference>,
    pub files: &'a HashMap<String, Vec<u8>>,
}

impl<'a> Slide<'a> {
    /// Parses raw slide XML-data and relationships to create a structured [`Slide`] instance.
    ///
    /// # Arguments
    ///
    /// - `xml`: Raw XML byte slice representing slide information.
    /// - `rel_path`: The internal relationship path of the slide.
    /// - `rels_data`: Optional relationships XML data (`.rels`) associated with slide.
    /// - `files`: A reference to the pptx file content map, for resource lookup.
    ///
    /// # Returns
    ///
    /// Returns a `Result`:
    /// - `Ok(Slide)`: Fully structured slide upon successful parsing.
    /// - `Err(Error)`: If XML parsing or slide building fails.
    ///
    /// # Errors
    ///
    /// Parsing may fail if XML structure is malformed or critical data is missing.
    pub fn parse(
        xml: &[u8],
        rel_path: String,
        rels_data: Option<&[u8]>,
        files: &'a HashMap<String, Vec<u8>>,
    ) -> Result<Slide<'a>, Error> {
        let slide_number = Self::extract_slide_number(&rel_path).unwrap_or(0);
        let elements: Vec<SlideElement> = parse_xml::parse_slide_xml(xml)?;

        let images = if let Some(rels_data) = rels_data {
            parse_rels::parse_slide_rels(rels_data)?
        } else {
            Vec::new()
        };

        Ok(Slide { rel_path, slide_number, elements, images, files })
    }

    /// Extracts the numeric slide identifier from a slide path.
    ///
    /// Helper method to parse slide numbers from internal pptx
    /// slide paths (e.g., "ppt/slides/slide1.xml" → `1`).
    fn extract_slide_number(path: &str) -> Option<u32> {
        path
            .split('/')
            .last()
            .and_then(|filename| {
                filename
                    .strip_prefix("slide")
                    .and_then(|s| s.strip_suffix(".xml"))
            })
            .and_then(|num_str| num_str.parse::<u32>().ok())
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
        slide_txt.push_str(format!("<!-- Slide {} -->\n\n", self.slide_number).as_str());

        for element in &self.elements {
            match element {
                SlideElement::Text(text) => {
                    for run in &text.runs {
                        slide_txt.push_str(&run.render_as_md());
                    }
                    slide_txt.push('\n');
                },
                SlideElement::Table(table) => {
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
                SlideElement::Image(image_ref) => {
                    let image_path = self.get_full_image_path(&image_ref.target);
                    
                    if let Some(image_data) = self.files.get(&image_path) {
                        let base64_string = general_purpose::STANDARD.encode(image_data);
                        let image_name = &image_ref.target.split('/').last()?;
                        let file_ext = &image_name.split('.').last()?;
                        slide_txt.push_str(format!("![{}](data:image/{};base64,{}\n)", image_name, file_ext, base64_string).as_str());
                    }
                    
                    slide_txt.push('\n');
                },
                SlideElement::List(list_element) => {
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

                        if level > previous_level {
                            counters[level] = 0;
                        } else if level < previous_level {
                            counters.truncate(level + 1);
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
    
    pub fn extract_images_as_base64(&self) -> Option<Vec<String>> {
        let mut images_base64 = Vec::new();

        for element in &self.elements {
            if let SlideElement::Image(image_ref) = element {
                let image_path = self.get_full_image_path(&image_ref.target);

                if let Some(image_data) = self.files.get(&image_path) {
                    let base64_string = general_purpose::STANDARD.encode(image_data);
                    images_base64.push(base64_string);
                } else {
                    return None;
                }
            }
        }

        Some(images_base64)
    }

    fn get_full_image_path(&self, target: &str) -> String {
        if target.starts_with("../") {
            let adjusted_target = target.trim_start_matches("../");
            format!("ppt/{}", adjusted_target)
        } else {
            format!("ppt/slides/{}", target)
        }
    }

    /// Links slide images references with their corresponding targets.
    ///
    /// Ensures that each image referenced by its ID is correctly 
    /// linked to the actual internal resource paths stored in the slide.
    /// This method is typically used internally after parsing a slide
    pub fn link_images(&mut self) {
        let id_to_target: HashMap<String, String> = self.images
            .iter()
            .map(|img_ref| (img_ref.id.clone(), img_ref.target.clone()))
            .collect();

        for element in &mut self.elements {
            if let SlideElement::Image(ref mut img_ref) = element {
                if let Some(target) = id_to_target.get(&img_ref.id) {
                    img_ref.target = target.clone();
                }
            }
        }
    }
}