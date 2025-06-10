use crate::{ImageReference, SlideElement};
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
pub struct Slide {
    pub rel_path: String,
    pub slide_number: u32,
    pub elements: Vec<SlideElement>,
    pub images: Vec<ImageReference>,
    pub image_data: HashMap<String, Vec<u8>>,
}

impl Slide {
    pub fn new(
        rel_path: String,
        slide_number: u32,
        elements: Vec<SlideElement>,
        images: Vec<ImageReference>,
        image_data: HashMap<String, Vec<u8>>,
    ) -> Self {
        Self {
            rel_path,
            slide_number,
            elements,
            images,
            image_data,
        }
    }

    // Extrahiere Slide-Nummer aus Pfad (statische Methode)
    pub fn extract_slide_number(path: &str) -> Option<u32> {
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

    // Link Bilder zu den entsprechenden Elementen
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

    // Konvertiert den Slide zu Markdown
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
                    // Table handling (wie vorher)
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
                    // Bild aus den vorab geladenen Daten holen
                    if let Some(image_data) = self.image_data.get(&image_ref.id) {
                        let base64_string = general_purpose::STANDARD.encode(image_data);
                        let image_name = &image_ref.target.split('/').last()?;
                        let file_ext = &image_name.split('.').last()?;
                        slide_txt.push_str(format!("![{}](data:image/{};base64,{})",
                                                   image_name, file_ext, base64_string).as_str());
                    }
                    slide_txt.push('\n');
                },
                SlideElement::List(list_element) => {
                    // List handling (wie vorher)
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
}