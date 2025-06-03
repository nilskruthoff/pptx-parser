use std::string::ParseError;
use crate::SlideElement;
use crate::parse_xml;
use super::{Error, };

#[derive(Debug)]
pub struct Slide {
    pub rel_path: String,
    pub slide_number: u32,
    pub elements: Vec<SlideElement>
}

impl Slide {
    pub fn parse(xml: &[u8], rel_path: String) -> Result<Slide, Error> {
        let slide_number = Self::extract_slide_number(&rel_path).unwrap_or(0);
        let elements: Vec<SlideElement> = parse_xml::parse_slide_xml(&xml)?;
        Ok(Slide { rel_path, slide_number, elements })
    }

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

    pub fn extract_text(&self) -> Option<String> {
        let mut slide_txt = String::new();

        for element in &self.elements {
            match element {
                SlideElement::Text(text) => {
                    for run in &text.runs {
                        slide_txt.push_str(&run.extract());
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

                        // Erstelle die Markdown-Zeile
                        let row_line = format!("| {} |", row_texts.join(" | "));
                        slide_txt.push_str(&row_line);
                        slide_txt.push('\n');

                        // Füge nach der ersten Zeile den Trenner hinzu
                        if is_header {
                            let separator_line = format!("|{}|", row_texts.iter().map(|_| " --- ").collect::<Vec<_>>().join("|"));
                            slide_txt.push_str(&separator_line);
                            slide_txt.push('\n');
                            is_header = false;
                        }
                    }
                    slide_txt.push('\n');
                },
                _ => ()
            }
        }
        Some(slide_txt)
    }
}