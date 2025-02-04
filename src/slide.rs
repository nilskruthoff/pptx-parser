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
}