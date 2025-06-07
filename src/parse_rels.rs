use crate::{Result, Error};
use crate::types::ImageReference;
use roxmltree::Document;

pub fn parse_slide_rels(xml_data: &[u8]) -> Result<Vec<ImageReference>> {
    let xml_str = std::str::from_utf8(xml_data).map_err(|_| Error::Unknown)?;
    let doc = Document::parse(xml_str)?;
    let root = doc.root_element();

    let mut images = Vec::new();
    for rel in root.children().filter(|n| n.is_element() && n.tag_name().name() == "Relationship") {
        if let Some(rel_type) = rel.attribute("Type") {
            if rel_type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" {
                if let Some(id) = rel.attribute("Id") {
                    if let Some(target) = rel.attribute("Target") {
                        images.push(ImageReference {
                            id: id.to_string(),
                            target: target.to_string(),
                        });
                    }
                }
            }
        }
    }

    Ok(images)
}