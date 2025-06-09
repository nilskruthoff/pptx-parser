use crate::{Result, Error};
use crate::types::ImageReference;
use roxmltree::Document;
use crate::constants::IMAGE_NAMESPACE;

/// Parses relationship (`.rels`) XML data from a PPTX slide, extracting image references.
///
/// PowerPoint slide relationships data contain mappings between resource IDs and their targets.
/// This function specifically extracts relationships pointing to embedded images.
///
/// # Arguments
///
/// - `xml_data`: Raw relationship XML data as a byte slice.
///
/// # Returns
///
/// Returns a `Result` containing:
/// - `Ok(Vec<ImageReference>)`: Vector containing extracted image IDs and their corresponding target paths.
/// - `Err(Error)`: If provided XML data isn't valid UTF-8 or XML parsing fails.
///
/// # Errors
///
/// An error is returned if:
/// - The XML data is not valid UTF-8.
/// - Malformed or invalid XML structure is detected.
/// ```
pub fn parse_slide_rels(xml_data: &[u8]) -> Result<Vec<ImageReference>> {
    let xml_str = std::str::from_utf8(xml_data).map_err(|_| Error::Unknown)?;
    let doc = Document::parse(xml_str)?;
    let root = doc.root_element();

    let mut images = Vec::new();
    for rel in root.children().filter(|n| n.is_element() && n.tag_name().name() == "Relationship") {
        if let Some(rel_type) = rel.attribute("Type") {
            if rel_type == IMAGE_NAMESPACE {
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