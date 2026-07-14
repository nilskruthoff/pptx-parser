use crate::constants::IMAGE_NAMESPACE;
use crate::types::ImageReference;
use crate::{Error, Result};
use roxmltree::Document;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Relationship {
    pub id: String,
    pub rel_type: String,
    pub target: String,
}

/// Parses package relationship (`.rels`) XML data and extracts all relationships.
pub fn parse_relationships(xml_data: &[u8]) -> Result<Vec<Relationship>> {
    let xml_str = std::str::from_utf8(xml_data).map_err(|_| Error::Unknown)?;
    let doc = Document::parse(xml_str)?;
    let root = doc.root_element();

    let mut relationships = Vec::new();
    for rel in root.children().filter(|n| n.is_element() && n.tag_name().name() == "Relationship") {
        if let (Some(id), Some(rel_type), Some(target)) = (
            rel.attribute("Id"),
            rel.attribute("Type"),
            rel.attribute("Target"),
        ) {
            relationships.push(Relationship {
                id: id.to_string(),
                rel_type: rel_type.to_string(),
                target: target.to_string(),
            });
        }
    }

    Ok(relationships)
}

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
    Ok(parse_relationships(xml_data)?
        .into_iter()
        .filter(|rel| rel.rel_type == IMAGE_NAMESPACE)
        .map(|rel| ImageReference {
            id: rel.id,
            target: rel.target,
        })
        .collect())
}

/// Extracts external hyperlink targets keyed by their relationship ID.
pub fn parse_hyperlink_rels(xml_data: &[u8]) -> Result<std::collections::HashMap<String, String>> {
    Ok(parse_relationships(xml_data)?
        .into_iter()
        .filter(|rel| rel.rel_type == crate::constants::HYPERLINK_NAMESPACE)
        .map(|rel| (rel.id, rel.target))
        .collect())
}

#[cfg(test)]
#[path = "../tests/unit/parse_rels.rs"]
mod tests;
