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

#[cfg(test)]
mod tests {
    use std::fs;
    use std::path::PathBuf;
    use super::*;

    fn load_xml(filename: &str) -> Vec<u8> {
        let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        path.push("tests");
        path.push("test_data");
        path.push("xml");
        path.push(filename);
        fs::read(path).expect("Unable to read test data file")
    }

    fn normalize_test_string(input: &str) -> String {
        input
            .trim_start_matches('\u{feff}') // remove BOM
            .replace("\r\n", "\n") // normalize line breaks
            .replace("    ", "\t") // replace 4 whitespaces with a tab
            .trim() // trim leading and trailing whitespace
            .to_string()
    }

    #[test]
    fn test_parse_slide_rels_with_images() {
        let xml_data = load_xml("rels_with_images.xml");
        match parse_slide_rels(&xml_data) {
            Ok(images) => {
                assert_eq!(images.len(), 2);
                assert_eq!(images[0].id, "rId1");
                assert_eq!(normalize_test_string(&images[0].target), normalize_test_string("../media/image1.png"));
                assert_eq!(images[1].id, "rId2");
                assert_eq!(normalize_test_string(&images[1].target), normalize_test_string("../media/image2.jpg"));
            },
            Err(_) => panic!("Fehler beim Parsen der Slide-Relationships mit Bildern")
        }
    }

    #[test]
    fn test_parse_slide_rels_empty() {
        let xml_data = load_xml("rels_without_images.xml");
        match parse_slide_rels(&xml_data) {
            Ok(images) => {
                assert_eq!(images.len(), 0);
            },
            Err(_) => panic!("Fehler beim Parsen der leeren Slide-Relationships")
        }
    }
}
