use crate::Result;
use crate::constants::IMAGE_NAMESPACE;
use crate::types::ImageReference;
use crate::xml::{attr, element_is, event, reader};
use quick_xml::events::Event;

const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Relationship {
    pub id: String,
    pub rel_type: String,
    pub target: String,
}

/// Parses package relationship (`.rels`) XML data and extracts all relationships.
pub fn parse_relationships(xml_data: &[u8]) -> Result<Vec<Relationship>> {
    let mut xml = reader(xml_data);
    let mut relationships = Vec::new();
    let mut depth = 0usize;
    loop {
        match event(&mut xml, "relationships")? {
            Event::Start(element) => {
                depth += 1;
                if element_is(&xml, &element, RELATIONSHIPS_NS, b"Relationship") {
                    push_relationship(&element, &mut relationships);
                }
            }
            Event::Empty(element)
                if element_is(&xml, &element, RELATIONSHIPS_NS, b"Relationship") =>
            {
                push_relationship(&element, &mut relationships);
            }
            Event::End(_) if depth > 0 => depth -= 1,
            Event::Eof if depth > 0 => {
                return Err(crate::Error::ParseError(
                    "Unexpected end of relationships XML",
                ));
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(relationships)
}

fn push_relationship(element: &quick_xml::events::BytesStart<'_>, out: &mut Vec<Relationship>) {
    if let (Some(id), Some(rel_type), Some(target)) = (
        attr(element, b"Id"),
        attr(element, b"Type"),
        attr(element, b"Target"),
    ) {
        out.push(Relationship {
            id,
            rel_type,
            target,
        });
    }
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
