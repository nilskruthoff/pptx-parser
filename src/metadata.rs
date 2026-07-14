use crate::{Error, Result, Slide};
use roxmltree::{Document, Node};

const CORE_PROPERTIES_NS: &str = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const DC_NS: &str = "http://purl.org/dc/elements/1.1/";
const DCTERMS_NS: &str = "http://purl.org/dc/terms/";
const META_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";

/// The common document properties of a presentation.
///
/// These values belong to the presentation as a whole rather than to an
/// individual slide. Timestamps are preserved in their source representation.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PresentationMetadata {
    pub title: Option<String>,
    pub author: Option<String>,
    pub last_modified_by: Option<String>,
    pub subject: Option<String>,
    pub description: Option<String>,
    pub keywords: Vec<String>,
    pub created_at: Option<String>,
    pub modified_at: Option<String>,
}

pub(crate) fn parse_pptx_metadata(core_xml: Option<&[u8]>) -> Result<PresentationMetadata> {
    let mut metadata = PresentationMetadata::default();

    if let Some(xml) = core_xml {
        let xml = std::str::from_utf8(xml)?;
        let document = Document::parse(xml)?;
        metadata.title = element_text(&document, DC_NS, "title");
        metadata.author = element_text(&document, DC_NS, "creator");
        metadata.last_modified_by = element_text(&document, CORE_PROPERTIES_NS, "lastModifiedBy");
        metadata.subject = element_text(&document, DC_NS, "subject");
        metadata.description = element_text(&document, DC_NS, "description");
        metadata.keywords = element_text(&document, CORE_PROPERTIES_NS, "keywords")
            .into_iter()
            .collect();
        metadata.created_at = element_text(&document, DCTERMS_NS, "created");
        metadata.modified_at = element_text(&document, DCTERMS_NS, "modified");
    }

    Ok(metadata)
}

pub(crate) fn parse_odp_metadata(meta_xml: Option<&[u8]>) -> Result<PresentationMetadata> {
    let mut metadata = PresentationMetadata::default();

    if let Some(xml) = meta_xml {
        let xml = std::str::from_utf8(xml)?;
        let document = Document::parse(xml)?;
        let last_modified_by = element_text(&document, DC_NS, "creator");
        metadata.title = element_text(&document, DC_NS, "title");
        metadata.author = element_text(&document, META_NS, "initial-creator")
            .or_else(|| last_modified_by.clone());
        metadata.last_modified_by = last_modified_by;
        metadata.subject = element_text(&document, DC_NS, "subject");
        metadata.description = element_text(&document, DC_NS, "description");
        metadata.keywords = document
            .descendants()
            .filter(|node| is_element(*node, META_NS, "keyword"))
            .filter_map(node_text)
            .collect();
        metadata.created_at = element_text(&document, META_NS, "creation-date");
        metadata.modified_at = element_text(&document, DC_NS, "date");
    }

    Ok(metadata)
}

pub(crate) fn render_presentation_markdown(
    metadata: &PresentationMetadata,
    include_metadata: bool,
    slides: Vec<Slide>,
) -> Result<String> {
    let mut parts = Vec::with_capacity(slides.len() + usize::from(include_metadata));
    if include_metadata {
        if let Some(comment) = render_metadata_comment(metadata) {
            parts.push(comment);
        }
    }
    for slide in slides {
        parts.push(slide.convert_to_md().ok_or(Error::ConversionFailed)?);
    }
    Ok(parts.join("\n\n"))
}

fn render_metadata_comment(metadata: &PresentationMetadata) -> Option<String> {
    let mut fields = Vec::new();
    push_field(&mut fields, "Title", metadata.title.as_deref());
    push_field(&mut fields, "Author", metadata.author.as_deref());
    push_field(
        &mut fields,
        "Last Modified By",
        metadata.last_modified_by.as_deref(),
    );
    push_field(&mut fields, "Subject", metadata.subject.as_deref());
    push_field(&mut fields, "Description", metadata.description.as_deref());
    if !metadata.keywords.is_empty() {
        fields.push(format!(
            "Keywords: {}",
            sanitize_comment_value(&metadata.keywords.join("; "))
        ));
    }
    push_field(&mut fields, "Created", metadata.created_at.as_deref());
    push_field(&mut fields, "Modified", metadata.modified_at.as_deref());
    if fields.is_empty() {
        None
    } else {
        Some(format!(
            "<!-- Presentation Metadata\n{}\n-->",
            fields.join("\n")
        ))
    }
}

fn push_field(fields: &mut Vec<String>, label: &str, value: Option<&str>) {
    if let Some(value) = value {
        fields.push(format!("{label}: {}", sanitize_comment_value(value)));
    }
}

fn sanitize_comment_value(value: &str) -> String {
    value
        .split_whitespace()
        .collect::<Vec<_>>()
        .join(" ")
        .replace("--", "&#45;&#45;")
}

fn element_text(document: &Document<'_>, namespace: &str, name: &str) -> Option<String> {
    document
        .descendants()
        .find(|node| is_element(*node, namespace, name))
        .and_then(node_text)
}

fn node_text(node: Node<'_, '_>) -> Option<String> {
    node.text()
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .map(str::to_string)
}

fn is_element(node: Node<'_, '_>, namespace: &str, name: &str) -> bool {
    node.is_element()
        && node.tag_name().namespace() == Some(namespace)
        && node.tag_name().name() == name
}

#[cfg(test)]
#[path = "../tests/unit/metadata.rs"]
mod tests;
