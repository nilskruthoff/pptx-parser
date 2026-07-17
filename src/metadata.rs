use crate::xml::{element_is, end_is, event, reader, reference, text};
use crate::{Error, Result, Slide};
use quick_xml::events::Event;

const CORE_PROPERTIES_NS: &str =
    "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const DC_NS: &str = "http://purl.org/dc/elements/1.1/";
const DCTERMS_NS: &str = "http://purl.org/dc/terms/";
const META_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";

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
    core_xml.map_or_else(
        || Ok(PresentationMetadata::default()),
        |xml| parse_metadata(xml, MetadataKind::Pptx, "PPTX core properties"),
    )
}

pub(crate) fn parse_odp_metadata(meta_xml: Option<&[u8]>) -> Result<PresentationMetadata> {
    meta_xml.map_or_else(
        || Ok(PresentationMetadata::default()),
        |xml| parse_metadata(xml, MetadataKind::Odp, "ODP metadata"),
    )
}

#[derive(Clone, Copy)]
enum MetadataKind {
    Pptx,
    Odp,
}

fn parse_metadata(data: &[u8], kind: MetadataKind, part: &str) -> Result<PresentationMetadata> {
    let mut xml = reader(data);
    let mut metadata = PresentationMetadata::default();
    let mut depth = 0usize;
    loop {
        match event(&mut xml, part)? {
            Event::Start(element) => {
                let field = if element_is(&xml, &element, DC_NS, b"title") {
                    Some(MetadataField::Title)
                } else if element_is(&xml, &element, DC_NS, b"creator") {
                    Some(MetadataField::Creator)
                } else if element_is(&xml, &element, DC_NS, b"subject") {
                    Some(MetadataField::Subject)
                } else if element_is(&xml, &element, DC_NS, b"description") {
                    Some(MetadataField::Description)
                } else if element_is(&xml, &element, CORE_PROPERTIES_NS, b"lastModifiedBy") {
                    Some(MetadataField::LastModifiedBy)
                } else if element_is(&xml, &element, CORE_PROPERTIES_NS, b"keywords")
                    || element_is(&xml, &element, META_NS, b"keyword")
                {
                    Some(MetadataField::Keyword)
                } else if element_is(&xml, &element, DCTERMS_NS, b"created")
                    || element_is(&xml, &element, META_NS, b"creation-date")
                {
                    Some(MetadataField::Created)
                } else if element_is(&xml, &element, DCTERMS_NS, b"modified")
                    || element_is(&xml, &element, DC_NS, b"date")
                {
                    Some(MetadataField::Modified)
                } else if element_is(&xml, &element, META_NS, b"initial-creator") {
                    Some(MetadataField::InitialCreator)
                } else {
                    None
                };
                if let Some(field) = field {
                    let end = crate::xml::local(element.name().as_ref()).to_vec();
                    if let Some(value) = read_element_text(&mut xml, &end, part)? {
                        assign_metadata(&mut metadata, kind, field, value);
                    }
                } else {
                    depth += 1;
                }
            }
            Event::End(_) if depth > 0 => depth -= 1,
            Event::Eof if depth > 0 => {
                return Err(Error::ParseError("Unexpected end of metadata XML"));
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if matches!(kind, MetadataKind::Odp) && metadata.author.is_none() {
        metadata.author = metadata.last_modified_by.clone();
    }
    Ok(metadata)
}

#[derive(Clone, Copy)]
enum MetadataField {
    Title,
    Creator,
    InitialCreator,
    LastModifiedBy,
    Subject,
    Description,
    Keyword,
    Created,
    Modified,
}

fn read_element_text(
    xml: &mut crate::xml::XmlReader<'_>,
    end: &[u8],
    part: &str,
) -> Result<Option<String>> {
    let mut value = String::new();
    loop {
        match event(xml, part)? {
            Event::Text(content) => value.push_str(&text(&content, part)?),
            Event::GeneralRef(content) => value.push_str(&reference(&content, part)?),
            Event::CData(content) => value.push_str(&String::from_utf8_lossy(content.as_ref())),
            Event::End(element) if end_is(element.name().as_ref(), end) => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of metadata XML")),
            _ => {}
        }
    }
    let value = value.trim();
    Ok((!value.is_empty()).then(|| value.to_string()))
}

fn assign_metadata(
    metadata: &mut PresentationMetadata,
    kind: MetadataKind,
    field: MetadataField,
    value: String,
) {
    match field {
        MetadataField::Title => metadata.title = Some(value),
        MetadataField::Creator if matches!(kind, MetadataKind::Pptx) => {
            metadata.author = Some(value)
        }
        MetadataField::Creator => metadata.last_modified_by = Some(value),
        MetadataField::InitialCreator => metadata.author = Some(value),
        MetadataField::LastModifiedBy => metadata.last_modified_by = Some(value),
        MetadataField::Subject => metadata.subject = Some(value),
        MetadataField::Description => metadata.description = Some(value),
        MetadataField::Keyword => metadata.keywords.push(value),
        MetadataField::Created => metadata.created_at = Some(value),
        MetadataField::Modified => metadata.modified_at = Some(value),
    }
}

pub(crate) fn render_presentation_markdown(
    metadata: &PresentationMetadata,
    include_metadata: bool,
    slides: Vec<Slide>,
) -> Result<String> {
    let mut parts = Vec::with_capacity(slides.len() + usize::from(include_metadata));
    if include_metadata && let Some(comment) = render_metadata_comment(metadata) {
        parts.push(comment);
    }
    for slide in slides {
        parts.push(slide.convert_to_md()?);
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
    (!fields.is_empty()).then(|| format!("<!-- Presentation Metadata\n{}\n-->", fields.join("\n")))
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

#[cfg(test)]
#[path = "../tests/unit/metadata.rs"]
mod tests;
