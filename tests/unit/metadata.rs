use super::*;

#[test]
fn parses_pptx_core_properties() {
    let core = br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"><dc:title>Deck</dc:title><dc:creator>Ada</dc:creator><cp:lastModifiedBy>Grace</cp:lastModifiedBy><dc:subject>Subject</dc:subject><dc:description>Description</dc:description><cp:keywords>rust; slides</cp:keywords><dcterms:created>2026-01-01T00:00:00Z</dcterms:created><dcterms:modified>2026-01-02T00:00:00Z</dcterms:modified></cp:coreProperties>"#;
    let metadata = parse_pptx_metadata(Some(core)).unwrap();
    assert_eq!(metadata.title.as_deref(), Some("Deck"));
    assert_eq!(metadata.author.as_deref(), Some("Ada"));
    assert_eq!(metadata.last_modified_by.as_deref(), Some("Grace"));
    assert_eq!(metadata.keywords, vec!["rust; slides"]);
}

#[test]
fn parses_odp_metadata() {
    let meta = br#"<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"><office:meta><dc:title>Deck</dc:title><meta:initial-creator>Ada</meta:initial-creator><dc:creator>Grace</dc:creator><meta:keyword>rust</meta:keyword><meta:keyword>slides</meta:keyword></office:meta></office:document-meta>"#;
    let metadata = parse_odp_metadata(Some(meta)).unwrap();
    assert_eq!(metadata.author.as_deref(), Some("Ada"));
    assert_eq!(metadata.last_modified_by.as_deref(), Some("Grace"));
    assert_eq!(metadata.keywords, vec!["rust", "slides"]);
}

#[test]
fn metadata_comment_is_safe_and_optional() {
    let metadata = PresentationMetadata {
        title: Some("Deck --> injected\nline".to_string()),
        ..PresentationMetadata::default()
    };
    let rendered = render_metadata_comment(&metadata).unwrap();
    assert!(!rendered[4..rendered.len() - 3].contains("--"));
    assert!(rendered.contains("Deck &#45;&#45;> injected line"));
    assert!(render_metadata_comment(&PresentationMetadata::default()).is_none());
}

#[test]
fn absent_metadata_is_empty_and_malformed_metadata_is_an_error() {
    assert_eq!(
        parse_pptx_metadata(None).unwrap(),
        PresentationMetadata::default()
    );
    assert!(parse_pptx_metadata(Some(b"<broken>")).is_err());
    assert!(parse_odp_metadata(Some(b"<broken>")).is_err());
}
