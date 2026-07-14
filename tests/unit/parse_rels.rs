use super::*;
use std::fs;
use std::path::PathBuf;

fn load_xml(filename: &str) -> Vec<u8> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("unit")
        .join("ooxml")
        .join(filename);
    fs::read(path).expect("read relationship fixture")
}

fn normalize_test_string(input: &str) -> String {
    input.trim_start_matches('\u{feff}').replace("\r\n", "\n").replace("    ", "\t").trim().to_string()
}

#[test]
fn parses_slide_relationships_with_images() {
    let images = parse_slide_rels(&load_xml("rels_with_images.xml")).expect("parse image relationships");
    assert_eq!(images.len(), 2);
    assert_eq!(images[0].id, "rId1");
    assert_eq!(normalize_test_string(&images[0].target), normalize_test_string("../media/image1.png"));
    assert_eq!(images[1].id, "rId2");
    assert_eq!(normalize_test_string(&images[1].target), normalize_test_string("../media/image2.jpg"));
}

#[test]
fn parses_empty_slide_relationships() {
    assert!(parse_slide_rels(&load_xml("rels_without_images.xml")).unwrap().is_empty());
}

#[test]
fn parses_hyperlink_relationships() {
    let hyperlinks = parse_hyperlink_rels(&load_xml("rels_with_images.xml")).expect("parse hyperlink relationships");
    assert_eq!(hyperlinks.get("rId3"), Some(&"https://example.com".to_string()));
    assert_eq!(hyperlinks.len(), 1);
}
