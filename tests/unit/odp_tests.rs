use super::*;
use crate::{ImageHandlingMode, ParserConfig, PresentationContainer, PresentationFormat, SlideElement};
use roxmltree::Document;
use std::fs::{self, File};
use std::io::Write;
use std::path::{Path, PathBuf};

fn load_odp_xml(filename: &str) -> Vec<u8> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("odp_test_data")
        .join(filename);
    fs::read(path).expect("read ODP XML fixture")
}

fn styles() -> StyleResolver {
    StyleResolver::from_documents(b"", &load_odp_xml("styles.xml")).expect("parse ODP styles")
}

fn temporary_odp_path(name: &str) -> PathBuf {
    std::env::temp_dir().join(format!("pptx-to-md-{name}-{}.odp", std::process::id()))
}

fn create_presentation_archive(path: &Path, files: Vec<(String, Vec<u8>)>) {
    let file = File::create(path).expect("create presentation fixture");
    let mut archive = zip::ZipWriter::new(file);
    let options = zip::write::SimpleFileOptions::default();
    for (name, bytes) in files {
        archive.start_file(name, options).expect("add presentation entry");
        archive.write_all(&bytes).expect("write presentation entry");
    }
    archive.finish().expect("finish presentation fixture");
}

#[test]
fn parses_heading_from_xml_fixture() {
    let xml = load_odp_xml("heading.xml");
    let document = Document::parse(std::str::from_utf8(&xml).unwrap()).unwrap();
    let mut elements = Vec::new();

    parse_text_container(document.root_element(), ElementPosition::default(), &styles(), &mut elements).unwrap();

    let SlideElement::Text(text, _) = &elements[0] else { panic!("expected heading text"); };
    assert_eq!(text.runs[0].text, "Heading\n");
    assert!(text.runs[0].formatting.bold);
}

#[test]
fn parses_nested_list_from_xml_fixture() {
    let xml = load_odp_xml("list.xml");
    let document = Document::parse(std::str::from_utf8(&xml).unwrap()).unwrap();

    let list = parse_list(document.root_element(), &styles(), 0);

    assert_eq!(list.items.len(), 2);
    assert_eq!(list.items[0].runs[0].text, "First\n");
    assert!(!list.items[0].is_ordered);
    assert_eq!(list.items[1].runs[0].text, "Second\n");
    assert_eq!(list.items[1].level, 1);
    assert!(list.items[1].is_ordered);
}

#[test]
fn parses_formatted_table_row_from_xml_fixture() {
    let xml = load_odp_xml("table.xml");
    let document = Document::parse(std::str::from_utf8(&xml).unwrap()).unwrap();
    let table_node = document.root_element();
    let header_row = table_node.children().find(|node| is_element(*node, TABLE_NS, "table-row")).unwrap();

    let row = parse_table_row(header_row, &styles());

    assert_eq!(row.cells.len(), 3);
    for cell in row.cells {
        assert_eq!(cell.runs[0].text, "Heading\n");
        assert!(cell.runs[0].formatting.bold);
    }
}

#[test]
fn parses_table_shape_from_xml_fixture() {
    let xml = load_odp_xml("table.xml");
    let document = Document::parse(std::str::from_utf8(&xml).unwrap()).unwrap();

    let table = parse_table(document.root_element(), &styles()).unwrap();

    assert_eq!(table.rows.len(), 2);
    assert_eq!(table.rows[0].cells.len(), 3);
    assert_eq!(table.rows[1].cells.len(), 3);
    assert_eq!(table.rows[1].cells[0].runs[0].text, "Value\n");
    assert!(table.rows[1].cells[1].runs.is_empty());
}

#[test]
fn parses_image_reference_from_xml_fixture() {
    let xml = load_odp_xml("image.xml");
    let document = Document::parse(std::str::from_utf8(&xml).unwrap()).unwrap();

    let image = parse_image(document.root_element()).expect("image reference");

    assert_eq!(image.id, "Pictures/image.bin");
    assert_eq!(image.target, "Pictures/image.bin");
}

#[test]
fn parses_complete_odp_content_fixture() {
    let path = temporary_odp_path("fixture-content");
    create_presentation_archive(&path, vec![
        ("mimetype".to_string(), b"application/vnd.oasis.opendocument.presentation".to_vec()),
        ("content.xml".to_string(), load_odp_xml("content.xml")),
        ("styles.xml".to_string(), load_odp_xml("styles.xml")),
        ("Pictures/image.bin".to_string(), vec![1, 2, 3]),
    ]);

    let config = ParserConfig::builder()
        .compress_images(false)
        .image_handling_mode(ImageHandlingMode::Manually)
        .build();
    let mut container = PresentationContainer::open_as(&path, config, PresentationFormat::Odp).unwrap();
    let slide = container.parse_all().unwrap().pop().unwrap();

    assert!(slide.elements.iter().any(|element| matches!(element, SlideElement::Text(_, _))));
    assert!(slide.elements.iter().any(|element| matches!(element, SlideElement::List(_, _))));
    assert!(slide.elements.iter().any(|element| matches!(element, SlideElement::Table(_, _))));
    assert_eq!(slide.image_data.get("Pictures/image.bin"), Some(&vec![1, 2, 3]));

    fs::remove_file(path).unwrap();
}

#[test]
fn detects_pptx_without_changing_the_existing_pptx_api() {
    let path = temporary_odp_path("pptx-detection");
    create_presentation_archive(&path, vec![
        ("[Content_Types].xml".to_string(), b"<Types/>".to_vec()),
        ("ppt/presentation.xml".to_string(), b"<p:presentation/>".to_vec()),
    ]);

    let mut container = PresentationContainer::open(&path, ParserConfig::default()).unwrap();
    assert_eq!(container.format(), PresentationFormat::Pptx);
    assert!(container.parse_all().unwrap().is_empty());

    fs::remove_file(path).unwrap();
}
