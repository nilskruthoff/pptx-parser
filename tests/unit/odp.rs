use super::*;
use crate::{ParserConfig, PresentationContainer, PresentationFormat, SlideElement};
use roxmltree::Document;
use std::fs;
use std::path::PathBuf;

#[test]
fn parses_text_from_odp_speaker_notes_only() {
    let xml = r#"
        <draw:page xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
                   xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
                   xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
            <draw:custom-shape><text:p>Slide content</text:p></draw:custom-shape>
            <presentation:notes><draw:custom-shape><text:p>Speaker note</text:p></draw:custom-shape></presentation:notes>
        </draw:page>
    "#;
    let document = Document::parse(xml).expect("parse ODP XML");
    let notes = parse_speaker_notes(document.root_element(), &StyleResolver::default())
        .expect("parse speaker notes");
    assert_eq!(notes.len(), 1);
    assert_eq!(notes[0].runs[0].text, "Speaker note\n");
}

fn odp_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("integration")
        .join("odp")
        .join("basic.odp")
}

fn text_from_slide(slide: &Slide) -> Vec<String> {
    slide
        .elements
        .iter()
        .filter_map(|element| match element {
            SlideElement::Text(text, _) => {
                Some(text.runs.iter().map(|run| run.text.as_str()).collect())
            }
            _ => None,
        })
        .collect()
}

fn speaker_note_text(slide: &Slide) -> String {
    slide
        .speaker_notes
        .iter()
        .flat_map(|note| note.runs.iter())
        .map(|run| run.text.as_str())
        .collect()
}

fn comment_text(slide: &Slide) -> String {
    slide
        .comments
        .iter()
        .flat_map(|comment| comment.runs.iter())
        .map(|run| run.text.as_str())
        .collect()
}

fn fixture_config() -> ParserConfig {
    ParserConfig::builder().extract_images(false).build()
}

fn load_odp_xml(filename: &str) -> Option<Vec<u8>> {
    fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("tests")
            .join("fixtures")
            .join("unit")
            .join("odp")
            .join(filename),
    )
    .ok()
}

macro_rules! load_odp_xml_or_skip {
    ($filename:expr) => {
        match load_odp_xml($filename) {
            Some(xml) => xml,
            None => return,
        }
    };
}

fn styles() -> Option<StyleResolver> {
    let styles = load_odp_xml("styles.xml")?;
    Some(StyleResolver::from_documents(b"", &styles).expect("parse ODP styles"))
}

fn open_real_odp_fixture() -> Option<PresentationContainer> {
    let path = odp_fixture_path();
    if !path.is_file() {
        return None;
    }

    Some(
        PresentationContainer::open_as(&path, fixture_config(), PresentationFormat::Odp)
            .expect("open ODP fixture"),
    )
}

#[test]
fn exposes_and_renders_odp_metadata_once() {
    let Some(mut container) = open_real_odp_fixture() else {
        return;
    };
    assert_eq!(
        container.metadata().created_at.as_deref(),
        Some("2026-07-12T21:10:22.756411800")
    );
    let markdown = container.convert_to_md().expect("convert ODP presentation");
    assert!(markdown.starts_with("<!-- Presentation Metadata\n"));
    assert_eq!(markdown.matches("Presentation Metadata").count(), 1);
}

#[test]
fn parses_heading_from_xml_fixture() {
    let xml = load_odp_xml_or_skip!("heading.xml");
    let document = Document::parse(std::str::from_utf8(&xml).unwrap()).unwrap();
    let mut elements = Vec::new();
    let Some(styles) = styles() else {
        return;
    };

    parse_text_container(
        document.root_element(),
        ElementPosition::default(),
        &styles,
        &mut elements,
    )
    .unwrap();

    let SlideElement::Text(text, _) = &elements[0] else {
        panic!("expected heading text");
    };
    assert_eq!(text.runs[0].text, "Heading\n");
    assert!(text.runs[0].formatting.bold);
}

#[test]
fn parses_hyperlink_from_odp_text_anchor() {
    let xml = r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink">Before <text:a xlink:href="https://example.com">linked <text:span>text</text:span></text:a></text:p>"#;
    let document = Document::parse(xml).expect("parse ODP paragraph");
    let Some(styles) = styles() else {
        return;
    };

    let runs = parse_paragraph(document.root_element(), &styles);
    let linked: String = runs
        .iter()
        .filter(|run| run.link_target.as_deref() == Some("https://example.com"))
        .map(|run| run.text.as_str())
        .collect();
    assert_eq!(linked, "linked text");
}

#[test]
fn parses_nested_list_from_xml_fixture() {
    let xml = load_odp_xml_or_skip!("list.xml");
    let document = Document::parse(std::str::from_utf8(&xml).unwrap()).unwrap();
    let Some(styles) = styles() else {
        return;
    };

    let list = parse_list(document.root_element(), &styles, 0);

    assert_eq!(list.items.len(), 2);
    assert!(!list.items[0].is_ordered);
    assert_eq!(list.items[1].level, 1);
    assert!(list.items[1].is_ordered);
}

#[test]
fn parses_formatted_table_from_xml_fixture() {
    let xml = load_odp_xml_or_skip!("table.xml");
    let document = Document::parse(std::str::from_utf8(&xml).unwrap()).unwrap();
    let table_node = document.root_element();
    let header_row = table_node
        .children()
        .find(|node| is_element(*node, TABLE_NS, "table-row"))
        .unwrap();
    let Some(styles) = styles() else {
        return;
    };

    let row = parse_table_row(header_row, &styles);
    assert_eq!(row.cells.len(), 3);
    assert!(row.cells.iter().all(|cell| cell.runs[0].formatting.bold));

    let table = parse_table(table_node, &styles).unwrap();
    assert_eq!(table.rows.len(), 2);
    assert_eq!(table.rows[1].cells.len(), 3);
    assert!(table.rows[1].cells[1].runs.is_empty());
}

#[test]
fn parses_complete_content_xml_fixture() {
    let xml = load_odp_xml_or_skip!("content.xml");
    let document = Document::parse(std::str::from_utf8(&xml).unwrap()).unwrap();
    let page = presentation_pages(&document)
        .next()
        .expect("presentation page");
    let Some(styles) = styles() else {
        return;
    };

    let elements = parse_page(page, &styles).unwrap();

    assert!(elements.iter().any(|element| matches!(
        element,
        SlideElement::Text(
            _,
            ElementPosition {
                x: 360_000,
                y: 720_000
            }
        )
    )));
    assert!(elements
        .iter()
        .any(|element| matches!(element, SlideElement::List(_, _))));
    assert!(elements
        .iter()
        .any(|element| matches!(element, SlideElement::Table(_, _))));
}

#[test]
fn parses_real_odp_fixture_and_preserves_slide_order() {
    let Some(mut container) = open_real_odp_fixture() else {
        return;
    };

    assert_eq!(container.format(), PresentationFormat::Odp);
    let slides = container.parse_all().expect("parse ODP fixture");

    assert_eq!(slides.len(), 7);
    assert!(text_from_slide(&slides[0])
        .join("\n")
        .contains("ODP Parser Fixture"));
    assert!(text_from_slide(&slides[1]).join("\n").contains("Lists"));
    assert!(text_from_slide(&slides[2]).join("\n").contains("Tables"));
    assert!(text_from_slide(&slides[3])
        .join("\n")
        .contains("Grouped elements"));
    assert!(text_from_slide(&slides[4])
        .join("\n")
        .contains("Sorting and empty content"));
    assert_eq!(speaker_note_text(&slides[5]), "Speaker notes\n");
    assert_eq!(comment_text(&slides[5]), "Comment\n");
    assert!(text_from_slide(&slides[6]).join("\n").contains("Image"));
}

#[test]
fn parses_title_and_run_formatting_from_real_odp() {
    let Some(mut container) = open_real_odp_fixture() else {
        return;
    };
    let slides = container.parse_all().expect("parse ODP fixture");

    let runs: Vec<_> = slides[0]
        .elements
        .iter()
        .filter_map(|element| match element {
            SlideElement::Text(text, _) => Some(text.runs.iter()),
            _ => None,
        })
        .flatten()
        .collect();

    assert!(runs
        .iter()
        .any(|run| run.text.contains("ODP Parser Fixture")));
    assert!(runs
        .iter()
        .any(|run| run.text.contains("Bold text") && run.formatting.bold));
    assert!(runs
        .iter()
        .any(|run| run.text.contains("Italic text") && run.formatting.italic));
    assert!(runs
        .iter()
        .any(|run| run.text.contains("Underlined text") && run.formatting.underlined));
    assert!(runs
        .iter()
        .any(|run| run.text.contains("Bold and italic text")
            && run.formatting.bold
            && run.formatting.italic));

    let markdown = slides[0].convert_to_md().expect("render first ODP slide");
    assert!(markdown.contains(
        "Plain paragraph\n\n**Bold text**\n\n_Italic text_\n\n<u>Underlined text</u>\n\n***Bold and italic text***"
    ));
}

#[test]
fn parses_bulleted_and_numbered_lists_from_real_odp() {
    let Some(mut container) = open_real_odp_fixture() else {
        return;
    };
    let slides = container.parse_all().expect("parse ODP fixture");

    let lists: Vec<_> = slides[1]
        .elements
        .iter()
        .filter_map(|element| match element {
            SlideElement::List(list, _) => Some(list),
            _ => None,
        })
        .collect();

    let items: Vec<_> = lists.iter().flat_map(|list| list.items.iter()).collect();
    assert!(items.iter().any(|item| item
        .runs
        .iter()
        .any(|run| run.text.contains("First bullet"))
        && !item.is_ordered));
    assert!(items.iter().any(|item| item
        .runs
        .iter()
        .any(|run| run.text.contains("Nested bullet"))
        && item.level == 1));
    assert!(items.iter().any(|item| item
        .runs
        .iter()
        .any(|run| run.text.contains("First number"))
        && item.is_ordered));
    assert!(items.iter().any(|item| item
        .runs
        .iter()
        .any(|run| run.text.contains("Nested number"))
        && item.level == 1));
    assert!(items.iter().any(
        |item| item.runs.iter().any(|run| run.text.contains("Link bullet")
            && run.link_target.as_deref() == Some("https://github.com/nilskruthoff/pptx-parser"))
    ));

    let markdown = slides[1].convert_to_md().expect("render list markdown");
    assert!(markdown.contains("[Link bullet](https://github.com/nilskruthoff/pptx-parser)"));
}

#[test]
fn parses_formatted_table_and_empty_cells_from_real_odp() {
    let Some(mut container) = open_real_odp_fixture() else {
        return;
    };
    let slides = container.parse_all().expect("parse ODP fixture");

    let table = slides[2]
        .elements
        .iter()
        .find_map(|element| match element {
            SlideElement::Table(table, _) => Some(table),
            _ => None,
        })
        .expect("table on third slide");

    assert_eq!(table.rows.len(), 3);
    assert_eq!(table.rows[0].cells.len(), 3);
    for cell in &table.rows[0].cells {
        assert!(cell.runs[0].formatting.bold);
        assert!(cell.runs[0].text.contains("Heading"));
    }
    assert_eq!(table.rows[1].cells[0].runs[0].text, "A1\n");
    assert_eq!(table.rows[2].cells[2].runs[0].text, "C2\n");
    assert_eq!(
        table.rows[2].cells[2].runs[0].link_target.as_deref(),
        Some("https://github.com/nilskruthoff/pptx-parser")
    );

    let markdown = slides[2].convert_to_md().expect("render table markdown");
    assert!(markdown.contains("[C2](https://github.com/nilskruthoff/pptx-parser)"));
    assert!(markdown.contains("| A2 | B2 | [C2](https://github.com/nilskruthoff/pptx-parser) |"));
}

#[test]
fn parses_grouped_text_and_keeps_vertical_text_order() {
    let Some(mut container) = open_real_odp_fixture() else {
        return;
    };
    let slides = container.parse_all().expect("parse ODP fixture");

    let grouped_text = text_from_slide(&slides[3]).join("\n");
    assert!(grouped_text.contains("Grouped Heading"));
    assert!(grouped_text.contains("Grouped Body"));

    let markdown = slides[4].convert_to_md().expect("render slide markdown");
    let first = markdown.find("First").expect("First text");
    let second = markdown.find("Second").expect("Second text");
    let third = markdown.find("Third").expect("Third text");
    assert!(first < second && second < third);
}
