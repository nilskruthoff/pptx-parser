use super::*;
use crate::{ParserConfig, PresentationFormat, Slide, SlideElement};
use std::path::PathBuf;

fn pptx_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("test_data")
        .join("test.pptx")
}

fn parse_pptx_fixture() -> Option<Vec<Slide>> {
    let path = pptx_fixture_path();
    if !path.is_file() {
        return None;
    }

    let mut container = PresentationContainer::open_as(
        &path,
        ParserConfig::builder().extract_images(false).build(),
        PresentationFormat::Pptx,
    )
    .expect("open PPTX fixture");

    assert_eq!(container.format(), PresentationFormat::Pptx);
    Some(container.parse_all().expect("parse PPTX fixture"))
}

fn slide_text(slide: &Slide) -> String {
    slide
        .elements
        .iter()
        .filter_map(|element| match element {
            SlideElement::Text(text, _) => Some(text.runs.iter().map(|run| run.text.as_str()).collect::<String>()),
            _ => None,
        })
        .collect()
}

fn list_item_text(item: &crate::ListItem) -> String {
    item.runs.iter().map(|run| run.text.as_str()).collect()
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

#[test]
fn parses_real_pptx_fixture_and_preserves_slide_order() {
    let Some(slides) = parse_pptx_fixture() else { return; };

    assert_eq!(slides.len(), 6);
    assert!(slide_text(&slides[0]).contains("PPTX Parser Fixtures"));
    assert!(slide_text(&slides[1]).contains("Lists"));
    assert!(slide_text(&slides[2]).contains("Tables"));
    assert!(slide_text(&slides[3]).contains("Grouped elements"));
    assert!(slide_text(&slides[4]).contains("Sorting and empty content"));
    assert_eq!(speaker_note_text(&slides[5]), "Speaker notes\n");
    assert_eq!(comment_text(&slides[5]), "Comment\n");
}

#[test]
fn parses_title_and_run_formatting_from_real_pptx() {
    let Some(slides) = parse_pptx_fixture() else { return; };
    assert!(slide_text(&slides[0]).contains("PPTX Parser Fixtures"));
    let runs: Vec<_> = slides[0]
        .elements
        .iter()
        .filter_map(|element| match element {
            SlideElement::Text(text, _) => Some(text.runs.iter()),
            _ => None,
        })
        .flatten()
        .collect();

    assert!(runs.iter().any(|run| run.text.contains("Bold") && run.formatting.bold));
    assert!(runs.iter().any(|run| run.text.contains("Italic") && run.formatting.italic));
    assert!(runs.iter().any(|run| run.text.contains("Underlined") && run.formatting.underlined));
    assert!(runs.iter().any(|run| run.text.contains("Bold") && run.formatting.bold && run.formatting.italic));
}

#[test]
fn parses_bulleted_and_numbered_lists_from_real_pptx() {
    let Some(slides) = parse_pptx_fixture() else { return; };
    let list = slides[1]
        .elements
        .iter()
        .find_map(|element| match element {
            SlideElement::List(list, _) => Some(list),
            _ => None,
        })
        .expect("list on second slide");

    assert!(list.items.iter().any(|item| list_item_text(item).contains("First bullet") && !item.is_ordered));
    assert!(list.items.iter().any(|item| list_item_text(item).contains("Nested bullet") && item.level == 1));
    assert!(list.items.iter().any(|item| list_item_text(item).contains("First number") && item.is_ordered));
    assert!(list.items.iter().any(|item| list_item_text(item).contains("Nested number") && item.level == 1 && item.is_ordered));
    assert!(list.items.iter().any(|item| {
        list_item_text(item).contains("Link bullet")
            && item
                .runs
                .iter()
                .all(|run| run.link_target.as_deref() == Some("https://github.com/nilskruthoff/pptx-parser"))
    }));

    let markdown = slides[1].convert_to_md().expect("render list markdown");
    assert!(markdown.contains("[Link ](https://github.com/nilskruthoff/pptx-parser)[bullet](https://github.com/nilskruthoff/pptx-parser)"));
}

#[test]
fn parses_formatted_table_from_real_pptx() {
    let Some(slides) = parse_pptx_fixture() else { return; };
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
    assert_eq!(table.rows[0].cells[0].runs[0].text, "Heading A");
    assert_eq!(table.rows[1].cells[1].runs[0].text, "B1");
    assert_eq!(table.rows[2].cells[2].runs[0].text, "C2");
    assert_eq!(
        table.rows[2].cells[2].runs[0].link_target.as_deref(),
        Some("https://github.com/nilskruthoff/pptx-parser")
    );

    let markdown = slides[2].convert_to_md().expect("render table markdown");
    assert!(markdown.contains("[C2](https://github.com/nilskruthoff/pptx-parser)"));
}

#[test]
fn parses_grouped_text_sorting_and_empty_cells_from_real_pptx() {
    let Some(slides) = parse_pptx_fixture() else { return; };

    let grouped = slide_text(&slides[3]);
    assert!(grouped.contains("Grouped Heading"));
    assert!(grouped.contains("Grouped Body"));

    let markdown = slides[4].convert_to_md().expect("render slide markdown");
    let first = markdown.find("First").expect("First text");
    let second = markdown.find("Second").expect("Second text");
    let third = markdown.find("Third").expect("Third text");
    assert!(first < second && second < third);

    let table = slides[4]
        .elements
        .iter()
        .find_map(|element| match element {
            SlideElement::Table(table, _) => Some(table),
            _ => None,
        })
        .expect("table on fifth slide");
    assert!(table.rows[0].cells[1].runs.is_empty());
    assert!(table.rows[1].cells[0].runs.is_empty());
    assert!(table.rows[1].cells[2].runs.is_empty());
}
