use base64::Engine as _;
use pptx_to_md::{
    ImageHandlingMode, ListKind, ParserConfig, PptxContainer, PresentationContainer,
    PresentationFormat, Slide, SlideBlockContent, SlideElement,
};
use std::fs;
use std::path::PathBuf;

fn pptx_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("integration")
        .join("pptx")
        .join("basic.pptx")
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

fn image_fixture_bytes() -> Vec<u8> {
    fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("tests")
            .join("fixtures")
            .join("integration")
            .join("example.jpg"),
    )
    .expect("read image fixture")
}

#[test]
fn pull_parser_matches_parallel_and_streaming_container_paths() {
    let path = pptx_fixture_path();
    if !path.is_file() {
        return;
    }

    let config = ParserConfig::builder().extract_images(true).build();
    let mut parallel = PptxContainer::open(&path, config.clone()).expect("open PPTX fixture");
    let slides = parallel
        .parse_all_multi_threaded()
        .expect("parse PPTX fixture in parallel");
    assert_eq!(slides.len(), parallel.slide_count as usize);

    let mut streamed = PptxContainer::open(&path, config).expect("open PPTX fixture");
    let streamed_count = streamed.iter_slides().fold(0, |count, slide| {
        slide.expect("stream PPTX slide");
        count + 1
    });
    assert_eq!(streamed_count, slides.len());
}

#[test]
fn exposes_and_renders_pptx_metadata_once() {
    let path = pptx_fixture_path();
    if !path.is_file() {
        return;
    }
    let mut container = PresentationContainer::open_as(
        &path,
        ParserConfig::builder().extract_images(false).build(),
        PresentationFormat::Pptx,
    )
    .expect("open PPTX fixture");
    assert_eq!(container.metadata().author.as_deref(), Some("Doe, John"));
    let markdown = container.convert_to_md().expect("convert presentation");
    assert!(markdown.starts_with("<!-- Presentation Metadata\n"));
    assert_eq!(markdown.matches("Presentation Metadata").count(), 1);
}

fn slide_text(slide: &Slide) -> String {
    slide
        .elements
        .iter()
        .filter_map(|element| match element {
            SlideElement::Text(text, _) => Some(
                text.runs
                    .iter()
                    .map(|run| run.text.as_str())
                    .collect::<String>(),
            ),
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

#[test]
fn parses_real_pptx_fixture_and_preserves_slide_order() {
    let Some(slides) = parse_pptx_fixture() else {
        return;
    };

    assert_eq!(slides.len(), 7);
    assert!(slide_text(&slides[0]).contains("PPTX Parser Fixtures"));
    assert!(slide_text(&slides[1]).contains("Lists"));
    assert!(slide_text(&slides[2]).contains("Tables"));
    assert!(slide_text(&slides[3]).contains("Grouped elements"));
    assert!(slide_text(&slides[4]).contains("Sorting and empty content"));
    assert_eq!(speaker_note_text(&slides[5]), "Speaker notes\n");
    assert_eq!(comment_text(&slides[5]), "Comment\n");
    assert!(slide_text(&slides[6]).contains("Image"));
}

#[test]
fn parses_title_and_run_formatting_from_real_pptx() {
    let Some(slides) = parse_pptx_fixture() else {
        return;
    };
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

    assert!(
        runs.iter()
            .any(|run| run.text.contains("Bold") && run.formatting.bold)
    );
    assert!(
        runs.iter()
            .any(|run| run.text.contains("Italic") && run.formatting.italic)
    );
    assert!(
        runs.iter()
            .any(|run| run.text.contains("Underlined") && run.formatting.underlined)
    );
    assert!(
        runs.iter()
            .any(|run| run.text.contains("Bold") && run.formatting.bold && run.formatting.italic)
    );

    let markdown = slides[0].convert_to_md().expect("render first slide");
    assert!(
        markdown.contains(
            "Plain paragraph\n\n**Bold text**\n\n_Italic text_\n\n<u>Underlined text</u>\n\n***Bold and italic text***"
        ),
        "{markdown}"
    );
}

#[test]
fn parses_bulleted_and_numbered_lists_from_real_pptx() {
    let Some(slides) = parse_pptx_fixture() else {
        return;
    };
    let paragraphs: Vec<_> = slides[1]
        .blocks
        .iter()
        .filter_map(|block| match &block.content {
            SlideBlockContent::Text(text) => Some(text.paragraphs.iter()),
            _ => None,
        })
        .flatten()
        .collect();

    assert!(paragraphs.iter().any(|paragraph| {
        paragraph.text().contains("First bullet")
            && matches!(
                paragraph.list.as_ref().map(|list| &list.kind),
                Some(ListKind::Bullet { .. })
            )
    }));
    assert!(paragraphs.iter().any(|paragraph| {
        paragraph.text().contains("Nested bullet")
            && paragraph.list.as_ref().is_some_and(|list| list.level == 1)
    }));
    assert!(paragraphs.iter().any(|paragraph| {
        paragraph.text().contains("First number")
            && matches!(
                paragraph.list.as_ref().map(|list| &list.kind),
                Some(ListKind::Ordered { .. })
            )
    }));
    assert!(paragraphs.iter().any(|paragraph| {
        paragraph.text().contains("Nested number")
            && paragraph.list.as_ref().is_some_and(|list| {
                list.level == 1 && matches!(list.kind, ListKind::Ordered { .. })
            })
    }));
    assert!(paragraphs.iter().any(|paragraph| {
        paragraph.text().contains("Link bullet")
            && paragraph.runs.iter().all(|run| {
                run.link_target.as_deref() == Some("https://github.com/nilskruthoff/pptx-parser")
            })
    }));

    let markdown = slides[1].convert_to_md().expect("render list markdown");
    assert!(markdown.contains("[Link bullet](https://github.com/nilskruthoff/pptx-parser)"));
}

#[test]
fn parses_formatted_table_from_real_pptx() {
    let Some(slides) = parse_pptx_fixture() else {
        return;
    };
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
    let Some(slides) = parse_pptx_fixture() else {
        return;
    };

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

#[test]
fn extracts_and_embeds_the_image_on_slide_seven() {
    let mut container = PresentationContainer::open_as(
        &pptx_fixture_path(),
        ParserConfig::builder()
            .extract_images(true)
            .compress_images(false)
            .image_handling_mode(ImageHandlingMode::InMarkdown)
            .build(),
        PresentationFormat::Pptx,
    )
    .expect("open PPTX fixture");
    let slides = container.parse_all().expect("parse PPTX fixture");
    let slide = slides
        .iter()
        .find(|slide| slide.slide_number == 7)
        .expect("image slide");

    assert!(slide_text(slide).contains("Image"));
    assert_eq!(slide.images.len(), 1);
    let image = slide.images.first().expect("image reference");
    assert!(slide.elements.iter().any(|element| {
        matches!(element, SlideElement::Image(reference, _) if reference.id == image.id)
    }));

    let expected_bytes = image_fixture_bytes();
    assert_eq!(slide.image_data.get(&image.id), Some(&expected_bytes));

    let expected_base64 = base64::engine::general_purpose::STANDARD.encode(expected_bytes);
    let markdown = slide.convert_to_md().expect("render image slide");
    assert!(markdown.contains("data:image/"), "{markdown}");
    assert!(markdown.contains(&expected_base64));
}

#[test]
fn presentation_container_exercises_parallel_and_streaming_pptx_wrappers() {
    let path = pptx_fixture_path();
    if !path.is_file() {
        return;
    }
    let config = ParserConfig::builder().extract_images(false).build();

    let mut parallel =
        PresentationContainer::open_as(&path, config.clone(), PresentationFormat::Pptx)
            .expect("open PPTX fixture");
    let slides = parallel
        .parse_all_multi_threaded()
        .expect("parse PPTX through presentation wrapper");

    let mut converted =
        PresentationContainer::open_as(&path, config.clone(), PresentationFormat::Pptx)
            .expect("open PPTX fixture");
    let markdown = converted
        .convert_to_md_multi_threaded()
        .expect("convert PPTX through presentation wrapper");
    assert!(!markdown.is_empty());

    let mut streamed = PresentationContainer::open_as(&path, config, PresentationFormat::Pptx)
        .expect("open PPTX fixture");
    let mut iterator = streamed.iter_slides();
    let mut streamed_count = 0;
    for slide in iterator.by_ref() {
        slide.expect("stream PPTX slide through presentation wrapper");
        streamed_count += 1;
    }
    assert_eq!(streamed_count, slides.len());
    assert!(iterator.next().is_none());
}
