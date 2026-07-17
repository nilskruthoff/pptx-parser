use base64::Engine as _;
use pptx_to_md::{
    ImageHandlingMode, ParserConfig, PresentationContainer, PresentationFormat, SlideElement,
};
use std::fs::{self, File};
use std::io::Write;
use std::path::{Path, PathBuf};

fn temporary_odp_path(name: &str) -> PathBuf {
    std::env::temp_dir().join(format!("pptx-to-md-{name}-{}.odp", std::process::id()))
}

fn odp_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("integration")
        .join("odp")
        .join("basic.odp")
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

fn create_presentation_archive(path: &Path, files: Vec<(String, Vec<u8>)>) {
    let file = File::create(path).expect("create presentation fixture");
    let mut archive = zip::ZipWriter::new(file);
    let options = zip::write::SimpleFileOptions::default();
    for (name, bytes) in files {
        archive
            .start_file(name, options)
            .expect("add presentation entry");
        archive.write_all(&bytes).expect("write presentation entry");
    }
    archive.finish().expect("finish presentation fixture");
}

#[test]
fn parses_odp_tables_lists_and_text_formatting() {
    let path = temporary_odp_path("elements");
    let content = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink">
  <office:automatic-styles>
    <style:style style:name="Bold" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>
    <text:list-style style:name="Bullets"><text:list-level-style-bullet text:level="1" text:bullet-char="-"/></text:list-style>
    <text:list-style style:name="Numbers"><text:list-level-style-number text:level="2"/></text:list-style>
  </office:automatic-styles>
  <office:body><office:presentation><draw:page draw:name="Test">
    <draw:frame svg:x="1cm" svg:y="2cm"><draw:text-box><text:p><text:span text:style-name="Bold">Bold text</text:span></text:p><text:list text:style-name="Bullets"><text:list-item><text:p>First</text:p><text:list text:style-name="Numbers"><text:list-item><text:p>Second</text:p></text:list-item></text:list></text:list-item></text:list></draw:text-box></draw:frame>
    <draw:frame svg:x="2cm" svg:y="3cm"><table:table><table:table-row><table:table-cell table:number-columns-repeated="2"><text:p>Header</text:p></table:table-cell></table:table-row><table:table-row><table:table-cell><text:p>Value</text:p></table:table-cell><table:covered-table-cell/></table:table-row></table:table></draw:frame>
    <draw:frame svg:x="3cm" svg:y="4cm"><draw:image xlink:href="Pictures/image.bin"/></draw:frame>
  </draw:page></office:presentation></office:body>
</office:document-content>"#;
    create_presentation_archive(
        &path,
        vec![
            (
                "mimetype".to_string(),
                b"application/vnd.oasis.opendocument.presentation".to_vec(),
            ),
            ("content.xml".to_string(), content.as_bytes().to_vec()),
            ("Pictures/image.bin".to_string(), vec![1, 2, 3]),
        ],
    );

    let config = ParserConfig::builder()
        .compress_images(false)
        .image_handling_mode(ImageHandlingMode::Manually)
        .build();
    let mut container =
        PresentationContainer::open_as(&path, config, PresentationFormat::Odp).unwrap();
    assert_eq!(container.format(), PresentationFormat::Odp);
    let slide = container.parse_all().unwrap().pop().unwrap();

    let text = slide
        .elements
        .iter()
        .find_map(|element| match element {
            SlideElement::Text(text, position) => Some((text, position)),
            _ => None,
        })
        .unwrap();
    assert!(
        text.0
            .runs
            .iter()
            .any(|run| run.text.contains("Bold text") && run.formatting.bold)
    );
    assert_eq!(text.1.x, 360_000);
    assert_eq!(text.1.y, 720_000);

    let list = slide
        .elements
        .iter()
        .find_map(|element| match element {
            SlideElement::List(list, _) => Some(list),
            _ => None,
        })
        .unwrap();
    assert_eq!(list.items.len(), 2);
    assert!(!list.items[0].is_ordered);
    assert_eq!(list.items[1].level, 1);
    assert!(list.items[1].is_ordered);

    let table = slide
        .elements
        .iter()
        .find_map(|element| match element {
            SlideElement::Table(table, _) => Some(table),
            _ => None,
        })
        .unwrap();
    assert_eq!(table.rows.len(), 2);
    assert_eq!(table.rows[0].cells.len(), 2);
    assert_eq!(table.rows[1].cells.len(), 2);
    assert_eq!(
        slide.image_data.get("Pictures/image.bin"),
        Some(&vec![1, 2, 3])
    );

    fs::remove_file(path).unwrap();
}

#[test]
fn detects_pptx_without_changing_the_existing_pptx_api() {
    let path = temporary_odp_path("pptx-detection");
    create_presentation_archive(
        &path,
        vec![
            ("[Content_Types].xml".to_string(), b"<Types/>".to_vec()),
            (
                "ppt/presentation.xml".to_string(),
                b"<p:presentation/>".to_vec(),
            ),
        ],
    );

    let mut container = PresentationContainer::open(&path, ParserConfig::default()).unwrap();
    assert_eq!(container.format(), PresentationFormat::Pptx);
    assert!(container.parse_all().unwrap().is_empty());

    let mut explicit_container =
        PresentationContainer::open_as(&path, ParserConfig::default(), PresentationFormat::Pptx)
            .unwrap();
    assert_eq!(explicit_container.format(), PresentationFormat::Pptx);
    assert!(explicit_container.parse_all().unwrap().is_empty());

    fs::remove_file(path).unwrap();
}

#[test]
fn extracts_and_embeds_the_image_on_slide_seven() {
    let mut container = PresentationContainer::open_as(
        &odp_fixture_path(),
        ParserConfig::builder()
            .extract_images(true)
            .compress_images(false)
            .image_handling_mode(ImageHandlingMode::InMarkdown)
            .build(),
        PresentationFormat::Odp,
    )
    .expect("open ODP fixture");
    let slides = container.parse_all().expect("parse ODP fixture");
    let slide = slides
        .iter()
        .find(|slide| slide.slide_number == 7)
        .expect("image slide");

    assert!(slide.elements.iter().any(|element| {
        matches!(element, SlideElement::Text(text, _) if text.runs.iter().any(|run| run.text.contains("Image")))
    }));
    assert_eq!(slide.images.len(), 1);
    let image = slide.images.first().expect("image reference");
    assert!(slide.elements.iter().any(|element| {
        matches!(element, SlideElement::Image(reference, _) if reference.id == image.id)
    }));

    let expected_bytes = image_fixture_bytes();
    assert_eq!(slide.image_data.get(&image.id), Some(&expected_bytes));

    let expected_base64 = base64::engine::general_purpose::STANDARD.encode(expected_bytes);
    let markdown = slide.convert_to_md().expect("render image slide");
    assert!(markdown.contains("data:image/"));
    assert!(markdown.contains(&expected_base64));
}

#[test]
fn presentation_container_exercises_odp_detection_and_wrapper_paths() {
    let path = odp_fixture_path();
    let config = ParserConfig::builder().extract_images(false).build();

    let detected = PresentationContainer::open(&path, config.clone()).expect("detect ODP fixture");
    assert_eq!(detected.format(), PresentationFormat::Odp);
    let _metadata = detected.metadata();

    let mut parallel = PresentationContainer::open(&path, config.clone()).expect("open ODP");
    let slides = parallel
        .parse_all_multi_threaded()
        .expect("parse ODP through parallel-compatible API");
    assert!(!slides.is_empty());

    let mut converted = PresentationContainer::open(&path, config.clone()).expect("open ODP");
    let markdown = converted.convert_to_md().expect("convert ODP");
    assert!(!markdown.is_empty());

    let mut converted_parallel =
        PresentationContainer::open(&path, config.clone()).expect("open ODP");
    assert_eq!(
        converted_parallel
            .convert_to_md_multi_threaded()
            .expect("convert ODP through parallel-compatible API"),
        markdown
    );

    let mut streamed = PresentationContainer::open(&path, config).expect("open ODP");
    let mut iterator = streamed.iter_slides();
    let mut streamed_count = 0;
    for slide in iterator.by_ref() {
        slide.expect("stream ODP slide");
        streamed_count += 1;
    }
    assert_eq!(streamed_count, slides.len());
    assert!(iterator.next().is_none());
}

#[test]
fn rejects_an_archive_with_an_unknown_presentation_format() {
    let path = temporary_odp_path("unsupported-format");
    create_presentation_archive(
        &path,
        vec![("unrelated.txt".to_string(), b"not a presentation".to_vec())],
    );

    let result = PresentationContainer::open(&path, ParserConfig::default());

    assert!(matches!(
        result,
        Err(pptx_to_md::Error::ParseError(
            "Unsupported presentation format"
        ))
    ));
    fs::remove_file(path).expect("remove unsupported presentation fixture");
}

fn normalize_semantic_markdown(markdown: &str) -> String {
    let mut normalized = String::new();
    let mut previous_blank = false;
    for line in markdown.lines().map(str::trim_end) {
        let blank = line.is_empty();
        if blank && previous_blank {
            continue;
        }
        normalized.push_str(line);
        normalized.push('\n');
        previous_blank = blank;
    }
    normalized.trim().to_string()
}

#[test]
fn equivalent_pptx_and_odp_slides_produce_equivalent_semantic_markdown() {
    let pptx_path = temporary_odp_path("semantic-parity-pptx");
    let pptx_slide = br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>Shared title</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:pPr><a:buChar char="-"/></a:pPr><a:r><a:t>Shared item</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#;
    create_presentation_archive(
        &pptx_path,
        vec![
            ("[Content_Types].xml".to_string(), b"<Types/>".to_vec()),
            (
                "ppt/presentation.xml".to_string(),
                b"<p:presentation/>".to_vec(),
            ),
            ("ppt/slides/slide1.xml".to_string(), pptx_slide.to_vec()),
        ],
    );

    let odp_path = temporary_odp_path("semantic-parity-odp");
    let odp_content = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"><office:body><office:presentation><draw:page><draw:frame presentation:class="title"><draw:text-box><text:p>Shared title</text:p></draw:text-box></draw:frame><draw:frame presentation:class="body"><draw:text-box><text:list><text:list-item><text:p>Shared item</text:p></text:list-item></text:list></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#;
    create_presentation_archive(
        &odp_path,
        vec![
            (
                "mimetype".to_string(),
                b"application/vnd.oasis.opendocument.presentation".to_vec(),
            ),
            ("content.xml".to_string(), odp_content.to_vec()),
        ],
    );

    let config = ParserConfig::builder().extract_images(false).build();
    let mut pptx =
        PresentationContainer::open_as(&pptx_path, config.clone(), PresentationFormat::Pptx)
            .unwrap();
    let mut odp =
        PresentationContainer::open_as(&odp_path, config, PresentationFormat::Odp).unwrap();

    let pptx_markdown = normalize_semantic_markdown(&pptx.convert_to_md().unwrap());
    let odp_markdown = normalize_semantic_markdown(&odp.convert_to_md().unwrap());
    assert_eq!(pptx_markdown, odp_markdown);

    fs::remove_file(pptx_path).unwrap();
    fs::remove_file(odp_path).unwrap();
}

#[test]
fn semantic_document_reports_missing_resources_without_losing_the_slide() {
    let path = temporary_odp_path("missing-image-diagnostic");
    let content = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:presentation><draw:page><draw:frame draw:name="Missing diagram"><draw:image xlink:href="Pictures/missing.png"/></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#;
    create_presentation_archive(
        &path,
        vec![
            (
                "mimetype".to_string(),
                b"application/vnd.oasis.opendocument.presentation".to_vec(),
            ),
            ("content.xml".to_string(), content.to_vec()),
        ],
    );

    let mut container = PresentationContainer::open(&path, ParserConfig::default()).unwrap();
    let document = container.parse_document().unwrap();

    assert_eq!(document.slides.len(), 1);
    assert_eq!(document.diagnostics.len(), 1);
    assert!(
        document.diagnostics[0]
            .message
            .contains("Image resource could not be loaded")
    );
    assert!(
        document.slides[0]
            .convert_to_md()
            .unwrap()
            .contains("[Image unavailable: Missing diagram]")
    );

    fs::remove_file(path).unwrap();
}
