use pptx_to_md::{
    ImageHandlingMode, ParserConfig, PresentationContainer, PresentationFormat, SlideElement,
};
use std::fs::{self, File};
use std::io::Write;
use std::path::{Path, PathBuf};

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
            ("mimetype".to_string(), b"application/vnd.oasis.opendocument.presentation".to_vec()),
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

    let text = slide.elements.iter().find_map(|element| match element {
        SlideElement::Text(text, position) => Some((text, position)),
        _ => None,
    }).unwrap();
    assert!(text.0.runs.iter().any(|run| run.text.contains("Bold text") && run.formatting.bold));
    assert_eq!(text.1.x, 360_000);
    assert_eq!(text.1.y, 720_000);

    let list = slide.elements.iter().find_map(|element| match element {
        SlideElement::List(list, _) => Some(list),
        _ => None,
    }).unwrap();
    assert_eq!(list.items.len(), 2);
    assert!(!list.items[0].is_ordered);
    assert_eq!(list.items[1].level, 1);
    assert!(list.items[1].is_ordered);

    let table = slide.elements.iter().find_map(|element| match element {
        SlideElement::Table(table, _) => Some(table),
        _ => None,
    }).unwrap();
    assert_eq!(table.rows.len(), 2);
    assert_eq!(table.rows[0].cells.len(), 2);
    assert_eq!(table.rows[1].cells.len(), 2);
    assert_eq!(slide.image_data.get("Pictures/image.bin"), Some(&vec![1, 2, 3]));

    fs::remove_file(path).unwrap();
}

#[test]
fn detects_pptx_without_changing_the_existing_pptx_api() {
    let path = temporary_odp_path("pptx-detection");
    create_presentation_archive(
        &path,
        vec![
            ("[Content_Types].xml".to_string(), b"<Types/>".to_vec()),
            ("ppt/presentation.xml".to_string(), b"<p:presentation/>".to_vec()),
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
