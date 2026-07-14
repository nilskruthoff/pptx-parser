use super::*;
use std::fs;
use std::io::Write;
use zip::write::SimpleFileOptions;

#[test]
fn sorts_slide_paths_numerically() {
    let mut slide_paths = vec!["ppt/slides/slide10.xml".to_string(), "ppt/slides/slide2.xml".to_string(), "ppt/slides/slide1.xml".to_string()];
    sort_slide_paths(&mut slide_paths);
    assert_eq!(slide_paths, vec!["ppt/slides/slide1.xml".to_string(), "ppt/slides/slide2.xml".to_string(), "ppt/slides/slide10.xml".to_string()]);
}

#[test]
fn resolves_target_path_with_parent_segments() {
    assert_eq!(PptxContainer::resolve_target_path("ppt/slides/slide3.xml", "../slideLayouts/slideLayout3.xml"), "ppt/slideLayouts/slideLayout3.xml");
}

#[test]
fn resolves_root_relative_target_path() {
    assert_eq!(PptxContainer::resolve_target_path("ppt/slides/slide3.xml", "/ppt/slideLayouts/slideLayout3.xml"), "ppt/slideLayouts/slideLayout3.xml");
}

#[test]
fn loads_speaker_notes_from_a_slide_relationship() {
    let path = std::env::temp_dir().join(format!("pptx-to-md-speaker-notes-{}.pptx", std::process::id()));
    let file = fs::File::create(&path).expect("create temporary PPTX");
    let mut archive = zip::ZipWriter::new(file);
    let options = SimpleFileOptions::default();
    archive.start_file("ppt/slides/slide1.xml", options).expect("start slide entry");
    archive.write_all(br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>"#).expect("write slide entry");
    archive.start_file("ppt/slides/_rels/slide1.xml.rels", options).expect("start relationship entry");
    archive.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/></Relationships>"#).expect("write relationship entry");
    archive.start_file("ppt/notesSlides/notesSlide1.xml", options).expect("start notes entry");
    archive.write_all(br#"<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>Presenter detail</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>"#).expect("write notes entry");
    archive.finish().expect("finish temporary PPTX");
    let mut container = PptxContainer::open(&path, ParserConfig::default()).expect("open temporary PPTX");
    let slides = container.parse_all().expect("parse temporary PPTX");
    assert_eq!(slides.len(), 1);
    assert_eq!(slides[0].speaker_notes[0].runs[0].text, "Presenter detail\n");
    drop(container);
    fs::remove_file(path).expect("remove temporary PPTX");
}
