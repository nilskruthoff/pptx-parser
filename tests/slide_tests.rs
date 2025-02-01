use pptx_parser::{Error, PptxContainer, parse_slide_xml, SlideElement};


#[test]
fn test_invalid_slide_number() {
    let container = PptxContainer::open("test.pptx".as_ref()).unwrap();
    let result = container.read_slide(999);
    assert!(matches!(result, Err(Error::SlideNotFound)));
}

#[test]
fn test_parse_slide_text() -> Result<(), Error> {
    let path = std::path::Path::new("tests/data/test_presentation.pptx");
    let container = PptxContainer::open(&path)?;
    let slide_paths = container.get_slide_paths();

    if let Some(first_slide_path) = slide_paths.first() {
        let slide_data = container.read_slide_by_path(first_slide_path)?;
        let slide = parse_slide_xml(slide_data)?;
        assert!(!slide.elements.is_empty());
        for element in slide.elements {
            if let SlideElement::Text(text_element) = element {
                println!("Found Text: {}", text_element.text);
                // Weitere Assertions können hier hinzugefügt werden
            }
        }
    } else {
        panic!("No Slide found.");
    }
    Ok(())
}