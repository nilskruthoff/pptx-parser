use pptx_parser::{Error, PptxContainer, parse_slide_xml, SlideElement};

#[test]
fn test_parse_slide_text() -> Result<(), Error> {
    let path = std::path::Path::new("test-data/sample.pptx");
    let container = PptxContainer::open(&path)?;
    let slide_paths = container.get_slide_paths();

    if let Some(first_slide_path) = slide_paths.first() {
        let slide_data = container.read_slide_by_path(first_slide_path)?;
        let slide = parse_slide_xml(slide_data)?;
        assert!(!slide.elements.is_empty());
        for element in slide.elements {
            if let SlideElement::Text(text_element) = element {
                println!("Found Text:");
            }
        }
    } else {
        panic!("No Slide found.");
    }
    Ok(())
}

#[test]
fn test_parse_all_slides() -> Result<(), Error> {
    let path = std::path::Path::new("test-data/sample.pptx");
    let container = PptxContainer::open(&path)?;
    let slide_paths = container.get_slide_paths();

    if !(slide_paths.len() > 0) { panic!("No Slide found.") }

    for slide_path in slide_paths {
        let slide_data = container.read_slide_by_path(&slide_path)?;
        let slide = parse_slide_xml(slide_data)?;
        for element in slide.elements {
            println!("{:?}", element);
        }
    }

    Ok(())
}

#[test]
fn test_parse_text() -> Result<(), Error> {
    let path = std::path::Path::new("test-data/sample.pptx");
    let container = PptxContainer::open(&path)?;
    let txt = container.extract_text()?;
    println!("{}", txt);
    Ok(())
}