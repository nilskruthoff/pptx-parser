use std::collections::HashMap;
use std::fs;
use std::fs::File;
use std::io::Write;
use std::path::{Path, PathBuf};
use pptx_to_md::{Error, Formatting, ListElement, ListItem, PptxContainer, Run, Slide, SlideElement, TableCell, TableElement, TableRow, TextElement};

fn load_test_data(filename: &str) -> String {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("tests");
    path.push("test_data");
    path.push(filename);
    fs::read_to_string(path).expect("Unable to read test data file")
}

fn normalize_test_string(input: &str) -> String {
    input
        .trim_start_matches('\u{feff}') // remove BOM
        .replace("\r\n", "\n") // normalize line breaks
        .replace("    ", "\t") // replace 4 whitespaces with a tab
        .trim() // trim leading and trailing whitespace
        .to_string()
}

#[test]
fn test_markdown_table_conversion() {
    let slide = Slide {
        rel_path: "ppt/slides/slide1.xml".to_string(),
        slide_number: 1,
        elements: vec![
            SlideElement::Table(TableElement {
                rows: vec![
                    TableRow { cells: vec![
                        TableCell { runs: vec![Run { text: "First name".into(), formatting: Formatting::default() }]},
                        TableCell { runs: vec![Run { text: "Last name".into(), formatting: Formatting::default() }]},
                        TableCell { runs: vec![Run { text: "Age".into(), formatting: Formatting::default() }]},
                    ]},
                    TableRow { cells: vec![
                        TableCell { runs: vec![Run { text: "John".into(), formatting: Formatting::default() }]},
                        TableCell { runs: vec![Run { text: "Doe".into(), formatting: Formatting::default() }]},
                        TableCell { runs: vec![Run { text: "21".into(), formatting: Formatting::default() }]},
                    ]},
                ]
            })
        ],
        images: vec![],
        image_data: HashMap::new(),
    };
    let md_result = slide.convert_to_md().unwrap();

    let expected_md = load_test_data("table_test.md");

    assert_eq!(
        normalize_test_string(&md_result),
        normalize_test_string(&expected_md)
    );
}

#[test]
fn test_markdown_list_conversion() {
    let slide = Slide {
        rel_path: "ppt/slides/slide2.xml".to_string(),
        slide_number: 2,
        elements: vec![
            SlideElement::List(ListElement {
                items: vec![
                    ListItem { level:0, is_ordered:false, runs: vec![Run{text: "Layer 1 Element 1".into(), formatting: Formatting::default()}]},
                    ListItem { level:1, is_ordered:false, runs: vec![Run{text: "Layer 2 Element 1".into(), formatting: Formatting::default()}]},
                    ListItem { level:1, is_ordered:false, runs: vec![Run{text: "Layer 2 Element 2".into(), formatting: Formatting::default()}]},
                    ListItem { level:0, is_ordered:false, runs: vec![Run{text: "Layer 1 Element 2".into(), formatting: Formatting::default()}]},
                ]
            })
        ],
        images: vec![],
        image_data: HashMap::new(),
    };

    let md_result = slide.convert_to_md().unwrap();
    let expected_md = load_test_data("list_test.md");
    
    assert_eq!(
        normalize_test_string(&md_result),
        normalize_test_string(&expected_md)
    );
}

#[test]
fn test_formatting_conversion() {
    let slide = Slide {
        rel_path: "ppt/slides/slide1.xml".to_string(),
        slide_number: 1,
        elements: vec![
            SlideElement::Text(TextElement { runs: vec![Run { text: "bold\n".into(), formatting: Formatting { bold: true, italic: false, underlined: false, lang: "en-US".into() } }]}),
            SlideElement::Text(TextElement { runs: vec![Run { text: "cursive\n".into(), formatting: Formatting { bold: false, italic: true, underlined: false, lang: "en-US".into() } }]}),
            SlideElement::Text(TextElement { runs: vec![Run { text: "underlined\n".into(), formatting: Formatting { bold: false, italic: false, underlined: true, lang: "en-US".into() } }]}),
            SlideElement::Text(TextElement { runs: vec![Run { text: "bold and cursive\n".into(), formatting: Formatting { bold: true, italic: true, underlined: false, lang: "en-US".into() } }]}),
            SlideElement::Text(TextElement { runs: vec![Run { text: "bold, cursive and underlined\n".into(), formatting: Formatting { bold: true, italic: true, underlined: true, lang: "en-US".into() } }]}),
        ],
        images: vec![],
        image_data: HashMap::new(),
    };

    let md_result = slide.convert_to_md().unwrap();
    let expected_md = load_test_data("formatting_test.md");

    assert_eq!(
        normalize_test_string(&md_result),
        normalize_test_string(&expected_md)
    );
}