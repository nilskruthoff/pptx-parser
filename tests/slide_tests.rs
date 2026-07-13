use pptx_to_md::ImageReference;
use pptx_to_md::{
    ElementPosition, Formatting, ListElement, ListItem, ParserConfig, Run, Slide, SlideElement,
    TableCell, TableElement, TableRow, TextElement,
};
use std::collections::HashMap;
use std::fs;
use std::path::PathBuf;

fn load_test_data(filename: &str) -> String {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("tests");
    path.push("test_data");
    path.push(filename);
    fs::read_to_string(path).expect("Unable to read test data file")
}

fn load_binary_test_data(filename: &str) -> Vec<u8> {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("tests");
    path.push("test_data");
    path.push(filename);
    fs::read(path).expect("Unable to read test data file")
}

fn mock_slide() -> Slide {
    Slide {
        rel_path: "ppt/slides/slide1.xml".to_string(),
        slide_number: 1,
        elements: vec![],
        speaker_notes: vec![],
        comments: vec![],
        images: vec![],
        image_data: HashMap::new(),
        config: ParserConfig::default(),
    }
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
        elements: vec![SlideElement::Table(
            TableElement {
                rows: vec![
                    TableRow {
                        cells: vec![
                            TableCell {
                                runs: vec![Run {
                                    text: "First name".into(),
                                    formatting: Formatting::default(),
                                    link_target: None,
                                }],
                            },
                            TableCell {
                                runs: vec![Run {
                                    text: "Last name".into(),
                                    formatting: Formatting::default(),
                                    link_target: None,
                                }],
                            },
                            TableCell {
                                runs: vec![Run {
                                    text: "Age".into(),
                                    formatting: Formatting::default(),
                                    link_target: None,
                                }],
                            },
                        ],
                    },
                    TableRow {
                        cells: vec![
                            TableCell {
                                runs: vec![Run {
                                    text: "John".into(),
                                    formatting: Formatting::default(),
                                    link_target: None,
                                }],
                            },
                            TableCell {
                                runs: vec![Run {
                                    text: "Doe".into(),
                                    formatting: Formatting::default(),
                                    link_target: None,
                                }],
                            },
                            TableCell {
                                runs: vec![Run {
                                    text: "21".into(),
                                    formatting: Formatting::default(),
                                    link_target: None,
                                }],
                            },
                        ],
                    },
                ],
            },
            ElementPosition::default(),
        )],
        speaker_notes: vec![],
        comments: vec![],
        images: vec![],
        image_data: HashMap::new(),
        config: ParserConfig::default(),
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
        elements: vec![SlideElement::List(
            ListElement {
                items: vec![
                    ListItem {
                        level: 0,
                        is_ordered: false,
                        runs: vec![Run {
                            text: "Layer 1 Element 1".into(),
                            formatting: Formatting::default(),
                            link_target: None,
                        }],
                    },
                    ListItem {
                        level: 1,
                        is_ordered: false,
                        runs: vec![Run {
                            text: "Layer 2 Element 1".into(),
                            formatting: Formatting::default(),
                            link_target: None,
                        }],
                    },
                    ListItem {
                        level: 1,
                        is_ordered: false,
                        runs: vec![Run {
                            text: "Layer 2 Element 2".into(),
                            formatting: Formatting::default(),
                            link_target: None,
                        }],
                    },
                    ListItem {
                        level: 0,
                        is_ordered: false,
                        runs: vec![Run {
                            text: "Layer 1 Element 2".into(),
                            formatting: Formatting::default(),
                            link_target: None,
                        }],
                    },
                ],
            },
            ElementPosition::default(),
        )],
        speaker_notes: vec![],
        comments: vec![],
        images: vec![],
        image_data: HashMap::new(),
        config: ParserConfig::default(),
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
            SlideElement::Text(
                TextElement {
                    runs: vec![Run {
                        text: "bold\n".into(),
                        formatting: Formatting {
                            bold: true,
                            italic: false,
                            underlined: false,
                            lang: "en-US".into(),
                        },
                        link_target: None,
                    }],
                },
                ElementPosition::default(),
            ),
            SlideElement::Text(
                TextElement {
                    runs: vec![Run {
                        text: "cursive\n".into(),
                        formatting: Formatting {
                            bold: false,
                            italic: true,
                            underlined: false,
                            lang: "en-US".into(),
                        },
                        link_target: None,
                    }],
                },
                ElementPosition::default(),
            ),
            SlideElement::Text(
                TextElement {
                    runs: vec![Run {
                        text: "underlined\n".into(),
                        formatting: Formatting {
                            bold: false,
                            italic: false,
                            underlined: true,
                            lang: "en-US".into(),
                        },
                        link_target: None,
                    }],
                },
                ElementPosition::default(),
            ),
            SlideElement::Text(
                TextElement {
                    runs: vec![Run {
                        text: "bold and cursive\n".into(),
                        formatting: Formatting {
                            bold: true,
                            italic: true,
                            underlined: false,
                            lang: "en-US".into(),
                        },
                        link_target: None,
                    }],
                },
                ElementPosition::default(),
            ),
            SlideElement::Text(
                TextElement {
                    runs: vec![Run {
                        text: "bold, cursive and underlined\n".into(),
                        formatting: Formatting {
                            bold: true,
                            italic: true,
                            underlined: true,
                            lang: "en-US".into(),
                        },
                        link_target: None,
                    }],
                },
                ElementPosition::default(),
            ),
        ],
        speaker_notes: vec![],
        comments: vec![],
        images: vec![],
        image_data: HashMap::new(),
        config: ParserConfig::default(),
    };

    let md_result = slide.convert_to_md().unwrap();
    let expected_md = load_test_data("formatting_test.md");

    assert_eq!(
        normalize_test_string(&md_result),
        normalize_test_string(&expected_md)
    );
}

#[test]
fn renders_links_in_all_text_markdown_contexts() {
    let link = || Run {
        text: "Example".into(),
        formatting: Formatting::default(),
        link_target: Some("https://example.com".into()),
    };
    let mut slide = mock_slide();
    slide.elements = vec![
        SlideElement::Text(
            TextElement { runs: vec![link()] },
            ElementPosition::default(),
        ),
        SlideElement::List(
            ListElement {
                items: vec![ListItem {
                    level: 0,
                    is_ordered: false,
                    runs: vec![link()],
                }],
            },
            ElementPosition::default(),
        ),
        SlideElement::Table(
            TableElement {
                rows: vec![TableRow {
                    cells: vec![TableCell { runs: vec![link()] }],
                }],
            },
            ElementPosition::default(),
        ),
    ];
    slide.speaker_notes = vec![TextElement { runs: vec![link()] }];
    slide.comments = vec![TextElement { runs: vec![link()] }];
    slide.config = ParserConfig::builder()
        .include_speaker_notes(true)
        .include_comments(true)
        .build();

    let markdown = slide.convert_to_md().expect("render markdown");
    assert_eq!(
        markdown.matches("[Example](https://example.com)").count(),
        5
    );
}

#[test]
fn extracts_slide_number_from_path() {
    assert_eq!(
        Slide::extract_slide_number("ppt/slides/slide5.xml"),
        Some(5)
    );
}

#[test]
fn extracts_image_extension() {
    assert_eq!(
        mock_slide().get_image_extension("../media/image1.png"),
        "png"
    );
}

#[test]
fn links_image_elements_to_relationship_targets() {
    let mut slide = mock_slide();
    slide.images.push(ImageReference {
        id: "rId2".to_string(),
        target: "../media/image1.png".to_string(),
    });
    slide.elements.push(SlideElement::Image(
        ImageReference {
            id: "rId2".to_string(),
            target: String::new(),
        },
        ElementPosition::default(),
    ));

    slide.link_images();

    let SlideElement::Image(reference, _) = &slide.elements[0] else {
        panic!("expected image element");
    };
    assert_eq!(reference.target, "../media/image1.png");
}

#[test]
fn compresses_images_to_smaller_valid_jpegs() {
    let mut slide = mock_slide();
    slide.config.quality = 50;
    let raw_image = load_binary_test_data("example-image.jpg");

    let compressed = slide.compress_image(&raw_image).expect("compress image");
    assert!(compressed.len() < raw_image.len());
    assert!(image::load_from_memory(&compressed).is_ok());
}

#[test]
fn renders_speaker_notes_as_markdown_blockquotes_when_enabled() {
    let mut slide = mock_slide();
    slide.config.include_slide_number_as_comment = false;
    slide.config.include_speaker_notes = true;
    slide.speaker_notes = vec![TextElement {
        runs: vec![
            Run {
                text: "First note\n".to_string(),
                formatting: Formatting::default(),
                link_target: None,
            },
            Run {
                text: "Second note".to_string(),
                formatting: Formatting::default(),
                link_target: None,
            },
        ],
    }];

    assert_eq!(
        slide.convert_to_md(),
        Some("> **Speaker Notes**\n>\n> First note\n> Second note\n".to_string())
    );
}

#[test]
fn does_not_render_speaker_notes_by_default() {
    let mut slide = mock_slide();
    slide.config.include_slide_number_as_comment = false;
    slide.speaker_notes = vec![TextElement {
        runs: vec![Run {
            text: "Hidden note".to_string(),
            formatting: Formatting::default(),
            link_target: None,
        }],
    }];

    assert_eq!(slide.convert_to_md(), Some(String::new()));
}

#[test]
fn renders_comments_separately_from_speaker_notes() {
    let mut slide = mock_slide();
    slide.config.include_slide_number_as_comment = false;
    slide.config.include_speaker_notes = true;
    slide.config.include_comments = true;
    slide.speaker_notes = vec![TextElement {
        runs: vec![Run {
            text: "Speaker notes".to_string(),
            formatting: Formatting::default(),
            link_target: None,
        }],
    }];
    slide.comments = vec![TextElement {
        runs: vec![Run {
            text: "Comment".to_string(),
            formatting: Formatting::default(),
            link_target: None,
        }],
    }];

    assert_eq!(
        slide.convert_to_md(),
        Some(
            "> **Speaker Notes**\n>\n> Speaker notes\n\n> **Comments**\n>\n> Comment\n".to_string()
        )
    );
}
