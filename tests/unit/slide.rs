use super::*;
use crate::{
    ElementPosition, Formatting, ListElement, ListItem, Run, TableCell, TableElement, TableRow,
    TextElement,
};
use std::collections::HashMap;
use std::fs;
use std::path::PathBuf;

fn load_test_data(filename: &str) -> String {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("tests");
    path.push("fixtures");
    path.push("unit");
    path.push("markdown");
    path.push(filename);
    fs::read_to_string(path).expect("Unable to read test data file")
}

fn load_binary_test_data(filename: &str) -> Vec<u8> {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("tests");
    path.push("fixtures");
    path.push("unit");
    path.push("media");
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
        blocks: vec![],
        diagnostics: vec![],
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

fn image_element(id: &str, target: &str) -> SlideElement {
    SlideElement::Image(
        ImageReference {
            id: id.to_string(),
            target: target.to_string(),
        },
        ElementPosition::default(),
    )
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
                                ..TableCell::default()
                            },
                            TableCell {
                                runs: vec![Run {
                                    text: "Last name".into(),
                                    formatting: Formatting::default(),
                                    link_target: None,
                                }],
                                ..TableCell::default()
                            },
                            TableCell {
                                runs: vec![Run {
                                    text: "Age".into(),
                                    formatting: Formatting::default(),
                                    link_target: None,
                                }],
                                ..TableCell::default()
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
                                ..TableCell::default()
                            },
                            TableCell {
                                runs: vec![Run {
                                    text: "Doe".into(),
                                    formatting: Formatting::default(),
                                    link_target: None,
                                }],
                                ..TableCell::default()
                            },
                            TableCell {
                                runs: vec![Run {
                                    text: "21".into(),
                                    formatting: Formatting::default(),
                                    link_target: None,
                                }],
                                ..TableCell::default()
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
        blocks: vec![],
        diagnostics: vec![],
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
        blocks: vec![],
        diagnostics: vec![],
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
                            lang: "en-US".into(),
                            ..Formatting::default()
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
                            italic: true,
                            lang: "en-US".into(),
                            ..Formatting::default()
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
                            underlined: true,
                            lang: "en-US".into(),
                            ..Formatting::default()
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
                            lang: "en-US".into(),
                            ..Formatting::default()
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
                            ..Formatting::default()
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
        blocks: vec![],
        diagnostics: vec![],
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
                    cells: vec![TableCell {
                        runs: vec![link()],
                        ..TableCell::default()
                    }],
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
fn escapes_raw_markdown_in_all_text_contexts() {
    let special = || Run {
        text: "# *literal* [label] \\ path\nnext".into(),
        formatting: Formatting::default(),
        link_target: None,
    };
    let mut slide = mock_slide();
    slide.config = ParserConfig::builder()
        .include_slide_number_as_comment(false)
        .include_speaker_notes(true)
        .include_comments(true)
        .build();
    slide.elements = vec![
        SlideElement::Text(
            TextElement {
                runs: vec![special()],
            },
            ElementPosition::default(),
        ),
        SlideElement::List(
            ListElement {
                items: vec![ListItem {
                    level: 0,
                    is_ordered: false,
                    runs: vec![special()],
                }],
            },
            ElementPosition::default(),
        ),
        SlideElement::Table(
            TableElement {
                rows: vec![TableRow {
                    cells: vec![TableCell {
                        runs: vec![special()],
                        ..TableCell::default()
                    }],
                }],
            },
            ElementPosition::default(),
        ),
    ];
    slide.speaker_notes = vec![TextElement {
        runs: vec![special()],
    }];
    slide.comments = vec![TextElement {
        runs: vec![special()],
    }];

    let markdown = slide.convert_to_md().expect("render markdown");
    assert_eq!(
        markdown.matches(r"\*literal\* \[label\] \\ path").count(),
        5
    );
    assert_eq!(markdown.matches(r"\# \*literal\*").count(), 4);
    assert_eq!(markdown.matches("<br>next").count(), 2);
    assert_eq!(markdown.matches("\\ path\n\nnext").count(), 1);
    assert_eq!(markdown.matches("\n> next").count(), 2);
}

#[test]
fn escapes_table_pipes_without_creating_columns() {
    let mut slide = mock_slide();
    slide.config.include_slide_number_as_comment = false;
    slide.elements = vec![SlideElement::Table(
        TableElement {
            rows: vec![TableRow {
                cells: vec![TableCell {
                    runs: vec![Run {
                        text: "left | right\r\nnext".into(),
                        formatting: Formatting::default(),
                        link_target: None,
                    }],
                    ..TableCell::default()
                }],
            }],
        },
        ElementPosition::default(),
    )];

    assert_eq!(
        slide.convert_to_md().unwrap(),
        "| left \\| right<br>next |\n| --- |\n\n"
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
        slide.convert_to_md().unwrap(),
        "> **Speaker Notes**\n>\n> First note\n> Second note\n"
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

    assert_eq!(slide.convert_to_md().unwrap(), String::new());
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
        slide.convert_to_md().unwrap(),
        "> **Speaker Notes**\n>\n> Speaker notes\n\n> **Comments**\n>\n> Comment\n"
    );
}

#[test]
fn loads_images_manually_without_compression_and_skips_missing_data() {
    let image_bytes = load_binary_test_data("example-image.jpg");
    let mut slide = mock_slide();
    slide.config = ParserConfig::builder()
        .compress_images(false)
        .image_handling_mode(ImageHandlingMode::Manually)
        .build();
    slide.elements = vec![
        image_element("present", "../media/example-image.jpg"),
        image_element("missing", "../media/missing.png"),
    ];
    slide
        .image_data
        .insert("present".to_string(), image_bytes.clone());

    let images = slide.load_images_manually().expect("load images manually");

    assert_eq!(images.len(), 1);
    assert_eq!(images[0].img_ref.id, "present");
    assert_eq!(
        images[0].base64_content,
        base64::engine::general_purpose::STANDARD.encode(image_bytes)
    );
}

#[test]
fn loads_and_compresses_images_manually() {
    let mut slide = mock_slide();
    slide.config = ParserConfig::builder()
        .compress_images(true)
        .quality(60)
        .image_handling_mode(ImageHandlingMode::Manually)
        .build();
    slide.elements = vec![image_element("image", "../media/example-image.jpg")];
    slide.image_data.insert(
        "image".to_string(),
        load_binary_test_data("example-image.jpg"),
    );

    let images = slide.load_images_manually().expect("load images manually");
    let compressed = base64::engine::general_purpose::STANDARD
        .decode(&images[0].base64_content)
        .expect("decode compressed image");

    assert_eq!(images.len(), 1);
    assert!(image::load_from_memory(&compressed).is_ok());
}

#[test]
fn invalid_image_data_cannot_be_compressed() {
    assert!(mock_slide().compress_image(b"not an image").is_none());
}

#[test]
fn save_mode_writes_the_image_and_renders_a_file_link() {
    let unique = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .expect("system time")
        .as_nanos();
    let output_dir = std::env::temp_dir().join(format!(
        "pptx-to-md-slide-save-{}-{unique}",
        std::process::id()
    ));
    let image_bytes = load_binary_test_data("example-image.jpg");
    let mut slide = mock_slide();
    slide.config = ParserConfig::builder()
        .compress_images(false)
        .image_handling_mode(ImageHandlingMode::Save)
        .image_output_path(output_dir.clone())
        .build();
    slide.elements = vec![image_element("rId1", "../media/example-image.jpg")];
    slide
        .image_data
        .insert("rId1".to_string(), image_bytes.clone());

    let markdown = slide.convert_to_md().expect("render slide");
    let saved_path = output_dir.join("slide1_image1_rId1.jpg");

    assert_eq!(
        fs::read(&saved_path).expect("read saved image"),
        image_bytes
    );
    assert!(markdown.contains("![slide1_image1_rId1.jpg](file://"));
    assert!(markdown.contains("slide1_image1_rId1.jpg)"));

    fs::remove_dir_all(output_dir).expect("remove image output directory");
}

#[test]
fn separates_multiple_elements_inside_quoted_sections() {
    let note = |text: &str| TextElement {
        runs: vec![Run {
            text: text.to_string(),
            formatting: Formatting::default(),
            link_target: None,
        }],
    };
    let mut slide = mock_slide();
    slide.config.include_speaker_notes = true;
    slide.config.include_comments = true;
    slide.speaker_notes = vec![note("First note"), note("Second note")];
    slide.comments = vec![note("First comment"), note("Second comment")];

    let markdown = slide.convert_to_md().expect("render slide");

    assert!(markdown.contains("> First note\n>\n> Second note\n"));
    assert!(markdown.contains("> First comment\n>\n> Second comment\n"));
}

fn semantic_text(text: &str, role: TextRole) -> SlideBlockContent {
    SlideBlockContent::Text(TextBlock {
        role,
        paragraphs: vec![Paragraph::plain(vec![Run {
            text: text.to_string(),
            formatting: Formatting::default(),
            link_target: None,
        }])],
    })
}

#[test]
fn semantic_renderer_uses_roles_and_configurable_reading_order() {
    let mut slide = mock_slide();
    slide.blocks = vec![
        SlideBlock {
            bounds: Bounds {
                x: 0,
                y: 0,
                width: 600,
                height: 50,
            },
            source_order: 0,
            content: semantic_text("Title", TextRole::Title),
        },
        SlideBlock {
            bounds: Bounds {
                x: 500,
                y: 100,
                width: 100,
                height: 100,
            },
            source_order: 1,
            content: semantic_text("Right column", TextRole::Body),
        },
        SlideBlock {
            bounds: Bounds {
                x: 0,
                y: 200,
                width: 100,
                height: 100,
            },
            source_order: 2,
            content: semantic_text("Left column", TextRole::Body),
        },
    ];
    let mut options = MarkdownOptions {
        include_slide_number_as_comment: false,
        ..MarkdownOptions::default()
    };

    let spatial = slide.to_markdown(&options).unwrap();
    assert!(spatial.starts_with("## Title"));
    assert!(spatial.find("Left column").unwrap() < spatial.find("Right column").unwrap());

    options.reading_order = ReadingOrder::Source;
    let source = slide.to_markdown(&options).unwrap();
    assert!(source.find("Right column").unwrap() < source.find("Left column").unwrap());
}

#[test]
fn semantic_renderer_uses_html_for_merged_tables_and_reports_unknown_blocks() {
    let mut slide = mock_slide();
    slide.blocks = vec![
        SlideBlock {
            bounds: Bounds::default(),
            source_order: 0,
            content: SlideBlockContent::Table(SemanticTable {
                rows: vec![SemanticTableRow {
                    cells: vec![SemanticTableCell {
                        paragraphs: vec![Paragraph::plain(vec![Run {
                            text: "Merged".to_string(),
                            formatting: Formatting::default(),
                            link_target: None,
                        }])],
                        row_span: 2,
                        column_span: 3,
                        covered: false,
                    }],
                }],
            }),
        },
        SlideBlock {
            bounds: Bounds::default(),
            source_order: 1,
            content: SlideBlockContent::Unsupported(UnsupportedBlock {
                kind: "chart".to_string(),
                fallback_text: Some("Revenue 2026".to_string()),
            }),
        },
    ];
    let markdown = slide
        .to_markdown(&MarkdownOptions {
            reading_order: ReadingOrder::Source,
            include_slide_number_as_comment: false,
            ..MarkdownOptions::default()
        })
        .unwrap();

    assert!(markdown.contains("<td rowspan=\"2\" colspan=\"3\">Merged</td>"));
    assert!(markdown.contains("Revenue 2026"));
    assert!(markdown.contains("<!-- Unsupported slide element: chart -->"));
}

#[test]
fn legacy_constructor_builds_semantic_blocks_for_every_element_kind() {
    let run = |text: &str| Run {
        text: text.to_string(),
        formatting: Formatting::default(),
        link_target: None,
    };
    let elements = vec![
        SlideElement::Text(
            TextElement {
                runs: vec![run("Text")],
            },
            ElementPosition { x: 1, y: 2 },
        ),
        SlideElement::List(
            ListElement {
                items: vec![ListItem {
                    level: 1,
                    is_ordered: true,
                    runs: vec![run("Second")],
                }],
            },
            ElementPosition { x: 3, y: 4 },
        ),
        SlideElement::Table(
            TableElement {
                rows: vec![TableRow {
                    cells: vec![TableCell {
                        runs: vec![run("Cell")],
                        ..TableCell::default()
                    }],
                }],
            },
            ElementPosition { x: 5, y: 6 },
        ),
        image_element("image", "../media/image.png"),
        SlideElement::Unknown,
    ];

    let slide = Slide::new(
        "ppt/slides/slide9.xml".to_string(),
        9,
        elements,
        Vec::new(),
        Vec::new(),
        Vec::new(),
        HashMap::new(),
        ParserConfig::default(),
    );

    assert_eq!(slide.blocks.len(), 5);
    let SlideBlockContent::Text(list) = &slide.blocks[1].content else {
        panic!("expected semantic list text")
    };
    assert!(matches!(
        list.paragraphs[0].list.as_ref().map(|list| &list.kind),
        Some(ListKind::Ordered { start: 1, .. })
    ));
    assert!(matches!(
        slide.blocks[2].content,
        SlideBlockContent::Table(_)
    ));
    assert!(matches!(
        slide.blocks[3].content,
        SlideBlockContent::Image(_)
    ));
    assert!(matches!(
        slide.blocks[4].content,
        SlideBlockContent::Unsupported(_)
    ));
}

#[test]
fn spatial_order_handles_dimensionless_blocks_and_full_width_separators() {
    let block = |text: &str, bounds: Bounds, source_order: usize| SlideBlock {
        bounds,
        source_order,
        content: semantic_text(text, TextRole::Body),
    };
    let mut slide = mock_slide();
    let options = MarkdownOptions {
        include_slide_number_as_comment: false,
        ..MarkdownOptions::default()
    };

    slide.blocks = vec![
        block(
            "Second",
            Bounds {
                x: 20,
                y: 10,
                width: 0,
                height: 0,
            },
            0,
        ),
        block(
            "First",
            Bounds {
                x: 10,
                y: 20,
                width: 0,
                height: 0,
            },
            1,
        ),
    ];
    let dimensionless = slide.to_markdown(&options).unwrap();
    assert!(dimensionless.find("Second").unwrap() < dimensionless.find("First").unwrap());

    slide.blocks = vec![
        block(
            "Right above",
            Bounds {
                x: 500,
                y: 100,
                width: 100,
                height: 50,
            },
            0,
        ),
        block(
            "Left above",
            Bounds {
                x: 0,
                y: 150,
                width: 100,
                height: 50,
            },
            1,
        ),
        block(
            "Separator",
            Bounds {
                x: 0,
                y: 300,
                width: 600,
                height: 40,
            },
            2,
        ),
        block(
            "Right below",
            Bounds {
                x: 500,
                y: 400,
                width: 100,
                height: 50,
            },
            3,
        ),
        block(
            "Left below",
            Bounds {
                x: 0,
                y: 450,
                width: 100,
                height: 50,
            },
            4,
        ),
    ];
    let spatial = slide.to_markdown(&options).unwrap();
    assert!(spatial.find("Left above").unwrap() < spatial.find("Right above").unwrap());
    assert!(spatial.find("Right above").unwrap() < spatial.find("Separator").unwrap());
    assert!(spatial.find("Separator").unwrap() < spatial.find("Left below").unwrap());
    assert!(spatial.find("Left below").unwrap() < spatial.find("Right below").unwrap());
}

#[test]
fn semantic_image_rendering_covers_missing_manual_and_invalid_compression_paths() {
    let image = SlideBlock {
        bounds: Bounds::default(),
        source_order: 0,
        content: SlideBlockContent::Image(ImageBlock {
            reference: ImageReference {
                id: "image".to_string(),
                target: "../media/image.png".to_string(),
            },
            alt_text: Some("Diagram".to_string()),
            mime_type: Some("image/png".to_string()),
        }),
    };
    let options = MarkdownOptions {
        include_slide_number_as_comment: false,
        ..MarkdownOptions::default()
    };

    let mut slide = mock_slide();
    slide.blocks = vec![image.clone()];
    assert!(
        slide
            .to_markdown(&options)
            .unwrap()
            .contains("Image unavailable: Diagram")
    );

    slide.config.image_handling_mode = ImageHandlingMode::Manually;
    assert_eq!(slide.to_markdown(&options).unwrap(), "\n");

    slide.config.image_handling_mode = ImageHandlingMode::InMarkdown;
    slide.config.compress_images = true;
    slide
        .image_data
        .insert("image".to_string(), b"invalid image".to_vec());
    assert!(
        slide
            .to_markdown(&options)
            .unwrap()
            .contains("Image unavailable: Diagram")
    );

    slide.config.image_handling_mode = ImageHandlingMode::Save;
    assert!(
        slide
            .to_markdown(&options)
            .unwrap()
            .contains("Image unavailable: Diagram")
    );
}
