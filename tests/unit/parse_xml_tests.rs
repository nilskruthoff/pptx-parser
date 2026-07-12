    use super::*;
    use std::fs;
    use std::path::PathBuf;

    fn load_xml(filename: &str) -> Option<String> {
        let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        path.push("tests");
        path.push("test_data");
        path.push("xml");
        path.push(filename);
        fs::read_to_string(path).ok()
    }
    macro_rules! load_xml_or_skip {
        ($filename:expr) => {
            match load_xml($filename) {
                Some(xml) => xml,
                None => return,
            }
        };
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
    fn test_parse_text() {
        let xml_data = load_xml_or_skip!("tx_body.xml");
        let doc = Document::parse(&*xml_data).expect("Parsing XML failes");
        let tx_body_node = doc.root_element();

        match parse_text(&tx_body_node) {
            Ok(text_element) => {
                assert_eq!(text_element.runs.len(), 3);
                assert_eq!(normalize_test_string(&text_element.runs[0].text), normalize_test_string("Hello"));
                assert_eq!(normalize_test_string(&text_element.runs[1].text), normalize_test_string("World"));
                assert_eq!(normalize_test_string(&text_element.runs[2].text), normalize_test_string("!"));
            },
            Err(_) => panic!("Fehler beim Parsen der XML-Datei"),
        }
    }

    #[test]
    fn test_parse_run_with_format() {
        let xml_data = load_xml_or_skip!("run_styles.xml");
        let doc = Document::parse(&*xml_data).expect("Parsing XML failed");
        let r_node = doc.root_element();

        match parse_run(&r_node) {
            Ok(run) => {
                assert_eq!(normalize_test_string(&run.text), normalize_test_string("Formatted text"));
                assert!(run.formatting.bold);
                assert!(run.formatting.italic);
                assert!(run.formatting.underlined);
                assert_eq!(run.formatting.lang, "de-DE");
            },
            Err(_) => panic!("Fehler beim Parsen des Runs mit Formatierung")
        }
    }

    #[test]
    fn test_parse_run_no_format() {
        let xml_data = load_xml_or_skip!("run_no_format.xml");
        let doc = Document::parse(&*xml_data).expect("Parsing XML failed");
        let r_node = doc.root_element();

        match parse_run(&r_node) {
            Ok(run) => {
                assert_eq!(normalize_test_string(&run.text), normalize_test_string("Unformatted text"));
                assert!(!run.formatting.bold);
                assert!(!run.formatting.italic);
                assert!(!run.formatting.underlined);
            },
            Err(_) => panic!("Fehler beim Parsen des Runs ohne Formatierung")
        }
    }
    
    #[test]
    fn test_parse_run_empty_text() {
        let xml_data = load_xml_or_skip!("run_empty.xml");
        let doc = Document::parse(&*xml_data).expect("Parsing XML failed");
        let r_node = doc.root_element();

        match parse_run(&r_node) {
            Ok(run) => {
                assert_eq!(run.text, "");
            },
            Err(_) => panic!("Failed to parse an empty Run")
        }
    }

    #[test]
    fn test_parse_paragraph_single() {
        let xml_data = load_xml_or_skip!("paragraph_single.xml");
        let doc = Document::parse(&*xml_data).expect("Parsing XML failed");
        let p_node = doc.root_element();

        match parse_paragraph(&p_node, true) {
            Ok(runs) => {
                assert_eq!(runs.len(), 1);
                assert_eq!(normalize_test_string(&runs[0].text), normalize_test_string("Single run\n"));
            },
            Err(_) => panic!("Failed to parse paragraph with a single run")
        }
    }

    #[test]
    fn test_parse_paragraph_multiple() {
        let xml_data = load_xml_or_skip!("paragraph_multiple.xml");
        let doc = Document::parse(&*xml_data).expect("Parsing XML failed");
        let p_node = doc.root_element();

        match parse_paragraph(&p_node, true) {
            Ok(runs) => {
                assert_eq!(runs.len(), 3);
                assert_eq!(normalize_test_string(&runs[0].text), normalize_test_string("First run"));
                assert_eq!(normalize_test_string(&runs[1].text), normalize_test_string("Second run"));
                assert_eq!(normalize_test_string(&runs[2].text), normalize_test_string("Third run\n"));
                assert!(runs[1].formatting.bold);
                assert!(runs[2].formatting.italic);
            },
            Err(_) => panic!("Failed to parse paragraph with multiple runs (`add_new_line: true)")
        }

        match parse_paragraph(&p_node, false) {
            Ok(runs) => {
                assert_eq!(runs.len(), 3);
                assert!(!runs[2].text.ends_with('\n'));
            },
            Err(_) => panic!("Failed to parse paragraph with multiple runs (`add_new_line: false)`")
        }
    }

    #[test]
    fn test_parse_paragraph_empty() {
        let xml_data = load_xml_or_skip!("paragraph_empty.xml");
        let doc = Document::parse(&*xml_data).expect("Parsing XML failed");
        let p_node = doc.root_element();

        match parse_paragraph(&p_node, true) {
            Ok(runs) => {
                assert_eq!(runs.len(), 0);
            },
            Err(_) => panic!("Failed to parse paragraph with empty runs")
        }
    }

    #[test]
    fn test_parse_list_properties_unordered() {
        // Test for unordered list properties
        let xml_data = load_xml_or_skip!("simple_list.xml");
        let doc = Document::parse(&*xml_data).expect("Failed to parse XML");

        let p_node = doc.root_element()
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "p")
            .expect("No paragraph element found");

        match parse_list_properties(&p_node) {
            Ok((level, is_ordered)) => {
                assert_eq!(level, 0, "List level should be 0");
                assert!(is_ordered, "List should be identified as ordered due to buChar element");
            },
            Err(_) => panic!("Failed to parse list properties")
        }
    }

    #[test]
    fn test_parse_list_properties_ordered() {
        // Test for ordered list properties
        let xml_data = load_xml_or_skip!("multilevel_list.xml");
        let doc = Document::parse(&*xml_data).expect("Failed to parse XML");

        // Get the first paragraph (level 0 with buAutoNum)
        let p_node = doc.root_element()
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "p")
            .expect("No paragraph element found");

        match parse_list_properties(&p_node) {
            Ok((level, is_ordered)) => {
                assert_eq!(level, 0, "List level should be 0");
                assert!(is_ordered, "List should be identified as ordered due to buAutoNum element");
            },
            Err(_) => panic!("Failed to parse ordered list properties")
        }

        // Get the second paragraph (level 1 with buChar)
        let p_node = doc.root_element()
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "p")
            .nth(1)
            .expect("Second paragraph element not found");

        match parse_list_properties(&p_node) {
            Ok((level, is_ordered)) => {
                assert_eq!(level, 1, "List level should be 1");
                assert!(is_ordered, "List should be identified as ordered due to buChar element");
            },
            Err(_) => panic!("Failed to parse level 1 list properties")
        }

        // Get the fourth paragraph (level 2 with buAutoNum)
        let p_node = doc.root_element()
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "p")
            .nth(3)
            .expect("Fourth paragraph element not found");

        match parse_list_properties(&p_node) {
            Ok((level, is_ordered)) => {
                assert_eq!(level, 2, "List level should be 2");
                assert!(is_ordered, "Level 2 list should be identified as ordered");
            },
            Err(_) => panic!("Failed to parse level 2 list properties")
        }
    }

    #[test]
    fn test_parse_simple_list() {
        // Test for parsing a complete simple list
        let xml_data = load_xml_or_skip!("simple_list.xml");
        let doc = Document::parse(&*xml_data).expect("Failed to parse XML");
        let tx_body_node = doc.root_element();

        match parse_list(&tx_body_node) {
            Ok(list) => {
                assert_eq!(list.items.len(), 3, "List should have 3 items");

                // Check the first item
                assert_eq!(list.items[0].level, 0, "First item should be level 0");
                assert!(list.items[0].is_ordered, "First item should be ordered (has buChar)");
                assert_eq!(normalize_test_string(&list.items[0].runs[0].text), normalize_test_string("First item\n"), "First item text mismatch");

                // Check the second item
                assert_eq!(list.items[1].level, 0, "Second item should be level 0");
                assert!(list.items[1].is_ordered, "Second item should be ordered (has buChar)");
                assert_eq!(normalize_test_string(&list.items[1].runs[0].text), normalize_test_string("Second item\n"), "Second item text mismatch");

                // Check the third item
                assert_eq!(list.items[2].level, 0, "Third item should be level 0");
                assert!(list.items[2].is_ordered, "Third item should be ordered (has buChar)");
                assert_eq!(normalize_test_string(&list.items[2].runs[0].text), normalize_test_string("Third item\n"), "Third item text mismatch");
            },
            Err(_) => panic!("Failed to parse simple list")
        }
    }

    #[test]
    fn test_parse_multilevel_list() {
        // Test for parsing a multilevel list
        let xml_data = load_xml_or_skip!("multilevel_list.xml");
        let doc = Document::parse(&*xml_data).expect("Failed to parse XML");
        let tx_body_node = doc.root_element();

        match parse_list(&tx_body_node) {
            Ok(list) => {
                assert_eq!(list.items.len(), 5, "List should have 5 items");

                // Check first item (level 0, ordered)
                assert_eq!(list.items[0].level, 0, "First item should be level 0");
                assert!(list.items[0].is_ordered, "First item should be ordered");
                assert_eq!(normalize_test_string(&list.items[0].runs[0].text), normalize_test_string("Main topic\n"), "First item text mismatch");

                // Check second item (level 1, unordered but detected as ordered due to buChar)
                assert_eq!(list.items[1].level, 1, "Second item should be level 1");
                assert!(list.items[1].is_ordered, "Second item should be detected as ordered due to buChar");
                assert_eq!(normalize_test_string(&list.items[1].runs[0].text), normalize_test_string("Subtopic bullet\n"), "Second item text mismatch");

                // Check fourth item (level 2, ordered)
                assert_eq!(list.items[3].level, 2, "Fourth item should be level 2");
                assert!(list.items[3].is_ordered, "Fourth item should be ordered");
                assert_eq!(normalize_test_string(&list.items[3].runs[0].text), normalize_test_string("Numbered sub-subtopic\n"), "Fourth item text mismatch");

                // Check fifth item (back to level 0)
                assert_eq!(list.items[4].level, 0, "Fifth item should be level 0");
                assert!(list.items[4].is_ordered, "Fifth item should be ordered");
                assert_eq!(normalize_test_string(&list.items[4].runs[0].text), normalize_test_string("Second main topic\n"), "Fifth item text mismatch");
            },
            Err(_) => panic!("Failed to parse multilevel list")
        }
    }

    /// Test for a simple table for a cell with a single paragraph
    #[test]
    fn test_parse_table_cell_simple() {
        let xml_data = load_xml_or_skip!("simple_table.xml");
        let doc = Document::parse(&*xml_data).expect("Parsing XML failed");

        let tc_node = doc.root_element()
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "tc")
            .expect("Couldn't find tc node");

        match parse_table_cell(&tc_node) {
            Ok(cell) => {
                assert_eq!(cell.runs.len(), 1);
                assert_eq!(normalize_test_string(&cell.runs[0].text), normalize_test_string("Cell 1,1"));
            },
            Err(_) => panic!("Failed to parse the table cell")
        }
    }

    /// Test for a complex table with multiple paragraphs in a table cell
    #[test]
    fn test_parse_table_cell_complex() {
        let xml_data = load_xml_or_skip!("complex_table.xml");
        let doc = Document::parse(&*xml_data).expect("Parsing XML failed");

        // second row, first cell
        let tc_node = doc.root_element()
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "tc")
            .nth(3)
            .expect("Failed to find table cell with multiple paragraphs");

        match parse_table_cell(&tc_node) {
            Ok(cell) => {
                assert_eq!(cell.runs.len(), 3);
                assert_eq!(normalize_test_string(&cell.runs[0].text), normalize_test_string("Multiple"));
                assert_eq!(normalize_test_string(&cell.runs[1].text), normalize_test_string("paragraphs"));
                assert_eq!(normalize_test_string(&cell.runs[2].text), normalize_test_string("in one cell"));
            },
            Err(_) => panic!("Failed to parse table cell with multiple paragraphs")
        }
    }
    #[test]
    fn test_parse_table_cell_empty() {
        let xml_data = load_xml_or_skip!("empty_table.xml");
        let doc = Document::parse(&*xml_data).expect("Parsing XML failed");

        let tc_node = doc.root_element()
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "tc")
            .expect("Failed to find empty table cell");

        match parse_table_cell(&tc_node) {
            Ok(cell) => {
                assert_eq!(cell.runs.len(), 0);
            },
            Err(_) => panic!("Failed to parse empty table cell")
        }
    }

    #[test]
    fn test_parse_table_row_simple() {
        let xml_data = load_xml_or_skip!("simple_table.xml");
        let doc = Document::parse(&*xml_data).expect("Parsing XML failed");

        let tr_node = doc.root_element()
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "tr")
            .expect("Couldn't find tc node");

        match parse_table_row(&tr_node) {
            Ok(row) => {
                assert_eq!(row.cells.len(), 2);
                assert_eq!(normalize_test_string(&row.cells[0].runs[0].text), normalize_test_string("Cell 1,1"));
                assert_eq!(normalize_test_string(&row.cells[1].runs[0].text), normalize_test_string("Cell 1,2"));
            },
            Err(_) => panic!("Failed to parse the table row")
        }
    }

    #[test]
    fn test_parse_table_row_complex() {
        let xml_data = load_xml_or_skip!("complex_table.xml");
        let doc = Document::parse(&*xml_data).expect("Parsing XML failed");

        let tr_node = doc.root_element()
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "tr")
            .nth(0) // Erste Zeile mit fetten Überschriften
            .expect("Couldn't find a table row with formatting");

        match parse_table_row(&tr_node) {
            Ok(row) => {
                assert_eq!(row.cells.len(), 3);
                for i in 0..3 {
                    assert!(row.cells[i].runs[0].formatting.bold);
                    assert!(normalize_test_string(&row.cells[i].runs[0].text).starts_with("Heading"));
                }
            },
            Err(_) => panic!("Failed to parse a table row with formatting")
        }
    }

    #[test]
    fn test_parse_simple_table() {
        // Test for a simple table with 2x2 structure
        let xml_data = load_xml_or_skip!("simple_table.xml");
        let doc = Document::parse(&*xml_data).expect("Failed to parse XML");

        let tbl_node = doc.root_element()
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "tbl")
            .expect("No table element found");

        match parse_table(&tbl_node) {
            Ok(table) => {
                assert_eq!(table.rows.len(), 2, "Table should have 2 rows");
                assert_eq!(table.rows[0].cells.len(), 2, "First row should have 2 cells");
                assert_eq!(table.rows[1].cells.len(), 2, "Second row should have 2 cells");

                // Check contents of the first row
                assert_eq!(normalize_test_string(&table.rows[0].cells[0].runs[0].text), normalize_test_string("Cell 1,1"), "Cell content mismatch");
                assert_eq!(normalize_test_string(&table.rows[0].cells[1].runs[0].text), normalize_test_string("Cell 1,2"), "Cell content mismatch");

                // Check contents of the second row
                assert_eq!(normalize_test_string(&table.rows[1].cells[0].runs[0].text), normalize_test_string("Cell 2,1"), "Cell content mismatch");
                assert_eq!(normalize_test_string(&table.rows[1].cells[1].runs[0].text), normalize_test_string("Cell 2,2"), "Cell content mismatch");
            },
            Err(_) => panic!("Failed to parse table structure")
        }
    }

    #[test]
    fn test_parse_complex_table() {
        // Test for a complex table with different formatting and multiple paragraphs
        let xml_data = load_xml_or_skip!("complex_table.xml");
        let doc = Document::parse(&*xml_data).expect("Failed to parse XML");

        let tbl_node = doc.root_element()
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "tbl")
            .expect("No table element found");

        match parse_table(&tbl_node) {
            Ok(table) => {
                assert_eq!(table.rows.len(), 2, "Table should have 2 rows");
                assert_eq!(table.rows[0].cells.len(), 3, "First row should have 3 cells");
                assert_eq!(table.rows[1].cells.len(), 3, "Second row should have 3 cells");

                // Check bold formatting in headers
                for i in 0..3 {
                    assert!(table.rows[0].cells[i].runs[0].formatting.bold, "Header cell should have bold formatting");
                    assert!(normalize_test_string(&table.rows[0].cells[i].runs[0].text).starts_with("Heading"), "Header should start with 'Heading'");
                }

                // Check the cell with multiple paragraphs
                assert_eq!(table.rows[1].cells[0].runs.len(), 3);
                assert_eq!(normalize_test_string(&table.rows[1].cells[0].runs[0].text), normalize_test_string("Multiple"), "First paragraph content mismatch");
                assert_eq!(normalize_test_string(&table.rows[1].cells[0].runs[1].text), normalize_test_string("paragraphs"), "Second paragraph content mismatch");
                assert_eq!(normalize_test_string(&table.rows[1].cells[0].runs[2].text), normalize_test_string("in one cell"), "Third paragraph content mismatch");

                // Check the cell with italic text
                assert!(table.rows[1].cells[1].runs[0].formatting.italic, "Text should have italic formatting");
                assert_eq!(normalize_test_string(&table.rows[1].cells[1].runs[0].text), normalize_test_string("Cursive"), "Italic text content mismatch");
            },
            Err(_) => panic!("Failed to parse complex table structure")
        }
    }

    #[test]
    fn test_parse_empty_table() {
        // Test for a table with empty cells
        let xml_data = load_xml_or_skip!("empty_table.xml");
        let doc = Document::parse(&*xml_data).expect("Failed to parse XML");

        let tbl_node = doc.root_element()
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "tbl")
            .expect("No table element found");

        match parse_table(&tbl_node) {
            Ok(table) => {
                assert_eq!(table.rows.len(), 2, "Table should have 2 rows");
                assert_eq!(table.rows[0].cells.len(), 2, "First row should have 2 cells");
                assert_eq!(table.rows[1].cells.len(), 2, "Second row should have 2 cells");

                // Check that empty cells have no runs
                assert_eq!(table.rows[0].cells[0].runs.len(), 0, "Empty cell should have no runs");
                assert_eq!(table.rows[0].cells[1].runs.len(), 0, "Empty cell should have no runs");
                assert_eq!(table.rows[1].cells[0].runs.len(), 0, "Empty cell should have no runs");

                // Check the one cell with content
                assert_eq!(table.rows[1].cells[1].runs.len(), 1, "Cell should have one run");
                assert_eq!(normalize_test_string(&table.rows[1].cells[1].runs[0].text), normalize_test_string("Only content"), "Cell content mismatch");
            },
            Err(_) => panic!("Failed to parse table with empty cells")
        }
    }

    #[test]
    fn test_parse_graphic_frame_with_table() {
        // Test for a graphic frame containing a table
        let xml_data = load_xml_or_skip!("simple_table.xml");
        let doc = Document::parse(&*xml_data).expect("Failed to parse XML");
        let node = doc.root_element();

        match parse_graphic_frame(&node) {
            Ok(Some(table)) => {
                assert_eq!(table.rows.len(), 2, "Table should have 2 rows");
                assert_eq!(table.rows[0].cells.len(), 2, "First row should have 2 cells");

                // Basic content check to confirm we got the right table
                assert_eq!(normalize_test_string(&table.rows[0].cells[0].runs[0].text), normalize_test_string("Cell 1,1"), "Cell content mismatch");
            },
            Ok(None) => panic!("Should have found a table, but got None"),
            Err(_) => panic!("Failed to parse graphic frame with table")
        }
    }

    #[test]
    fn test_parse_graphic_frame_without_table() {
        // Test for a graphic frame that doesn't contain a table
        let xml_data = load_xml_or_skip!("non_table_graphic.xml");
        let doc = Document::parse(&*xml_data).expect("Failed to parse XML");
        let node = doc.root_element();

        match parse_graphic_frame(&node) {
            Ok(None) => {
                // This is the expected result - no table found
            },
            Ok(Some(_)) => panic!("Found a table where none should exist"),
            Err(_) => panic!("Failed to parse non-table graphic frame")
        }
    }

    #[test]
    fn test_parse_pic_with_image() {
        // Test for parsing a picture with a valid image reference
        let xml_data = load_xml_or_skip!("pic_with_image.xml");
        let doc = Document::parse(&*xml_data).expect("Failed to parse XML");
        let pic_node = doc.root_element();

        match parse_pic(&pic_node) {
            Ok(image_ref) => {
                assert_eq!(image_ref.id, "rId2", "Image reference ID should be 'rId2'");
                assert_eq!(image_ref.target, "", "Image target should be empty initially");
            },
            Err(e) => panic!("Failed to parse picture: {:?}", e)
        }
    }

    #[test]
    fn test_parse_pic_without_embed() {
        // Test for parsing a picture without an embed attribute
        let xml_data = load_xml_or_skip!("pic_without_embed.xml");
        let doc = Document::parse(&*xml_data).expect("Failed to parse XML");
        let pic_node = doc.root_element();

        match parse_pic(&pic_node) {
            Ok(_) => panic!("Should have failed due to missing embed attribute"),
            Err(Error::ImageNotFound) => {
                // This is the expected behavior - should fail with ImageNotFound
            },
            Err(e) => panic!("Expected ImageNotFound error but got: {:?}", e)
        }
    }

    #[test]
    fn test_parse_pic_without_blip() {
        // Test for parsing a picture without a blip node
        let xml_data = load_xml_or_skip!("pic_without_blip.xml");
        let doc = Document::parse(&*xml_data).expect("Failed to parse XML");
        let pic_node = doc.root_element();

        match parse_pic(&pic_node) {
            Ok(_) => panic!("Should have failed due to missing blip node"),
            Err(Error::ImageNotFound) => {
                // This is the expected behavior - should fail with ImageNotFound
            },
            Err(e) => panic!("Expected ImageNotFound error but got: {:?}", e)
        }
    }

    #[test]
    fn test_parse_slide_xml_reads_graphic_frame_position_from_p_xfrm() {
        let xml = r#"
            <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                   xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:cSld>
                <p:spTree>
                  <p:nvGrpSpPr/>
                  <p:grpSpPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="0" cy="0"/>
                      <a:chOff x="0" y="0"/>
                      <a:chExt cx="0" cy="0"/>
                    </a:xfrm>
                  </p:grpSpPr>
                  <p:graphicFrame>
                    <p:nvGraphicFramePr/>
                    <p:xfrm>
                      <a:off x="111" y="222"/>
                      <a:ext cx="333" cy="444"/>
                    </p:xfrm>
                    <a:graphic>
                      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                        <a:tbl>
                          <a:tr h="1">
                            <a:tc>
                              <a:txBody>
                                <a:bodyPr/>
                                <a:lstStyle/>
                                <a:p><a:r><a:t>Cell</a:t></a:r></a:p>
                              </a:txBody>
                              <a:tcPr/>
                            </a:tc>
                          </a:tr>
                        </a:tbl>
                      </a:graphicData>
                    </a:graphic>
                  </p:graphicFrame>
                </p:spTree>
              </p:cSld>
            </p:sld>
        "#;

        let elements = parse_slide_xml(xml.as_bytes()).expect("Failed to parse slide XML");

        let table = elements.iter().find(|element| matches!(element, SlideElement::Table(_, _)))
            .expect("Expected table element");

        match table {
            SlideElement::Table(_, position) => assert_eq!(*position, ElementPosition { x: 111, y: 222 }),
            element => panic!("Expected table element, got {:?}", element),
        }
    }

    #[test]
    fn test_parse_slide_xml_applies_group_transform_to_child_positions() {
        let xml = r#"
            <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                   xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:cSld>
                <p:spTree>
                  <p:nvGrpSpPr/>
                  <p:grpSpPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="0" cy="0"/>
                      <a:chOff x="0" y="0"/>
                      <a:chExt cx="0" cy="0"/>
                    </a:xfrm>
                  </p:grpSpPr>
                  <p:grpSp>
                    <p:nvGrpSpPr/>
                    <p:grpSpPr>
                      <a:xfrm>
                        <a:off x="1000" y="2000"/>
                        <a:ext cx="4000" cy="6000"/>
                        <a:chOff x="100" y="200"/>
                        <a:chExt cx="2000" cy="3000"/>
                      </a:xfrm>
                    </p:grpSpPr>
                    <p:sp>
                      <p:nvSpPr/>
                      <p:spPr>
                        <a:xfrm>
                          <a:off x="600" y="1700"/>
                          <a:ext cx="1000" cy="1000"/>
                        </a:xfrm>
                      </p:spPr>
                      <p:txBody>
                        <a:bodyPr/>
                        <a:lstStyle/>
                        <a:p><a:r><a:t>Grouped</a:t></a:r></a:p>
                      </p:txBody>
                    </p:sp>
                  </p:grpSp>
                </p:spTree>
              </p:cSld>
            </p:sld>
        "#;

        let elements = parse_slide_xml(xml.as_bytes()).expect("Failed to parse slide XML");

        let text = elements.iter().find(|element| matches!(element, SlideElement::Text(_, _)))
            .expect("Expected text element");

        match text {
            SlideElement::Text(_, position) => assert_eq!(*position, ElementPosition { x: 2000, y: 5000 }),
            element => panic!("Expected text element, got {:?}", element),
        }
    }

    #[test]
    fn test_parse_slide_xml_uses_inherited_placeholder_position() {
        let xml = r#"
            <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                   xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:cSld>
                <p:spTree>
                  <p:nvGrpSpPr/>
                  <p:grpSpPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="0" cy="0"/>
                      <a:chOff x="0" y="0"/>
                      <a:chExt cx="0" cy="0"/>
                    </a:xfrm>
                  </p:grpSpPr>
                  <p:sp>
                    <p:nvSpPr>
                      <p:cNvPr id="2" name="Title"/>
                      <p:cNvSpPr/>
                      <p:nvPr><p:ph type="title"/></p:nvPr>
                    </p:nvSpPr>
                    <p:spPr/>
                    <p:txBody>
                      <a:bodyPr/>
                      <a:lstStyle/>
                      <a:p><a:r><a:t>Inherited title</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp>
                </p:spTree>
              </p:cSld>
            </p:sld>
        "#;

        let mut positions = HashMap::new();
        positions.insert(
            PlaceholderKey {
                kind: Some("title".to_string()),
                idx: None,
            },
            ElementPosition { x: 695326, y: 333375 },
        );

        let inherited = InheritedPositions { positions };
        let elements = parse_slide_xml_with_inherited_positions(xml.as_bytes(), &inherited)
            .expect("Failed to parse slide XML");

        let text = elements.iter().find(|element| matches!(element, SlideElement::Text(_, _)))
            .expect("Expected text element");

        match text {
            SlideElement::Text(_, position) => assert_eq!(*position, ElementPosition { x: 695326, y: 333375 }),
            element => panic!("Expected text element, got {:?}", element),
        }
    }

    #[test]
    fn test_extract_inherited_positions_falls_back_to_master_placeholder() {
        let master_xml = r#"
            <p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:cSld>
                <p:spTree>
                  <p:nvGrpSpPr/>
                  <p:grpSpPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="0" cy="0"/>
                      <a:chOff x="0" y="0"/>
                      <a:chExt cx="0" cy="0"/>
                    </a:xfrm>
                  </p:grpSpPr>
                  <p:sp>
                    <p:nvSpPr>
                      <p:cNvPr id="2" name="Footer"/>
                      <p:cNvSpPr/>
                      <p:nvPr><p:ph type="ftr" idx="11"/></p:nvPr>
                    </p:nvSpPr>
                    <p:spPr>
                      <a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm>
                    </p:spPr>
                  </p:sp>
                </p:spTree>
              </p:cSld>
            </p:sldMaster>
        "#;

        let layout_xml = r#"
            <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:cSld>
                <p:spTree>
                  <p:nvGrpSpPr/>
                  <p:grpSpPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="0" cy="0"/>
                      <a:chOff x="0" y="0"/>
                      <a:chExt cx="0" cy="0"/>
                    </a:xfrm>
                  </p:grpSpPr>
                  <p:sp>
                    <p:nvSpPr>
                      <p:cNvPr id="4" name="Footer"/>
                      <p:cNvSpPr/>
                      <p:nvPr><p:ph type="ftr" idx="11"/></p:nvPr>
                    </p:nvSpPr>
                    <p:spPr/>
                  </p:sp>
                </p:spTree>
              </p:cSld>
            </p:sldLayout>
        "#;

        let master_positions = extract_inherited_positions(master_xml.as_bytes(), &InheritedPositions::default())
            .expect("Failed to parse master positions");
        let layout_positions = extract_inherited_positions(layout_xml.as_bytes(), &master_positions)
            .expect("Failed to parse layout positions");

        let position = layout_positions.resolve(&PlaceholderKey {
            kind: Some("ftr".to_string()),
            idx: Some("11".to_string()),
        }).expect("Expected footer placeholder position");

        assert_eq!(position, ElementPosition { x: 100, y: 200 });
    }

    #[test]
    fn test_extract_inherited_positions_prefers_layout_placeholder_with_same_idx() {
        let master_xml = r#"
            <p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:cSld>
                <p:spTree>
                  <p:nvGrpSpPr/>
                  <p:grpSpPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="0" cy="0"/>
                      <a:chOff x="0" y="0"/>
                      <a:chExt cx="0" cy="0"/>
                    </a:xfrm>
                  </p:grpSpPr>
                  <p:sp>
                    <p:nvSpPr>
                      <p:cNvPr id="2" name="Footer"/>
                      <p:cNvSpPr/>
                      <p:nvPr><p:ph type="ftr" idx="11"/></p:nvPr>
                    </p:nvSpPr>
                    <p:spPr>
                      <a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm>
                    </p:spPr>
                  </p:sp>
                </p:spTree>
              </p:cSld>
            </p:sldMaster>
        "#;

        let layout_xml = r#"
            <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:cSld>
                <p:spTree>
                  <p:nvGrpSpPr/>
                  <p:grpSpPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="0" cy="0"/>
                      <a:chOff x="0" y="0"/>
                      <a:chExt cx="0" cy="0"/>
                    </a:xfrm>
                  </p:grpSpPr>
                  <p:sp>
                    <p:nvSpPr>
                      <p:cNvPr id="4" name="Footer"/>
                      <p:cNvSpPr/>
                      <p:nvPr><p:ph idx="11"/></p:nvPr>
                    </p:nvSpPr>
                    <p:spPr>
                      <a:xfrm><a:off x="900" y="800"/><a:ext cx="300" cy="400"/></a:xfrm>
                    </p:spPr>
                  </p:sp>
                </p:spTree>
              </p:cSld>
            </p:sldLayout>
        "#;

        let master_positions = extract_inherited_positions(master_xml.as_bytes(), &InheritedPositions::default())
            .expect("Failed to parse master positions");
        let layout_positions = extract_inherited_positions(layout_xml.as_bytes(), &master_positions)
            .expect("Failed to parse layout positions");

        let position = layout_positions.resolve(&PlaceholderKey {
            kind: Some("ftr".to_string()),
            idx: Some("11".to_string()),
        }).expect("Expected footer placeholder position");

        assert_eq!(position, ElementPosition { x: 900, y: 800 });
    }
