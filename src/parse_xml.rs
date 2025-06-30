use crate::constants::{A_NAMESPACE, P_NAMESPACE, RELS_NAMESPACE};
use crate::types::{SlideElement, TableCell, TableElement, TableRow, TextElement};
use crate::{ElementPosition, Error, Formatting, ImageReference, ListElement, ListItem, Result, Run};
use roxmltree::{Document, Node};

enum ParsedContent {
    Text(TextElement),
    List(ListElement),
}

/// Parses raw XML slide data from a PowerPoint (pptx) file and extracts all slide elements.
///
/// This function processes a single PowerPoint slide's XML data to identify and parse its
/// contained elements into structured variants such as text blocks, tables, images, and lists.
/// Unrecognized or malformed elements will result in inclusion of a [`SlideElement::Unknown`] variant.
///
/// # Arguments
///
/// - `xml_data`: Byte slice containing raw XML data of a PowerPoint slide.
///
/// # Returns
///
/// Returns a `Result` containing either:
/// - `Vec<SlideElement>`: Vector of successfully parsed slide elements.
/// - `Error`: Error information encapsulated in [`crate::Error`] if parsing fails at XML parsing level.
///
/// # Errors
///
/// Parsing may fail and return [`Error`] if:
/// - The provided XML data isn't valid UTF-8.
/// - The XML structure is malformed or missing essential schema elements (`<p:cSld>` or `<p:spTree>` tags).
///
/// # Notes
///
/// - The function strictly follows Microsoft's Open XML slide schema.
/// - For best results, ensure input XML data is extracted directly from PPTX files or equivalent sources.
pub fn parse_slide_xml(xml_data: &[u8]) -> Result<Vec<SlideElement>> {
    let xml_str = std::str::from_utf8(xml_data).map_err(|_| Error::Unknown)?;
    let doc = Document::parse(xml_str)?;
    let root = doc.root_element();
    let ns = root.tag_name().namespace();

    let c_sld = root
        .descendants()
        .find(|n| n.tag_name().name() == "cSld" && n.tag_name().namespace() == ns)
        .ok_or(format!("No <p:cSld> tag was found for: {:?}", ns)).map_err(|_| Error::Unknown)?;

    let sp_tree = c_sld
        .children()
        .find(|n| n.tag_name().name() == "spTree" && n.tag_name().namespace() == ns)
        .ok_or(format!("No <p:spTree> tag was found for: {:?}", ns)).map_err(|_| Error::Unknown)?;

    let mut elements = Vec::new();
    for child_node in sp_tree.children().filter(|n| n.is_element()) {
        elements.extend(parse_group(&child_node)?);
    }

    Ok(elements)
}

/// Parst eine gesamte Gruppe und alle untergeordneten Kind-Elemente rekursiv
fn parse_group(node: &Node) -> Result<Vec<SlideElement>> {
    let mut elements = Vec::new();

    let tag_name = node.tag_name().name();
    let namespace = node.tag_name().namespace().unwrap_or("");

    if namespace != P_NAMESPACE {
        return Ok(elements);
    }

    let position = extract_position(node);

    match tag_name {
        "sp" => {
            let position = extract_position(node);
            match parse_sp(node)? {
                ParsedContent::Text(text) => elements.push(SlideElement::Text(text, position)),
                ParsedContent::List(list) => elements.push(SlideElement::List(list, position)),
            }
        },
        "graphicFrame" => {
            if let Some(graphic_element) = parse_graphic_frame(&node)? {
                elements.push(SlideElement::Table(graphic_element, position));
            }
        },
        "pic" => {
            let image_reference = parse_pic(&node)?;
            elements.push(SlideElement::Image(image_reference, position));
        },
        "grpSp" => {
            for child in node.children().filter(|n| n.is_element()) {
                elements.extend(parse_group(&child)?);
            }
        },
        _ => elements.push(SlideElement::Unknown),
    }

    Ok(elements)
}

/// Parses the text body node (`<p:txBody>`) ito search for shape nodes (`<a:sp>`) and
/// evaluates if a shape is a formatted list or a common text
fn parse_sp(sp_node: &Node) -> Result<ParsedContent> {
    let tx_body_node = sp_node.children()
        .find(|n| n.tag_name().name() == "txBody" && n.tag_name().namespace() == Some(P_NAMESPACE))
        .ok_or(Error::Unknown)?;

    let is_list = tx_body_node.descendants().any(|n| {
        n.is_element()
            && n.tag_name().name() == "pPr"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
            && (
            n.attribute("lvl").is_some() ||
                n.children().any(|child| {
                    child.is_element() && (
                        child.tag_name().name() == "buAutoNum" ||
                            child.tag_name().name() == "buChar"
                    )
                })
        )
    });

    if is_list {
        Ok(ParsedContent::List(parse_list(&tx_body_node)?))
    } else {
        Ok(ParsedContent::Text(parse_text(&tx_body_node)?))
    }
}

/// Parses the text body node (`<p:txBody>`) for all paragraph nodes (`<a:p>`) containing text runs
/// # Returns
/// Returns a `Result` containing either:
/// - `SlideElement::Text`: A text element containing all text runs
/// - `Error`: Error information encapsulated in [`crate::Error`] if parsing fails at XML parsing level.
fn parse_text(tx_body_node: &Node) -> Result<TextElement> {
    let mut runs = Vec::new();

    for p_node in tx_body_node.children().filter(|n| {
        n.is_element()
            && n.tag_name().name() == "p"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
    }) {
        let mut paragraph_runs = parse_paragraph(&p_node, true)?;
        runs.append(&mut paragraph_runs);
    }

    Ok(TextElement { runs })
}

fn parse_graphic_frame(node: &Node) -> Result<Option<TableElement>> {
    let graphic_data_node = node
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "graphicData"
                && n.tag_name().namespace() == Some(A_NAMESPACE)
                && n.attribute("uri") == Some("http://schemas.openxmlformats.org/drawingml/2006/table")
        });

    if let Some(graphic_data) = graphic_data_node {
        if let Some(tbl_node) = graphic_data
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "tbl" && n.tag_name().namespace() == Some(A_NAMESPACE))
        {
            let table = parse_table(&tbl_node)?;
            return Ok(Some(table));
        }
    }

    Ok(None)
}

/// Parses a table node (`<a:tbl>`) and extracts all
/// table rows ('<a:tr>') elements to construct a `TableElement`.
fn parse_table(tbl_node: &Node) -> Result<TableElement> {
    let mut rows = Vec::new();

    for tr_node in tbl_node.children().filter(|n| {
        n.is_element()
            && n.tag_name().name() == "tr"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
    }) {
        let row = parse_table_row(&tr_node)?;
        rows.push(row);
    }

    Ok(TableElement { rows })
}

/// Parses a table row node (`'<a:tr>'`) and extracts all
/// table cells ('<a:tc>') elements to construct a full `TableRow`.
fn parse_table_row(tr_node: &Node) -> Result<TableRow> {
    let mut cells = Vec::new();

    for tc_node in tr_node.children().filter(|n| {
        n.is_element()
            && n.tag_name().name() == "tc"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
    }) {
        let cell = parse_table_cell(&tc_node)?;
        cells.push(cell);
    }

    Ok(TableRow { cells })
}

/// Parses a table cell node (`'<a:tc>'`) and extracts all
/// paragraph nodes ('<a:p>') to construct a `TableCell`.
fn parse_table_cell(tc_node: &Node) -> Result<TableCell> {
    let mut runs = Vec::new();

    if let Some(tx_body_node) = tc_node.children().find(|n| {
        n.is_element()
            && n.tag_name().name() == "txBody"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
    }) {
        for p_node in tx_body_node.children().filter(|n| {
            n.is_element()
                && n.tag_name().name() == "p"
                && n.tag_name().namespace() == Some(A_NAMESPACE)
        }) {
            let mut paragraph_runs = parse_paragraph(&p_node, false)?;
            runs.append(&mut paragraph_runs);
        }
    }

    Ok(TableCell { runs })
}

/// Parses an image node (`<a:pic>`) to extract an image reference.
///
/// This function locates and processes the `<a:blip>` element inside a given
/// image node, extracting the necessary attributes to build an `ImageReference` object
/// that is necessary to link the image and extracts the relative image path from the media dir
///
/// # Returns
///
/// Returns a `Result` with:
/// - `SlideElement::Image`: A `SlideElement` containing the image's reference `ID` to link it if successfully parsed.
/// - `Error::ImageNotFound`: If the `<blip>` element or necessary attributes are missing.
fn parse_pic(pic_node: &Node) -> Result<ImageReference> {
    let blip_node = pic_node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "blip" && n.tag_name().namespace() == Some(A_NAMESPACE))
        .ok_or(Error::ImageNotFound)?;

    let embed_attr = blip_node.attribute((RELS_NAMESPACE, "embed"))
        .or_else(|| blip_node.attribute("r:embed"))
        .ok_or(Error::ImageNotFound)?;

    let image_ref = ImageReference {
        id: embed_attr.to_string(),
        target: String::new(),
    };

    Ok(image_ref)
}

/// Parses the paragraph node (`<a:p>`) that is already identified as a list from the text body node (`<p:txBody>`)
/// and extracts the _text runs_, the _level of indentation_ and weather its _ordered_ or _unordered_
///
/// # Returns
/// - `SlideElement::List`: A complete lists with all children of type `ListElement`
/// - `Error`: Error information encapsulated in [`crate::Error`] if parsing fails at XML parsing level.
fn parse_list(tx_body_node: &Node) -> Result<ListElement> {
    let mut items = Vec::new();

    for p_node in tx_body_node.children().filter(|n| {
        n.is_element()
            && n.tag_name().name() == "p"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
    }) {
        let (level, is_ordered) = parse_list_properties(&p_node)?;

        let runs = parse_paragraph(&p_node, true)?;

        items.push(ListItem { level, is_ordered, runs });
    }

    Ok(ListElement { items })
}

/// Extracts list properties from a paragraph node (``<a:p>`).
///
/// This function analyzes a paragraph node to determine its list level and
/// whether it's an ordered or unordered list in a PowerPoint slide's XML structure.
///
/// # Returns
///
/// Returns a `Result` containing:
/// - `Ok((level, is_ordered))`: A tuple where `level` (u32) indicates the list depth level and `is_ordered` (bool) indicates if the list is ordered or unordered.
/// - `Err(Error)`: When parsing fails due to structural inconsistencies in the XML node.
fn parse_list_properties(p_node: &Node) -> Result<(u32, bool)> {
    let mut level = 0;
    let mut is_ordered = false;

    if let Some(p_pr_node) = p_node.children().find(|n| {
        n.is_element()
            && n.tag_name().name() == "pPr"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
    }) {
        if let Some(lvl_attr) = p_pr_node.attribute("lvl") {
            level = lvl_attr.parse::<u32>().unwrap_or(0);
        }

        is_ordered = p_pr_node.children().any(|n| {
            n.is_element() && n.tag_name().namespace() == Some(A_NAMESPACE) && n.tag_name().name() == "buAutoNum"
        });

        if !is_ordered {
            is_ordered = p_pr_node.children().any(|n| {
                n.is_element() && n.tag_name().namespace() == Some(A_NAMESPACE) && n.tag_name().name() == "buChar"
            });
        }
    }

    Ok((level, is_ordered))
}

/// Parses a single text paragraph node (`<a:p>`) into multiple text runs.
///
/// # Notes
/// Searches for the last run and adds a newline character
fn parse_paragraph(p_node: &Node, add_new_line: bool) -> Result<Vec<Run>> {
    let run_nodes: Vec<_> = p_node.children().filter(|n| {
        n.is_element()
            && n.tag_name().name() == "r"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
    }).collect();

    let count = run_nodes.len();
    let mut runs: Vec<Run> = Vec::new();

    for (idx, r_node) in run_nodes.iter().enumerate() {
        let mut run = parse_run(r_node)?;
        
        if add_new_line && idx == count - 1 {
            run.text.push('\n');
        }

        runs.push(run);
    }
    Ok(runs)
}

/// Parses a single run properties node (`<a:rPr>`) and extracting the text content from the text node (`<a:t>`)
/// as well as the format including _bold_, _italic_, _underlined_ and the _language_
fn parse_run(r_node: &Node) -> Result<Run> {
    let mut text = String::new();
    let mut formatting = Formatting::default();

    if let Some(r_pr_node) = r_node.children().find(|n| {
        n.is_element()
            && n.tag_name().name() == "rPr"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
    }) {
        if let Some(b_attr) = r_pr_node.attribute("b") {
            formatting.bold = b_attr == "1" || b_attr.eq_ignore_ascii_case("true");
        }
        if let Some(i_attr) = r_pr_node.attribute("i") {
            formatting.italic = i_attr == "1" || i_attr.eq_ignore_ascii_case("true");
        }
        if let Some(u_attr) = r_pr_node.attribute("u") {
            formatting.underlined = u_attr != "none";
        }
        if let Some(lang_attr) = r_pr_node.attribute("lang") {
            formatting.lang = lang_attr.to_string();
        }
    }

    if let Some(t_node) = r_node.children().find(|n| {
        n.is_element()
            && n.tag_name().name() == "t"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
    }) {
        if let Some(t) = t_node.text() {
            text.push_str(t);
        }
    }
    Ok(Run { text, formatting })
}

fn extract_position(node: &Node) -> ElementPosition {
    let default = ElementPosition::default();

    node.descendants()
        .find(|n| n.tag_name().namespace() == Some(A_NAMESPACE) && n.tag_name().name() == "xfrm")
        .and_then(|xfrm| {
            let x = xfrm
                .children()
                .find(|n| n.tag_name().name() == "off" && n.tag_name().namespace() == Some(A_NAMESPACE))
                .and_then(|off| off.attribute("x")?.parse::<i64>().ok())?;

            let y = xfrm
                .children()
                .find(|n| n.tag_name().name() == "off" && n.tag_name().namespace() == Some(A_NAMESPACE))
                .and_then(|off| off.attribute("y")?.parse::<i64>().ok())?;

            Some(ElementPosition { x, y })
        })
        .unwrap_or(default)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;
    use std::path::PathBuf;

    fn load_xml(filename: &str) -> String {
        let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        path.push("tests");
        path.push("test_data");
        path.push("xml");
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
    fn test_parse_text() {
        let xml_data = load_xml("tx_body.xml");
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
            _ => {}
        }
    }

    #[test]
    fn test_parse_run_with_format() {
        let xml_data = load_xml("run_styles.xml");
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
        let xml_data = load_xml("run_no_format.xml");
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
        let xml_data = load_xml("run_empty.xml");
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
        let xml_data = load_xml("paragraph_single.xml");
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
        let xml_data = load_xml("paragraph_multiple.xml");
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
        let xml_data = load_xml("paragraph_empty.xml");
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
        let xml_data = load_xml("simple_list.xml");
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
        let xml_data = load_xml("multilevel_list.xml");
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
        let xml_data = load_xml("simple_list.xml");
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
            Ok(_) => panic!("Expected a List element but got something else"),
            Err(_) => panic!("Failed to parse simple list")
        }
    }

    #[test]
    fn test_parse_multilevel_list() {
        // Test for parsing a multilevel list
        let xml_data = load_xml("multilevel_list.xml");
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
            Ok(_) => panic!("Expected a List element but got something else"),
            Err(_) => panic!("Failed to parse multilevel list")
        }
    }

    /// Test for a simple table for a cell with a single paragraph
    #[test]
    fn test_parse_table_cell_simple() {
        let xml_data = load_xml("simple_table.xml");
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
        let xml_data = load_xml("complex_table.xml");
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
        let xml_data = load_xml("empty_table.xml");
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
        let xml_data = load_xml("simple_table.xml");
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
        let xml_data = load_xml("complex_table.xml");
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
        let xml_data = load_xml("simple_table.xml");
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
        let xml_data = load_xml("complex_table.xml");
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
        let xml_data = load_xml("empty_table.xml");
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
        let xml_data = load_xml("simple_table.xml");
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
            Ok(_) => panic!("Found a different slide element, expected a table"),
            Err(_) => panic!("Failed to parse graphic frame with table")
        }
    }

    #[test]
    fn test_parse_graphic_frame_without_table() {
        // Test for a graphic frame that doesn't contain a table
        let xml_data = load_xml("non_table_graphic.xml");
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
        let xml_data = load_xml("pic_with_image.xml");
        let doc = Document::parse(&*xml_data).expect("Failed to parse XML");
        let pic_node = doc.root_element();

        match parse_pic(&pic_node) {
            Ok(image_ref) => {
                assert_eq!(image_ref.id, "rId2", "Image reference ID should be 'rId2'");
                assert_eq!(image_ref.target, "", "Image target should be empty initially");
            },
            Ok(_) => panic!("Expected an Image element but got something else"),
            Err(e) => panic!("Failed to parse picture: {:?}", e)
        }
    }

    #[test]
    fn test_parse_pic_without_embed() {
        // Test for parsing a picture without an embed attribute
        let xml_data = load_xml("pic_without_embed.xml");
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
        let xml_data = load_xml("pic_without_blip.xml");
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
}