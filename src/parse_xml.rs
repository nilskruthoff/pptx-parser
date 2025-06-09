use crate::constants::{A_NAMESPACE, P_NAMESPACE, RELS_NAMESPACE};
use crate::types::{SlideElement, TableCell, TableElement, TableRow, TextElement};
use crate::SlideElement::Unknown;
use crate::{Error, Formatting, ImageReference, ListElement, ListItem, Result, Run};
use roxmltree::{Document, Node};


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
        let tag_name = child_node.tag_name().name();
        let namespace = child_node.tag_name().namespace().unwrap_or("");
        if namespace == P_NAMESPACE {
            match tag_name {
                "sp" => {
                    let slide = parse_sp(&child_node)?;
                    elements.push(slide);
                },
                "graphicFrame" => {
                    if let Some(element) = parse_graphic_frame(&child_node)? {
                        elements.push(element);
                    }
                },
                "pic" => {
                    let image_element = parse_pic(&child_node)?;
                    elements.push(image_element);
                },
                _ => {
                    elements.push(Unknown)
                }
            }
        }
    }

    Ok(elements)
}

/// Parses the text body node (`<p:txBody>`) ito search for shape nodes (`<a:sp>`) and
/// evaluates if a shape is formatted list or a common text
fn parse_sp(sp_node: &Node) -> Result<SlideElement> {
    let tx_body_node = sp_node.children().find(|n| {
        n.is_element()
            && n.tag_name().name() == "txBody"
            && n.tag_name().namespace() == Some(P_NAMESPACE)
    }).ok_or(Error::Unknown)?;

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
        parse_list(&tx_body_node)
    } else {
        parse_text(&tx_body_node)
    }
}

/// Parses the text body node (`<p:txBody>`) for all paragraph nodes (`<a:p>`) containing text runs
/// # Returns
/// Returns a `Result` containing either:
/// - `SlideElement::Text`: A text element containing all text runs
/// - `Error`: Error information encapsulated in [`crate::Error`] if parsing fails at XML parsing level.
fn parse_text(tx_body_node: &Node) -> Result<SlideElement> {
    let mut runs = Vec::new();

    for p_node in tx_body_node.children().filter(|n| {
        n.is_element()
            && n.tag_name().name() == "p"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
    }) {
        let mut paragraph_runs = parse_paragraph(&p_node, true)?;
        runs.append(&mut paragraph_runs);
    }

    Ok(SlideElement::Text(TextElement { runs }))
}

fn parse_graphic_frame(node: &Node) -> Result<Option<SlideElement>> {
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
            return Ok(Some(SlideElement::Table(table)));
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
fn parse_pic(pic_node: &Node) -> Result<SlideElement> {
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

    Ok(SlideElement::Image(image_ref))
}

/// Parses the paragraph node (`<a:p>`) that is already identified as a list from the text body node (`<p:txBody>`)
/// and extracts the _text runs_, the _level of indentation_ and weather its _ordered_ or _unordered_
///
/// # Returns
/// - `SlideElement::List`: A complete lists with all children of type `ListElement`
/// - `Error`: Error information encapsulated in [`crate::Error`] if parsing fails at XML parsing level.
fn parse_list(tx_body_node: &Node) -> Result<SlideElement> {
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

    Ok(SlideElement::List(ListElement { items }))
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