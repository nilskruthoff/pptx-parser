use crate::constants::{A_NAMESPACE, P_NAMESPACE, RELS_NAMESPACE};
use crate::types::{SlideElement, TableCell, TableElement, TableRow, TextElement};
use crate::{ElementPosition, Error, Formatting, ImageReference, ListElement, ListItem, Result, Run};
use roxmltree::{Document, Node};
use std::collections::HashMap;

enum ParsedContent {
    Text(TextElement),
    List(ListElement),
}

/// Represents the accumulated coordinate transformation from a local element space
/// into slide-level coordinates.
///
/// Grouped shapes in PPTX can define their own origin (`chOff`) and size (`chExt`)
/// which must be mapped into the enclosing group frame (`off`/`ext`). This helper
/// stores the current scaling and translation so child elements can be converted
/// into their effective slide position.
#[derive(Debug, Clone, Copy)]
struct CoordinateTransform {
    scale_x: f64,
    scale_y: f64,
    translate_x: f64,
    translate_y: f64,
}

impl CoordinateTransform {
    fn identity() -> Self {
        Self {
            scale_x: 1.0,
            scale_y: 1.0,
            translate_x: 0.0,
            translate_y: 0.0,
        }
    }

    fn apply(self, position: ElementPosition) -> ElementPosition {
        ElementPosition {
            x: (position.x as f64 * self.scale_x + self.translate_x).round() as i64,
            y: (position.y as f64 * self.scale_y + self.translate_y).round() as i64,
        }
    }

    fn then(self, next: CoordinateTransform) -> CoordinateTransform {
        CoordinateTransform {
            scale_x: self.scale_x * next.scale_x,
            scale_y: self.scale_y * next.scale_y,
            translate_x: self.scale_x * next.translate_x + self.translate_x,
            translate_y: self.scale_y * next.translate_y + self.translate_y,
        }
    }
}

/// Identifies a placeholder shape across slide, layout, and master documents.
///
/// PowerPoint placeholders are primarily matched by `type` and `idx`. Some
/// layouts omit the type attribute, so resolution falls back from exact match
/// to looser combinations when needed.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct PlaceholderKey {
    kind: Option<String>,
    idx: Option<String>,
}

/// Holds inherited placeholder positions resolved from slide layouts and masters.
#[derive(Debug, Clone, Default)]
pub struct InheritedPositions {
    positions: HashMap<PlaceholderKey, ElementPosition>,
}

impl InheritedPositions {
    fn resolve(&self, key: &PlaceholderKey) -> Option<ElementPosition> {
        self.positions
            .get(key)
            .copied()
            .or_else(|| {
                key.idx.as_ref().and_then(|idx| {
                    self.positions
                        .iter()
                        .find(|(candidate, _)| candidate.idx.as_deref() == Some(idx.as_str()))
                        .map(|(_, position)| *position)
                })
            })
            .or_else(|| {
                key.kind.as_ref().and_then(|kind| {
                    self.positions
                        .iter()
                        .find(|(candidate, _)| candidate.kind.as_deref() == Some(kind.as_str()))
                        .map(|(_, position)| *position)
                })
            })
    }
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
    parse_slide_xml_with_inherited_positions(xml_data, &InheritedPositions::default())
}

/// Parses raw slide XML and resolves missing placeholder positions from inherited sources.
pub fn parse_slide_xml_with_inherited_positions(
    xml_data: &[u8],
    inherited_positions: &InheritedPositions,
) -> Result<Vec<SlideElement>> {
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
        elements.extend(parse_group(&child_node, CoordinateTransform::identity(), inherited_positions)?);
    }

    Ok(elements)
}

/// Parses a slide node and all nested child nodes recursively.
///
/// The `transform` argument carries the accumulated group transformation so elements
/// inside nested `<p:grpSp>` containers can be positioned in slide coordinates.
fn parse_group(node: &Node, transform: CoordinateTransform, inherited_positions: &InheritedPositions) -> Result<Vec<SlideElement>> {
    let mut elements = Vec::new();

    let tag_name = node.tag_name().name();
    let namespace = node.tag_name().namespace().unwrap_or("");

    if namespace != P_NAMESPACE {
        return Ok(elements);
    }

    let position = extract_position(node, transform, inherited_positions);

    match tag_name {
        "sp" => {
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
            let child_transform = transform.then(extract_group_transform(node));
            for child in node.children().filter(|n| n.is_element()) {
                elements.extend(parse_group(&child, child_transform, inherited_positions)?);
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

/// Extracts the effective slide position for a node.
///
/// The node's local coordinates are read from its transform XML and then converted
/// through the accumulated parent-group transform.
fn extract_position(node: &Node, transform: CoordinateTransform, inherited_positions: &InheritedPositions) -> ElementPosition {
    extract_raw_position(node)
        .map(|position| transform.apply(position))
        .or_else(|| extract_placeholder_key(node).and_then(|key| inherited_positions.resolve(&key)))
        .unwrap_or_default()
}

/// Reads a node's local position directly from its XML transform element.
///
/// Different PPTX element types store their transform in different places:
/// shapes and pictures use `<a:xfrm>` inside `<p:spPr>`, while tables inside
/// `<p:graphicFrame>` use `<p:xfrm>`.
fn extract_raw_position(node: &Node) -> Option<ElementPosition> {
    let xfrm = match node.tag_name().name() {
        "sp" | "pic" => node
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(P_NAMESPACE))
            .and_then(|sp_pr| {
                sp_pr.children().find(|n| {
                    n.is_element() && n.tag_name().name() == "xfrm" && n.tag_name().namespace() == Some(A_NAMESPACE)
                })
            }),
        "graphicFrame" => node.children().find(|n| {
            n.is_element() && n.tag_name().name() == "xfrm" && n.tag_name().namespace() == Some(P_NAMESPACE)
        }),
        "grpSp" => node
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "grpSpPr" && n.tag_name().namespace() == Some(P_NAMESPACE))
            .and_then(|grp_sp_pr| {
                grp_sp_pr.children().find(|n| {
                    n.is_element() && n.tag_name().name() == "xfrm" && n.tag_name().namespace() == Some(A_NAMESPACE)
                })
            }),
        _ => None,
    }?;

    Some(ElementPosition {
        x: extract_child_attr_i64(&xfrm, A_NAMESPACE, "off", "x")?,
        y: extract_child_attr_i64(&xfrm, A_NAMESPACE, "off", "y")?,
    })
}

/// Builds the transformation that maps a group's local child coordinates
/// into the coordinate system of its parent.
///
/// PowerPoint groups define both a child origin/extent (`chOff`/`chExt`) and a
/// rendered origin/extent (`off`/`ext`). The resulting transform combines scaling
/// and translation so all nested elements can be reported in slide space.
fn extract_group_transform(node: &Node) -> CoordinateTransform {
    let Some(xfrm) = node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "grpSpPr" && n.tag_name().namespace() == Some(P_NAMESPACE))
        .and_then(|grp_sp_pr| {
            grp_sp_pr.children().find(|n| {
                n.is_element() && n.tag_name().name() == "xfrm" && n.tag_name().namespace() == Some(A_NAMESPACE)
            })
        }) else {
        return CoordinateTransform::identity();
    };

    let off_x = extract_child_attr_i64(&xfrm, A_NAMESPACE, "off", "x").unwrap_or(0) as f64;
    let off_y = extract_child_attr_i64(&xfrm, A_NAMESPACE, "off", "y").unwrap_or(0) as f64;
    let ch_off_x = extract_child_attr_i64(&xfrm, A_NAMESPACE, "chOff", "x").unwrap_or(0) as f64;
    let ch_off_y = extract_child_attr_i64(&xfrm, A_NAMESPACE, "chOff", "y").unwrap_or(0) as f64;
    let ext_x = extract_child_attr_i64(&xfrm, A_NAMESPACE, "ext", "cx").unwrap_or(0) as f64;
    let ext_y = extract_child_attr_i64(&xfrm, A_NAMESPACE, "ext", "cy").unwrap_or(0) as f64;
    let ch_ext_x = extract_child_attr_i64(&xfrm, A_NAMESPACE, "chExt", "cx").unwrap_or(0) as f64;
    let ch_ext_y = extract_child_attr_i64(&xfrm, A_NAMESPACE, "chExt", "cy").unwrap_or(0) as f64;

    let scale_x = if ch_ext_x == 0.0 { 1.0 } else { ext_x / ch_ext_x };
    let scale_y = if ch_ext_y == 0.0 { 1.0 } else { ext_y / ch_ext_y };

    CoordinateTransform {
        scale_x,
        scale_y,
        translate_x: off_x - ch_off_x * scale_x,
        translate_y: off_y - ch_off_y * scale_y,
    }
}

/// Reads an integer attribute from a direct child node with the given namespace and tag name.
fn extract_child_attr_i64(node: &Node, namespace: &str, child_name: &str, attr_name: &str) -> Option<i64> {
    node.children()
        .find(|n| n.is_element() && n.tag_name().name() == child_name && n.tag_name().namespace() == Some(namespace))
        .and_then(|child| child.attribute(attr_name))
        .and_then(|value| value.parse::<i64>().ok())
}

fn extract_placeholder_key(node: &Node) -> Option<PlaceholderKey> {
    let placeholder = node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "ph" && n.tag_name().namespace() == Some(P_NAMESPACE))?;

    Some(PlaceholderKey {
        kind: placeholder.attribute("type").map(|value| value.to_string()),
        idx: placeholder.attribute("idx").map(|value| value.to_string()),
    })
}

/// Extracts placeholder positions from a layout or master XML document.
///
/// Placeholder nodes that do not define their own transform inherit their
/// position from the provided `inherited_positions`.
pub fn extract_inherited_positions(
    xml_data: &[u8],
    inherited_positions: &InheritedPositions,
) -> Result<InheritedPositions> {
    let xml_str = std::str::from_utf8(xml_data).map_err(|_| Error::Unknown)?;
    let doc = Document::parse(xml_str)?;
    let root = doc.root_element();

    let c_sld = root
        .descendants()
        .find(|n| n.tag_name().name() == "cSld" && n.tag_name().namespace() == root.tag_name().namespace())
        .ok_or(Error::Unknown)?;

    let sp_tree = c_sld
        .children()
        .find(|n| n.tag_name().name() == "spTree" && n.tag_name().namespace() == root.tag_name().namespace())
        .ok_or(Error::Unknown)?;

    let mut positions = inherited_positions.positions.clone();
    collect_placeholder_positions(&sp_tree, CoordinateTransform::identity(), inherited_positions, &mut positions);

    Ok(InheritedPositions { positions })
}

fn collect_placeholder_positions(
    node: &Node,
    transform: CoordinateTransform,
    inherited_positions: &InheritedPositions,
    positions: &mut HashMap<PlaceholderKey, ElementPosition>,
) {
    for child in node.children().filter(|n| n.is_element()) {
        if child.tag_name().namespace() != Some(P_NAMESPACE) {
            continue;
        }

        if child.tag_name().name() == "grpSp" {
            let child_transform = transform.then(extract_group_transform(&child));
            collect_placeholder_positions(&child, child_transform, inherited_positions, positions);
            continue;
        }

        if let Some(key) = extract_placeholder_key(&child) {
            let position = extract_raw_position(&child)
                .map(|local| transform.apply(local))
                .or_else(|| inherited_positions.resolve(&key));

            if let Some(position) = position {
                insert_placeholder_position(positions, key, position);
            }
        }
    }
}

fn insert_placeholder_position(
    positions: &mut HashMap<PlaceholderKey, ElementPosition>,
    key: PlaceholderKey,
    position: ElementPosition,
) {
    if let Some(idx) = key.idx.as_deref() {
        positions.retain(|candidate, _| candidate.idx.as_deref() != Some(idx));
    }

    positions.insert(key, position);
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
}
