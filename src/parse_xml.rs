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
    parse_slide_xml_with_hyperlinks(xml_data, &InheritedPositions::default(), &HashMap::new())
}

/// Parses the user-authored text from a PPTX notes slide.
///
/// Notes slides also contain placeholders for dates, footers, and slide numbers.
/// Only the `body` placeholder is speaker-note content and is therefore returned.
#[cfg(test)]
pub(crate) fn parse_speaker_notes_xml(xml_data: &[u8]) -> Result<Vec<TextElement>> {
    parse_speaker_notes_xml_with_hyperlinks(xml_data, &HashMap::new())
}

pub(crate) fn parse_speaker_notes_xml_with_hyperlinks(xml_data: &[u8], hyperlinks: &HashMap<String, String>) -> Result<Vec<TextElement>> {
    let xml_str = std::str::from_utf8(xml_data).map_err(|_| Error::Unknown)?;
    let doc = Document::parse(xml_str)?;
    let root = doc.root_element();
    let ns = root.tag_name().namespace();

    let c_sld = root
        .descendants()
        .find(|node| node.tag_name().name() == "cSld" && node.tag_name().namespace() == ns)
        .ok_or(Error::Unknown)?;
    let sp_tree = c_sld
        .children()
        .find(|node| node.tag_name().name() == "spTree" && node.tag_name().namespace() == ns)
        .ok_or(Error::Unknown)?;

    let mut notes = Vec::new();
    for shape in sp_tree.children().filter(|node| {
        node.is_element()
            && node.tag_name().name() == "sp"
            && node.tag_name().namespace() == Some(P_NAMESPACE)
            && is_notes_body_placeholder(*node)
    }) {
        if let Some(text_body) = shape.children().find(|node| {
            node.is_element()
                && node.tag_name().name() == "txBody"
                && node.tag_name().namespace() == Some(P_NAMESPACE)
        }) {
            let text = parse_text_with_hyperlinks(&text_body, hyperlinks)?;
            if !text.runs.is_empty() {
                notes.push(text);
            }
        }
    }
    Ok(notes)
}

/// Parses text content from both legacy and modern PPTX comment parts.
pub(crate) fn parse_comments_xml_with_hyperlinks(
    xml_data: &[u8],
    hyperlinks: &HashMap<String, String>,
) -> Result<Vec<TextElement>> {
    let xml_str = std::str::from_utf8(xml_data).map_err(|_| Error::Unknown)?;
    let doc = Document::parse(xml_str)?;
    let mut comments = Vec::new();

    for text_body in doc
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "txBody")
    {
        let text = parse_text_with_hyperlinks(&text_body, hyperlinks)?;
        if !text.runs.is_empty() {
            comments.push(text);
        }
    }

    if comments.is_empty() {
        for text in doc.descendants().filter(|node| {
            node.is_element()
                && node.tag_name().name() == "text"
                && node.tag_name().namespace() == Some(P_NAMESPACE)
        }) {
            let content = text.text().unwrap_or_default();
            if !content.is_empty() {
                comments.push(TextElement {
                    runs: vec![Run {
                        text: format!("{}\n", content),
                        formatting: Formatting::default(),
                        link_target: None,
                    }],
                });
            }
        }
    }

    Ok(comments)
}

fn is_notes_body_placeholder(shape: Node<'_, '_>) -> bool {
    shape.descendants().any(|node| {
        node.is_element()
            && node.tag_name().name() == "ph"
            && node.tag_name().namespace() == Some(P_NAMESPACE)
            && node.attribute("type") == Some("body")
    })
}

/// Parses raw slide XML and resolves missing placeholder positions from inherited sources.
pub fn parse_slide_xml_with_inherited_positions(
    xml_data: &[u8],
    inherited_positions: &InheritedPositions,
) -> Result<Vec<SlideElement>> {
    parse_slide_xml_with_hyperlinks(xml_data, inherited_positions, &HashMap::new())
}

/// Parses slide XML while resolving text-run hyperlink relationship IDs.
pub fn parse_slide_xml_with_hyperlinks(
    xml_data: &[u8],
    inherited_positions: &InheritedPositions,
    hyperlinks: &HashMap<String, String>,
) -> Result<Vec<SlideElement>> {
    let xml_str = std::str::from_utf8(xml_data).map_err(|_| Error::Unknown)?;
    let doc = Document::parse(xml_str)?;
    let root = doc.root_element();
    let ns = root.tag_name().namespace();

    let c_sld = root
        .descendants()
        .find(|n| n.tag_name().name() == "cSld" && n.tag_name().namespace() == ns)
        .ok_or(format!("No <p:cSld> tag was found for: {:?}", ns))
        .map_err(|_| Error::Unknown)?;

    let sp_tree = c_sld
        .children()
        .find(|n| n.tag_name().name() == "spTree" && n.tag_name().namespace() == ns)
        .ok_or(format!("No <p:spTree> tag was found for: {:?}", ns))
        .map_err(|_| Error::Unknown)?;

    let mut elements = Vec::new();
    for child_node in sp_tree.children().filter(|n| n.is_element()) {
        elements.extend(parse_group(
            &child_node,
            CoordinateTransform::identity(),
            inherited_positions,
            hyperlinks,
        )?);
    }

    Ok(elements)
}

/// Parses a slide node and all nested child nodes recursively.
///
/// The `transform` argument carries the accumulated group transformation so elements
/// inside nested `<p:grpSp>` containers can be positioned in slide coordinates.
fn parse_group(
    node: &Node,
    transform: CoordinateTransform,
    inherited_positions: &InheritedPositions,
    hyperlinks: &HashMap<String, String>,
) -> Result<Vec<SlideElement>> {
    let mut elements = Vec::new();

    let tag_name = node.tag_name().name();
    let namespace = node.tag_name().namespace().unwrap_or("");

    if namespace != P_NAMESPACE {
        return Ok(elements);
    }

    let position = extract_position(node, transform, inherited_positions);

    match tag_name {
        "sp" => match parse_sp(node, hyperlinks)? {
            ParsedContent::Text(text) => elements.push(SlideElement::Text(text, position)),
            ParsedContent::List(list) => elements.push(SlideElement::List(list, position)),
        },
        "graphicFrame" => {
            if let Some(graphic_element) = parse_graphic_frame_with_hyperlinks(&node, hyperlinks)? {
                elements.push(SlideElement::Table(graphic_element, position));
            }
        }
        "pic" => {
            let image_reference = parse_pic(&node)?;
            elements.push(SlideElement::Image(image_reference, position));
        }
        "grpSp" => {
            let child_transform = transform.then(extract_group_transform(node));
            for child in node.children().filter(|n| n.is_element()) {
                elements.extend(parse_group(
                    &child,
                    child_transform,
                    inherited_positions,
                    hyperlinks,
                )?);
            }
        }
        _ => elements.push(SlideElement::Unknown),
    }

    Ok(elements)
}

/// Parses the text body node (`<p:txBody>`) ito search for shape nodes (`<a:sp>`) and
/// evaluates if a shape is a formatted list or a common text
fn parse_sp(sp_node: &Node, hyperlinks: &HashMap<String, String>) -> Result<ParsedContent> {
    let tx_body_node = sp_node
        .children()
        .find(|n| n.tag_name().name() == "txBody" && n.tag_name().namespace() == Some(P_NAMESPACE))
        .ok_or(Error::Unknown)?;

    let is_list = tx_body_node.descendants().any(|n| {
        n.is_element()
            && n.tag_name().name() == "pPr"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
            && (n.attribute("lvl").is_some()
                || n.children().any(|child| {
                    child.is_element()
                        && (child.tag_name().name() == "buAutoNum"
                            || child.tag_name().name() == "buChar")
                }))
    });

    if is_list {
        Ok(ParsedContent::List(parse_list_with_hyperlinks(
            &tx_body_node,
            hyperlinks,
        )?))
    } else {
        Ok(ParsedContent::Text(parse_text_with_hyperlinks(
            &tx_body_node,
            hyperlinks,
        )?))
    }
}

/// Parses the text body node (`<p:txBody>`) for all paragraph nodes (`<a:p>`) containing text runs
/// # Returns
/// Returns a `Result` containing either:
/// - `SlideElement::Text`: A text element containing all text runs
/// - `Error`: Error information encapsulated in [`crate::Error`] if parsing fails at XML parsing level.
#[cfg(test)]
fn parse_text(tx_body_node: &Node) -> Result<TextElement> {
    parse_text_with_hyperlinks(tx_body_node, &HashMap::new())
}

fn parse_text_with_hyperlinks(
    tx_body_node: &Node,
    hyperlinks: &HashMap<String, String>,
) -> Result<TextElement> {
    let mut runs = Vec::new();

    for p_node in tx_body_node.children().filter(|n| {
        n.is_element()
            && n.tag_name().name() == "p"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
    }) {
        let mut paragraph_runs = parse_paragraph_with_hyperlinks(&p_node, true, hyperlinks)?;
        runs.append(&mut paragraph_runs);
    }

    Ok(TextElement { runs })
}

#[cfg(test)]
fn parse_graphic_frame(node: &Node) -> Result<Option<TableElement>> {
    parse_graphic_frame_with_hyperlinks(node, &HashMap::new())
}

fn parse_graphic_frame_with_hyperlinks(
    node: &Node,
    hyperlinks: &HashMap<String, String>,
) -> Result<Option<TableElement>> {
    let graphic_data_node = node.descendants().find(|n| {
        n.is_element()
            && n.tag_name().name() == "graphicData"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
            && n.attribute("uri") == Some("http://schemas.openxmlformats.org/drawingml/2006/table")
    });

    if let Some(graphic_data) = graphic_data_node {
        if let Some(tbl_node) = graphic_data.children().find(|n| {
            n.is_element()
                && n.tag_name().name() == "tbl"
                && n.tag_name().namespace() == Some(A_NAMESPACE)
        }) {
            let table = parse_table_with_hyperlinks(&tbl_node, hyperlinks)?;
            return Ok(Some(table));
        }
    }

    Ok(None)
}

/// Parses a table node (`<a:tbl>`) and extracts all
/// table rows ('<a:tr>') elements to construct a `TableElement`.
#[cfg(test)]
fn parse_table(tbl_node: &Node) -> Result<TableElement> {
    parse_table_with_hyperlinks(tbl_node, &HashMap::new())
}

fn parse_table_with_hyperlinks(tbl_node: &Node, hyperlinks: &HashMap<String, String>) -> Result<TableElement> {
    let mut rows = Vec::new();

    for tr_node in tbl_node.children().filter(|n| {
        n.is_element()
            && n.tag_name().name() == "tr"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
    }) {
        let row = parse_table_row_with_hyperlinks(&tr_node, hyperlinks)?;
        rows.push(row);
    }

    Ok(TableElement { rows })
}

/// Parses a table row node (`'<a:tr>'`) and extracts all
/// table cells ('<a:tc>') elements to construct a full `TableRow`.
#[cfg(test)]
fn parse_table_row(tr_node: &Node) -> Result<TableRow> {
    parse_table_row_with_hyperlinks(tr_node, &HashMap::new())
}

fn parse_table_row_with_hyperlinks(tr_node: &Node, hyperlinks: &HashMap<String, String>, ) -> Result<TableRow> {
    let mut cells = Vec::new();

    for tc_node in tr_node.children().filter(|n| {
        n.is_element()
            && n.tag_name().name() == "tc"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
    }) {
        let cell = parse_table_cell_with_hyperlinks(&tc_node, hyperlinks)?;
        cells.push(cell);
    }

    Ok(TableRow { cells })
}

/// Parses a table cell node (`'<a:tc>'`) and extracts all
/// paragraph nodes ('<a:p>') to construct a `TableCell`.
#[cfg(test)]
fn parse_table_cell(tc_node: &Node) -> Result<TableCell> {
    parse_table_cell_with_hyperlinks(tc_node, &HashMap::new())
}

fn parse_table_cell_with_hyperlinks(tc_node: &Node, hyperlinks: &HashMap<String, String>, ) -> Result<TableCell> {
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
            let mut paragraph_runs = parse_paragraph_with_hyperlinks(&p_node, false, hyperlinks)?;
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
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "blip"
                && n.tag_name().namespace() == Some(A_NAMESPACE)
        })
        .ok_or(Error::ImageNotFound)?;

    let embed_attr = blip_node
        .attribute((RELS_NAMESPACE, "embed"))
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
#[cfg(test)]
fn parse_list(tx_body_node: &Node) -> Result<ListElement> {
    parse_list_with_hyperlinks(tx_body_node, &HashMap::new())
}

fn parse_list_with_hyperlinks(tx_body_node: &Node, hyperlinks: &HashMap<String, String>, ) -> Result<ListElement> {
    let mut items = Vec::new();

    for p_node in tx_body_node.children().filter(|n| {
        n.is_element()
            && n.tag_name().name() == "p"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
    }) {
        let (level, is_ordered) = parse_list_properties(&p_node)?;

        let runs = parse_paragraph_with_hyperlinks(&p_node, true, hyperlinks)?;

        items.push(ListItem { level, is_ordered, runs, });
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
            n.is_element()
                && n.tag_name().namespace() == Some(A_NAMESPACE)
                && n.tag_name().name() == "buAutoNum"
        });

        if !is_ordered {
            is_ordered = p_pr_node.children().any(|n| {
                n.is_element()
                    && n.tag_name().namespace() == Some(A_NAMESPACE)
                    && n.tag_name().name() == "buChar"
            });
        }
    }

    Ok((level, is_ordered))
}

/// Parses a single text paragraph node (`<a:p>`) into multiple text runs.
///
/// # Notes
/// Searches for the last run and adds a newline character
#[cfg(test)]
fn parse_paragraph(p_node: &Node, add_new_line: bool) -> Result<Vec<Run>> {
    parse_paragraph_with_hyperlinks(p_node, add_new_line, &HashMap::new())
}

fn parse_paragraph_with_hyperlinks(
    p_node: &Node,
    add_new_line: bool,
    hyperlinks: &HashMap<String, String>,
) -> Result<Vec<Run>> {
    let run_nodes: Vec<_> = p_node
        .children()
        .filter(|n| {
            n.is_element()
                && n.tag_name().name() == "r"
                && n.tag_name().namespace() == Some(A_NAMESPACE)
        })
        .collect();

    let count = run_nodes.len();
    let mut runs: Vec<Run> = Vec::new();

    for (idx, r_node) in run_nodes.iter().enumerate() {
        let mut run = parse_run_with_hyperlinks(r_node, hyperlinks)?;

        if add_new_line && idx == count - 1 {
            run.text.push('\n');
        }

        runs.push(run);
    }
    Ok(runs)
}

/// Parses a single run properties node (`<a:rPr>`) and extracting the text content from the text node (`<a:t>`)
/// as well as the format including _bold_, _italic_, _underlined_ and the _language_
#[cfg(test)]
fn parse_run(r_node: &Node) -> Result<Run> {
    parse_run_with_hyperlinks(r_node, &HashMap::new())
}

fn parse_run_with_hyperlinks(r_node: &Node, hyperlinks: &HashMap<String, String>) -> Result<Run> {
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
    let link_target = r_node
        .children()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "rPr"
                && n.tag_name().namespace() == Some(A_NAMESPACE)
        })
        .and_then(|r_pr| {
            r_pr.children().find(|n| {
                n.is_element()
                    && n.tag_name().name() == "hlinkClick"
                    && n.tag_name().namespace() == Some(A_NAMESPACE)
            })
        })
        .and_then(|link| {
            link.attribute((RELS_NAMESPACE, "id"))
                .or_else(|| link.attribute("r:id"))
        })
        .and_then(|id| hyperlinks.get(id))
        .cloned();

    Ok(Run { text, formatting, link_target})
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
            .find(|n| {
                n.is_element()
                    && n.tag_name().name() == "spPr"
                    && n.tag_name().namespace() == Some(P_NAMESPACE)
            })
            .and_then(|sp_pr| {
                sp_pr.children().find(|n| {
                    n.is_element()
                        && n.tag_name().name() == "xfrm"
                        && n.tag_name().namespace() == Some(A_NAMESPACE)
                })
            }),
        "graphicFrame" => node.children().find(|n| {
            n.is_element()
                && n.tag_name().name() == "xfrm"
                && n.tag_name().namespace() == Some(P_NAMESPACE)
        }),
        "grpSp" => node
            .children()
            .find(|n| {
                n.is_element()
                    && n.tag_name().name() == "grpSpPr"
                    && n.tag_name().namespace() == Some(P_NAMESPACE)
            })
            .and_then(|grp_sp_pr| {
                grp_sp_pr.children().find(|n| {
                    n.is_element()
                        && n.tag_name().name() == "xfrm"
                        && n.tag_name().namespace() == Some(A_NAMESPACE)
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
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "grpSpPr"
                && n.tag_name().namespace() == Some(P_NAMESPACE)
        })
        .and_then(|grp_sp_pr| {
            grp_sp_pr.children().find(|n| {
                n.is_element()
                    && n.tag_name().name() == "xfrm"
                    && n.tag_name().namespace() == Some(A_NAMESPACE)
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
fn extract_child_attr_i64(node: &Node, namespace: &str, child_name: &str, attr_name: &str, ) -> Option<i64> {
    node.children()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == child_name
                && n.tag_name().namespace() == Some(namespace)
        })
        .and_then(|child| child.attribute(attr_name))
        .and_then(|value| value.parse::<i64>().ok())
}

fn extract_placeholder_key(node: &Node) -> Option<PlaceholderKey> {
    let placeholder = node.descendants().find(|n| {
        n.is_element()
            && n.tag_name().name() == "ph"
            && n.tag_name().namespace() == Some(P_NAMESPACE)
    })?;

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
        .find(|n| { n.tag_name().name() == "spTree" && n.tag_name().namespace() == root.tag_name().namespace()})
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
#[path = "../tests/unit/parse_xml_tests.rs"]
mod tests;
