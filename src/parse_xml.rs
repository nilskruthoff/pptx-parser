use crate::types::{SlideElement, TextElement, TableElement, TableRow, TableCell};
use crate::constants::{P_NAMESPACE, A_NAMESPACE};
use roxmltree::{Document, Node};
use crate::{Result, Error, Formatting, Run};
use crate::SlideElement::Unknown;

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
                _ => {
                    elements.push(Unknown)
                }
            }
        }
    }

    Ok(elements)
}

fn parse_sp(sp_node: &Node) -> Result<SlideElement> {
    // Search for <p:txBody> element
    let tx_body_node = if let Some(node) = sp_node.children().find(|n| {
        n.is_element()
            && n.tag_name().name() == "txBody"
            && n.tag_name().namespace() == Some(P_NAMESPACE)
    }) {
        node
    } else {
        return Ok(Unknown);
    };

    let mut runs = Vec::new();
    // Iterate over <a:p> elements within txBody nodes
    for p_node in tx_body_node.children().filter(|n| {
        n.is_element()
            && n.tag_name().name() == "p"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
    }) {
        let mut paragraph_runs = parse_paragraph(&p_node)?;
        runs.append(&mut paragraph_runs); // collect all runs in every paragraph node
    }

    Ok(SlideElement::Text(TextElement { runs }))
}
fn parse_paragraph(p_node: &Node) -> Result<Vec<Run>> {
    let mut runs: Vec<Run> = Vec::new();

    // iterate over <a:r> elements within paragraph elements
    for r_node in p_node.children().filter(|n| {
        n.is_element()
            && n.tag_name().name() == "r"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
    }) {
        let run = parse_run(&r_node)?;
        runs.push(run);
    }
    Ok(runs)
}
fn parse_run(r_node: &Node) -> Result<Run> {
    let mut text = String::new();
    let mut formatting = Formatting::default();

    if let Some(rPr_node) = r_node.children().find(|n| {
        n.is_element()
            && n.tag_name().name() == "rPr"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
    }) {
        if let Some(b_attr) = rPr_node.attribute("b") {
            formatting.bold = b_attr == "1" || b_attr.eq_ignore_ascii_case("true");
        }
        if let Some(i_attr) = rPr_node.attribute("i") {
            formatting.italic = i_attr == "1" || i_attr.eq_ignore_ascii_case("true");
        }
        if let Some(u_attr) = rPr_node.attribute("u") {
            formatting.underlined = u_attr != "none";
        }
        if let Some(lang_attr) = rPr_node.attribute("lang") {
            formatting.lang = lang_attr.to_string();
        }
    }

    // search the <a:t> element within the run element
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

fn parse_graphic_frame(node: &Node) -> Result<Option<SlideElement>> {
    // Suche nach <a:graphicData> mit Tabellen-URI
    let graphic_data_node = node
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "graphicData"
                && n.tag_name().namespace() == Some(A_NAMESPACE)
                && n.attribute("uri") == Some("http://schemas.openxmlformats.org/drawingml/2006/table")
        });

    if let Some(graphic_data) = graphic_data_node {
        // Suche nach <a:tbl> innerhalb von <a:graphicData>
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
            let mut paragraph_runs = parse_paragraph(&p_node)?;
            runs.append(&mut paragraph_runs);
        }
    }

    Ok(TableCell { runs })
}