use crate::types::{Slide, SlideElement, TextElement};
use crate::constants::{P_NAMESPACE, A_NAMESPACE};
use roxmltree::{Document, Node};
use crate::{Result, Error, Formatting, Run};

pub fn parse_slide_xml(xml_data: &[u8]) -> Result<Slide> {
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
                    println!("{:?}", slide);
                    // if let Some(text_element) = parse_sp_node(&child_node)? {
                    //     elements.push(SlideElement::Text(text_element));
                    // }
                },
                _ => {
                    println!("Unbekanntes Element unter spTree: {}", tag_name);
                }
            }
        }
    }
    Ok(Slide { elements })
}

fn parse_sp(sp_node: &Node) -> Result<SlideElement> {
    // Suche nach dem <p:txBody> Element
    let tx_body_node = sp_node.children().find(|n| {
        n.is_element()
            && n.tag_name().name() == "txBody"
            && n.tag_name().namespace() == Some(P_NAMESPACE)
    }).ok_or(Error::ParseError("txBody node not found"))?;
    let mut runs = Vec::new();
    // Iteriere über alle <a:p>-Knoten innerhalb von <p:txBody>
    for p_node in tx_body_node.children().filter(|n| {
        n.is_element()
            && n.tag_name().name() == "p"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
    }) {
        let mut paragraph_runs = parse_paragraph(&p_node)?;
        runs.append(&mut paragraph_runs); // Sammle die Runs aus jedem Paragraphen
    }
    // Gib am Ende alle gesammelten Runs zurück
    Ok(SlideElement::Text(TextElement { runs }))
}
fn parse_paragraph(p_node: &Node) -> Result<Vec<Run>> {
    let mut runs: Vec<Run> = Vec::new();
    // Iteriere über alle <a:r>-Knoten innerhalb des <a:p>-Knotens
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
    // Suche nach dem <a:rPr>-Element für Formatierungen
    if let Some(rPr_node) = r_node.children().find(|n| {
        n.is_element()
            && n.tag_name().name() == "rPr"
            && n.tag_name().namespace() == Some(A_NAMESPACE)
    }) {
        // Iteriere über die Kindelemente von <a:rPr> für spezifische Formatierungselemente
        for prop in rPr_node.children().filter(|n| n.is_element()) {
            match prop.tag_name().name() {
                "b" => formatting.bold = true,
                "i" => formatting.italic = true,
                "u" => formatting.underlined = true,
                "lang" => formatting.lang = prop.text().unwrap_or_default().to_string(),
                _ => {}
            }
        }
        // Alternative: Direkt Attribute von <a:rPr> auswerten
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
    // Suche nach dem <a:t> Element innerhalb des <a:r>-Knotens
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