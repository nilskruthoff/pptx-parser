use crate::types::{Slide, SlideElement, TextElement, Formatting};
use crate::constants::{P_NAMESPACE, A_NAMESPACE};
use roxmltree::{Document, Node};
use crate::{Result, Error};

pub fn parse_slide_xml(xml_data: &[u8]) -> Result<Slide> {
    let xml_str = std::str::from_utf8(xml_data).map_err(|_| Error::Unknown)?;
    let doc = Document::parse(xml_str)?;
    let mut elements = Vec::new();
    // Suche nach allen <p:sp> Elementen
    for node in doc.descendants().filter(|n| n.has_tag_name(("sp", P_NAMESPACE))) {
        // Versuche, ein TextElement aus diesem sp-Element zu parsen
        if let Some(text_element) = parse_text_shape(&node)? {
            elements.push(SlideElement::Text(text_element));
        }
    }
    Ok(Slide { elements })
}

fn parse_text_shape(node: &Node) -> Result<Option<TextElement>> {
    // Suche nach dem <p:txBody> Element
    let tx_body = node
        .children()
        .find(|n| n.has_tag_name(("txBody", P_NAMESPACE)));
    if let Some(tx_body) = tx_body {
        // Innerhalb von txBody nach <a:p> Paragraphen suchen
        for paragraph in tx_body
            .children()
            .filter(|n| n.has_tag_name(("p", A_NAMESPACE)))
        {
            // Innerhalb des Paragraphen nach Runs suchen
            for run in paragraph
                .children()
                .filter(|n| n.has_tag_name(("r", A_NAMESPACE)))
            {
                if let Some(text_element) = parse_run(&run)? {
                    // Wir nehmen an, dass jedes Run ein eigenes TextElement ist
                    return Ok(Some(text_element));
                }
            }
        }
    }
    Ok(None)
}

// src/parser.rs
fn parse_run(node: &Node) -> Result<Option<TextElement>> {
    let mut text = String::new();
    let mut formatting = Formatting::default();
    // Suche nach dem <a:t> Element innerhalb des Runs
    if let Some(t_node) = node
        .children()
        .find(|n| n.has_tag_name(("t", A_NAMESPACE)))
    {
        if let Some(t) = t_node.text() {
            text.push_str(t);
        }
    }
    // Suche nach Formatierungsinformationen im <a:rPr> Element
    if let Some(rpr_node) = node
        .children()
        .find(|n| n.has_tag_name(("rPr", A_NAMESPACE)))
    {
        parse_run_properties(&rpr_node, &mut formatting)?;
    }
    if text.is_empty() {
        Ok(None)
    } else {
        Ok(Some(TextElement { text, formatting }))
    }
}

fn parse_run_properties(node: &Node, formatting: &mut Formatting) -> Result<()> {
    // Prüfung auf Attribute wie fett, kursiv etc.
    if node.attribute("b").is_some() {
        formatting.bold = true;
    }
    if node.attribute("i").is_some() {
        formatting.italic = true;
    }
    if node.attribute("u").is_some() {
        formatting.underline = true;
    }
    Ok(())
}