use super::*;
use crate::Baseline;
use std::fs;
use std::path::PathBuf;

fn fixture(name: &str) -> Vec<u8> {
    fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/unit/ooxml")
            .join(name),
    )
    .unwrap()
}

fn at_element<'a>(data: &'a [u8], namespace: &str, name: &[u8]) -> XmlReader<'a> {
    let mut xml = reader(data);
    loop {
        match event(&mut xml, "test XML").unwrap() {
            Event::Start(element) if element_is(&xml, &element, namespace, name) => return xml,
            Event::Eof => panic!("element not found"),
            _ => {}
        }
    }
}

#[test]
fn parses_text_runs_and_paragraph_boundaries() {
    let data = fixture("tx_body.xml");
    let mut xml = at_element(&data, P_NAMESPACE, b"txBody");
    let text = content_to_text(parse_text_body(&mut xml, true, &HashMap::new()).unwrap());
    assert_eq!(text.runs.len(), 3);
    assert_eq!(text.runs[0].text, "Hello");
    assert_eq!(text.runs[2].text, "!\n");

    let data = fixture("paragraph_multiple.xml");
    let mut xml = at_element(&data, A_NAMESPACE, b"p");
    let paragraph = parse_paragraph_events(&mut xml, false, &HashMap::new()).unwrap();
    assert_eq!(paragraph.runs.len(), 3);
    assert!(paragraph.runs[1].formatting.bold);
    assert!(paragraph.runs[2].formatting.italic);
    assert!(!paragraph.runs[2].text.ends_with('\n'));
}

#[test]
fn parses_run_formatting_text_and_hyperlink() {
    let data = fixture("run_styles.xml");
    let mut xml = at_element(&data, A_NAMESPACE, b"r");
    let run = parse_run_events(&mut xml, &HashMap::new()).unwrap();
    assert_eq!(run.text.trim(), "Formatted text");
    assert!(run.formatting.bold && run.formatting.italic && run.formatting.underlined);
    assert_eq!(run.formatting.lang, "de-DE");

    let data = br#"<a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:rPr><a:hlinkClick r:id="rId7"/></a:rPr><a:t>Example &amp; Co</a:t></a:r>"#;
    let mut xml = at_element(data, A_NAMESPACE, b"r");
    let links = HashMap::from([("rId7".into(), "https://example.com".into())]);
    let run = parse_run_events(&mut xml, &links).unwrap();
    assert_eq!(run.text, "Example & Co");
    assert_eq!(run.link_target.as_deref(), Some("https://example.com"));
}

#[test]
fn parses_lists_with_existing_marker_semantics() {
    let data = fixture("multilevel_list.xml");
    let mut xml = at_element(&data, P_NAMESPACE, b"txBody");
    let block = parse_text_body(&mut xml, true, &HashMap::new()).unwrap();
    assert_eq!(block.paragraphs.len(), 5);
    assert_eq!(block.paragraphs[1].list.as_ref().unwrap().level, 1);
    assert!(matches!(
        block.paragraphs[0].list.as_ref().unwrap().kind,
        ListKind::Ordered { .. }
    ));
    assert!(matches!(
        block.paragraphs[1].list.as_ref().unwrap().kind,
        ListKind::Bullet { .. }
    ));
    assert!(block.paragraphs[0]
        .runs
        .last()
        .unwrap()
        .text
        .ends_with('\n'));
}

#[test]
fn inherits_list_styles_and_preserves_explicit_list_overrides() {
    let layout = br#"<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:pPr lvl="0"/></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sldLayout>"#;
    let inherited =
        extract_inherited_positions(layout, &InheritedPositions::default()).unwrap();
    let slide = br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr><p:txBody><a:p><a:r><a:t>Inherited bullet</a:t></a:r></a:p><a:p><a:pPr><a:buNone/></a:pPr><a:r><a:t>Plain paragraph</a:t></a:r></a:p><a:p><a:pPr><a:buAutoNum type="arabicPeriod" startAt="4"/></a:pPr><a:r><a:t>Fourth</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#;

    let parsed =
        parse_slide_document_with_hyperlinks(slide, &inherited, &HashMap::new()).unwrap();
    let SlideBlockContent::Text(text) = &parsed.blocks[0].content else {
        panic!("expected semantic text block")
    };

    assert_eq!(text.role, TextRole::Body);
    assert_eq!(parsed.blocks[0].bounds, Bounds { x: 10, y: 20, width: 300, height: 400 });
    assert!(matches!(
        text.paragraphs[0].list.as_ref().map(|list| &list.kind),
        Some(ListKind::Bullet { .. })
    ));
    assert!(text.paragraphs[1].list.is_none());
    assert!(matches!(
        text.paragraphs[2].list.as_ref().map(|list| &list.kind),
        Some(ListKind::Ordered { start: 4, .. })
    ));
}

#[test]
fn parses_fields_breaks_and_extended_run_formatting() {
    let data = br#"<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:p><a:fld><a:rPr strike="sngStrike" baseline="30000" sz="1850"/><a:t>Field</a:t></a:fld><a:br/><a:r><a:t>Next</a:t></a:r></a:p></p:txBody>"#;
    let mut xml = at_element(data, P_NAMESPACE, b"txBody");
    let block = parse_text_body(&mut xml, false, &HashMap::new()).unwrap();

    assert_eq!(block.paragraphs[0].text(), "Field\nNext");
    assert!(block.paragraphs[0].runs[0].formatting.strikethrough);
    assert_eq!(
        block.paragraphs[0].runs[0].formatting.baseline,
        Baseline::Superscript
    );
    assert_eq!(
        block.paragraphs[0].runs[0].formatting.font_size_points,
        Some(18.5)
    );

    let inherited = br#"<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:p><a:pPr><a:defRPr b="1" i="1"/></a:pPr><a:r><a:rPr b="0"/><a:t>Override</a:t></a:r><a:r><a:t>Inherited</a:t></a:r></a:p></p:txBody>"#;
    let mut xml = at_element(inherited, P_NAMESPACE, b"txBody");
    let block = parse_text_body(&mut xml, false, &HashMap::new()).unwrap();
    assert!(!block.paragraphs[0].runs[0].formatting.bold);
    assert!(block.paragraphs[0].runs[0].formatting.italic);
    assert!(block.paragraphs[0].runs[1].formatting.bold);
    assert!(block.paragraphs[0].runs[1].formatting.italic);
}

#[test]
fn parses_tables_and_empty_cells() {
    let data = fixture("complex_table.xml");
    let mut xml = at_element(&data, A_NAMESPACE, b"tbl");
    let table = parse_table_events(&mut xml, &HashMap::new()).unwrap();
    assert_eq!(table.rows.len(), 2);
    assert_eq!(table.rows[0].cells.len(), 3);
    assert!(table.rows[0].cells[0].runs[0].formatting.bold);
    assert_eq!(table.rows[1].cells[0].runs.len(), 3);

    let data = fixture("empty_table.xml");
    let mut xml = at_element(&data, A_NAMESPACE, b"tbl");
    let table = parse_table_events(&mut xml, &HashMap::new()).unwrap();
    assert!(table.rows[0].cells.iter().all(|cell| cell.runs.is_empty()));
}

#[test]
fn preserves_pptx_table_spans_for_html_rendering() {
    let data = br#"<a:tbl xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:tr><a:tc gridSpan="2" rowSpan="3"><a:txBody><a:p><a:r><a:t>Merged</a:t></a:r></a:p></a:txBody></a:tc><a:tc hMerge="1"></a:tc></a:tr></a:tbl>"#;
    let mut xml = at_element(data, A_NAMESPACE, b"tbl");
    let table = parse_table_events(&mut xml, &HashMap::new()).unwrap();

    assert_eq!(table.rows[0].cells[0].column_span, 2);
    assert_eq!(table.rows[0].cells[0].row_span, 3);
    assert!(table.rows[0].cells[1].covered);
    assert_eq!(table.rows[0].cells[0].paragraphs[0].text(), "Merged");
}

#[test]
fn parses_graphics_and_picture_failures() {
    let table_fixture = String::from_utf8(fixture("simple_table.xml")).unwrap();
    let data =
        format!(r#"<p:graphicFrame xmlns:p="{P_NAMESPACE}">{table_fixture}</p:graphicFrame>"#);
    let mut xml = at_element(data.as_bytes(), P_NAMESPACE, b"graphicFrame");
    let (table, _) = parse_graphic_frame(&mut xml, &HashMap::new()).unwrap();
    assert_eq!(table.unwrap().rows.len(), 2);

    let data = fixture("pic_with_image.xml");
    let mut xml = at_element(&data, P_NAMESPACE, b"pic");
    assert_eq!(parse_picture(&mut xml).unwrap().0.id, "rId2");

    for name in ["pic_without_embed.xml", "pic_without_blip.xml"] {
        let data = fixture(name);
        let mut xml = at_element(&data, P_NAMESPACE, b"pic");
        assert!(matches!(parse_picture(&mut xml), Err(Error::ImageNotFound)));
    }
}

#[test]
fn parses_positions_groups_and_alternative_prefixes() {
    let xml = br#"
      <q:sld xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main">
        <q:cSld><q:spTree><q:grpSp><q:grpSpPr><d:xfrm><d:off x="100" y="200"/><d:ext cx="200" cy="400"/><d:chOff x="10" y="20"/><d:chExt cx="100" cy="200"/></d:xfrm></q:grpSpPr>
          <q:sp><q:nvSpPr/><q:spPr><d:xfrm><d:off x="20" y="30"/></d:xfrm></q:spPr><q:txBody><d:p><d:r><d:t>Grouped</d:t></d:r></d:p></q:txBody></q:sp>
        </q:grpSp></q:spTree></q:cSld>
      </q:sld>"#;
    let elements = parse_slide_xml(xml).unwrap();
    let grouped = elements
        .iter()
        .find_map(|element| match element {
            SlideElement::Text(_, position) => Some(position),
            _ => None,
        })
        .unwrap();
    assert_eq!(*grouped, ElementPosition { x: 120, y: 220 });
}

#[test]
fn resolves_layout_and_master_placeholder_positions() {
    let master = br#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="title" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="42" y="84"/></a:xfrm></p:spPr></p:sp></p:spTree></p:cSld></p:sldMaster>"#;
    let inherited = extract_inherited_positions(master, &InheritedPositions::default()).unwrap();
    let slide = br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="title" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:p><a:r><a:t>Title</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#;
    let elements = parse_slide_xml_with_inherited_positions(slide, &inherited).unwrap();
    assert_eq!(elements[0].position(), ElementPosition { x: 42, y: 84 });
}

#[test]
fn reads_only_body_notes_and_rejects_malformed_xml() {
    let xml = br#"<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>Note</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:nvPr><p:ph type="sldNum"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>7</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>"#;
    let notes = parse_speaker_notes_xml(xml).unwrap();
    assert_eq!(notes.len(), 1);
    assert_eq!(notes[0].runs[0].text, "Note\n");
    assert!(parse_slide_xml(b"<p:sld").is_err());
}
