use super::*;
use std::fs;
use std::path::PathBuf;

fn fixture(name: &str) -> Vec<u8> {
    fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/unit/odp")
            .join(name),
    )
    .unwrap()
}

#[test]
fn indexes_pages_with_inherited_namespace_context() {
    let xml = br#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><o:body><o:presentation><d:page/><d:page><d:g/></d:page></o:presentation></o:body></o:document-content>"#;
    let pages = index_pages(xml).unwrap();
    assert_eq!(pages.len(), 2);
    assert!(
        pages[0]
            .namespaces
            .iter()
            .any(|(name, _)| name == "xmlns:d")
    );
    let fragment = page_fragment(&xml[pages[1].range.clone()], &pages[1].namespaces);
    let parsed = parse_page_fragment(&fragment, &StyleResolver::default()).unwrap();
    assert!(parsed.elements.is_empty());
}

#[test]
fn parses_fixture_page_styles_lists_tables_and_position() {
    let content = fixture("content.xml");
    let style_xml = fixture("styles.xml");
    let styles = StyleResolver::from_documents(&content, &style_xml).unwrap();
    let pages = index_pages(&content).unwrap();
    let fragment = page_fragment(&content[pages[0].range.clone()], &pages[0].namespaces);
    let parsed = parse_page_fragment(&fragment, &styles).unwrap();
    assert!(parsed.elements.iter().any(|element| matches!(
        element,
        SlideElement::Text(
            _,
            ElementPosition {
                x: 360_000,
                y: 720_000
            }
        )
    )));
    assert!(
        parsed
            .elements
            .iter()
            .any(|element| matches!(element, SlideElement::List(_, _)))
    );
    let table = parsed
        .elements
        .iter()
        .find_map(|element| match element {
            SlideElement::Table(table, _) => Some(table),
            _ => None,
        })
        .unwrap();
    assert_eq!(table.rows[0].cells.len(), 3);
    assert!(table.rows[0].cells[0].runs[0].formatting.bold);
}

#[test]
fn parses_links_special_text_notes_and_comments_in_one_page_pass() {
    let xml = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:presentation><draw:page><draw:custom-shape><text:p>A<text:s text:c="2"/><text:a xlink:href="https://example.com">link</text:a><text:tab/><text:line-break/>B</text:p></draw:custom-shape><presentation:notes><draw:custom-shape><text:p>Note</text:p></draw:custom-shape></presentation:notes><office:annotation><text:p>Comment</text:p></office:annotation></draw:page></office:presentation></office:body></office:document-content>"#;
    let pages = index_pages(xml).unwrap();
    let fragment = page_fragment(&xml[pages[0].range.clone()], &pages[0].namespaces);
    let parsed = parse_page_fragment(&fragment, &StyleResolver::default()).unwrap();
    assert_eq!(parsed.speaker_notes[0].runs[0].text, "Note\n");
    assert_eq!(parsed.comments[0].runs[0].text, "Comment\n");
    let SlideElement::Text(text, _) = &parsed.elements[0] else {
        panic!()
    };
    assert!(
        text.runs
            .iter()
            .any(|run| run.link_target.as_deref() == Some("https://example.com"))
    );
    assert_eq!(
        text.runs
            .iter()
            .map(|run| run.text.as_str())
            .collect::<String>(),
        "A  link\t\nB\n"
    );
}

#[test]
fn expands_repeated_and_spanned_table_cells() {
    let xml = br#"<table:table xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><table:table-row table:number-rows-repeated="2"><table:table-cell table:number-columns-repeated="2" table:number-columns-spanned="2"><text:p>X</text:p></table:table-cell><table:covered-table-cell/></table:table-row></table:table>"#;
    let mut reader = reader(xml);
    loop {
        match event(&mut reader, "test").unwrap() {
            Event::Start(element) if element_is(&reader, &element, TABLE_NS, b"table") => break,
            _ => {}
        }
    }
    let table = parse_table(&mut reader, &StyleResolver::default()).unwrap();
    assert_eq!(table.rows.len(), 2);
    assert_eq!(table.rows[0].cells.len(), 5);
    assert_eq!(table.rows[0].cells[0].column_span, 2);
    assert!(table.rows[0].cells[1].covered);
}

#[test]
fn maps_odp_roles_bounds_and_unsupported_elements_to_semantic_blocks() {
    let xml = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><office:body><office:presentation><draw:page><draw:frame presentation:class="title" svg:x="1cm" svg:y="2cm" svg:width="10cm" svg:height="3cm"><draw:text-box><text:p>Semantic title</text:p></draw:text-box></draw:frame><draw:line></draw:line></draw:page></office:presentation></office:body></office:document-content>"#;
    let pages = index_pages(xml).unwrap();
    let fragment = page_fragment(&xml[pages[0].range.clone()], &pages[0].namespaces);
    let parsed = parse_page_fragment(&fragment, &StyleResolver::default()).unwrap();

    let SlideBlockContent::Text(title) = &parsed.blocks[0].content else {
        panic!("expected title block")
    };
    assert_eq!(title.role, TextRole::Title);
    assert_eq!(title.paragraphs[0].text(), "Semantic title\n");
    assert_eq!(parsed.blocks[0].bounds.width, 3_600_000);
    assert_eq!(parsed.blocks[0].bounds.height, 1_080_000);
    assert!(matches!(
        parsed.blocks[1].content,
        SlideBlockContent::Unsupported(_)
    ));
    assert_eq!(parsed.diagnostics.len(), 1);
}

#[test]
fn converts_lengths_transforms_and_reports_bad_xml() {
    assert_eq!(parse_length("1cm"), Some(360_000));
    assert_eq!(parse_length("2.5mm"), Some(90_000));
    assert_eq!(
        parse_translate("translate(1cm 2cm)"),
        ElementPosition {
            x: 360_000,
            y: 720_000
        }
    );
    assert!(index_pages(b"<broken").is_err());
}

#[test]
fn parses_empty_images_accessible_titles_frame_tables_and_headings() {
    let xml = br#"
      <office:document-content
          xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
          xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
          xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
          xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
          xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
          xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
          xmlns:xlink="http://www.w3.org/1999/xlink">
        <office:body><office:presentation><draw:page>
          <draw:image xlink:href="Pictures/direct.png" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"/>
          <draw:frame draw:name="Frame fallback" svg:x="5cm" svg:y="1cm" svg:width="2cm" svg:height="2cm">
            <draw:image xlink:href="Pictures/framed.png"></draw:image>
            <svg:title>Accessible &amp; image</svg:title>
          </draw:frame>
          <draw:frame svg:x="1cm" svg:y="7cm" svg:width="6cm" svg:height="2cm">
            <table:table><table:table-row><table:table-cell><text:p>Frame cell</text:p></table:table-cell></table:table-row></table:table>
          </draw:frame>
          <draw:frame svg:x="1cm" svg:y="10cm" svg:width="6cm" svg:height="1cm">
            <draw:text-box><text:h>Section heading</text:h></draw:text-box>
          </draw:frame>
          <presentation:notes><draw:frame><draw:image xlink:href="Pictures/note.png"/></draw:frame></presentation:notes>
        </draw:page></office:presentation></office:body>
      </office:document-content>"#;

    let pages = index_pages(xml).unwrap();
    let fragment = page_fragment(&xml[pages[0].range.clone()], &pages[0].namespaces);
    let parsed = parse_page_fragment(&fragment, &StyleResolver::default()).unwrap();

    assert_eq!(parsed.elements.len(), 4);
    let SlideBlockContent::Image(direct) = &parsed.blocks[0].content else {
        panic!("expected direct image")
    };
    assert_eq!(direct.reference.target, "Pictures/direct.png");
    assert_eq!(parsed.blocks[0].bounds.x, 360_000);
    assert_eq!(parsed.blocks[0].bounds.height, 1_440_000);

    let SlideBlockContent::Image(framed) = &parsed.blocks[1].content else {
        panic!("expected framed image")
    };
    assert_eq!(framed.alt_text.as_deref(), Some("Accessible & image"));

    let SlideBlockContent::Table(table) = &parsed.blocks[2].content else {
        panic!("expected frame table")
    };
    assert_eq!(table.rows[0].cells[0].paragraphs[0].text(), "Frame cell\n");

    let SlideBlockContent::Text(heading) = &parsed.blocks[3].content else {
        panic!("expected heading")
    };
    assert_eq!(heading.role, TextRole::Heading);
    assert_eq!(heading.paragraphs[0].text(), "Section heading\n");
    assert!(parsed.speaker_notes.is_empty());
}
