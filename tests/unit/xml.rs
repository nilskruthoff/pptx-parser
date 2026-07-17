use super::*;

#[test]
fn skip_element_reports_unexpected_eof() {
    let mut xml = reader(b"<root><child>");
    assert!(matches!(event(&mut xml, "test.xml"), Ok(Event::Start(_))));
    assert!(matches!(event(&mut xml, "test.xml"), Ok(Event::Start(_))));

    let result = skip_element(&mut xml, b"child", "test.xml");

    assert!(matches!(
        result,
        Err(Error::ParseError("Unexpected end of XML element"))
    ));
}
