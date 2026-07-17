use crate::{Error, Result};
use quick_xml::events::{BytesRef, BytesStart, BytesText, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::borrow::Cow;

pub(crate) type XmlReader<'a> = NsReader<&'a [u8]>;

pub(crate) fn reader(data: &[u8]) -> XmlReader<'_> {
    let mut reader = NsReader::from_reader(data);
    reader.config_mut().trim_text(false);
    reader
}

pub(crate) fn event<'a>(reader: &mut XmlReader<'a>, part: &str) -> Result<Event<'a>> {
    reader.read_event().map_err(|source| Error::Xml {
        part: part.to_string(),
        source,
    })
}

pub(crate) fn local(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

pub(crate) fn element_is(
    reader: &XmlReader<'_>,
    element: &BytesStart<'_>,
    namespace: &str,
    name: &[u8],
) -> bool {
    let (resolved, local_name) = reader.resolver().resolve_element(element.name());
    local_name.as_ref() == name
        && matches!(resolved, ResolveResult::Bound(value) if value.as_ref() == namespace.as_bytes())
}

pub(crate) fn end_is(name: &[u8], wanted: &[u8]) -> bool {
    local(name) == wanted
}

pub(crate) fn attr(element: &BytesStart<'_>, wanted: &[u8]) -> Option<String> {
    element
        .attributes()
        .with_checks(false)
        .flatten()
        .find_map(|attribute| {
            (local(attribute.key.as_ref()) == wanted).then(|| {
                let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                quick_xml::escape::unescape(&value)
                    .map(Cow::into_owned)
                    .unwrap_or(value)
            })
        })
}

pub(crate) fn text(event: &BytesText<'_>, part: &str) -> Result<String> {
    let decoded = event.decode().map_err(|source| Error::Xml {
        part: part.to_string(),
        source: source.into(),
    })?;
    Ok(quick_xml::escape::unescape(&decoded)
        .map(Cow::into_owned)
        .unwrap_or_else(|_| decoded.into_owned()))
}

pub(crate) fn reference(event: &BytesRef<'_>, part: &str) -> Result<String> {
    let decoded = event.decode().map_err(|source| Error::Xml {
        part: part.to_string(),
        source: source.into(),
    })?;
    let escaped = format!("&{decoded};");
    Ok(quick_xml::escape::unescape(&escaped)
        .map(Cow::into_owned)
        .unwrap_or(escaped))
}

pub(crate) fn skip_element(reader: &mut XmlReader<'_>, end: &[u8], part: &str) -> Result<()> {
    let mut depth = 1usize;
    while depth > 0 {
        match event(reader, part)? {
            Event::Start(_) => depth += 1,
            Event::End(element) => {
                depth -= 1;
                if depth == 0 && end_is(element.name().as_ref(), end) {
                    break;
                }
            }
            Event::Eof => return Err(Error::ParseError("Unexpected end of XML element")),
            _ => {}
        }
    }
    Ok(())
}

#[cfg(test)]
#[path = "../tests/unit/xml.rs"]
mod tests;
