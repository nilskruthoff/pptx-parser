use crate::metadata::{parse_odp_metadata, render_presentation_markdown};
use crate::xml::{
    XmlReader, attr, element_is, end_is, event, reader, reference, skip_element, text,
};
use crate::{
    ElementPosition, Error, Formatting, ImageReference, ListElement, ListItem, Paragraph,
    ParseDiagnostic, ParserConfig, PresentationMetadata, Result, Run, Slide, SlideBlock,
    SlideBlockContent, SlideElement, TableCell, TableElement, TableRow, TextBlock, TextElement,
    TextRole,
};
use quick_xml::events::{BytesStart, Event};
use std::collections::HashMap;
use std::io::Read;
use std::ops::Range;
use std::path::Path;

const DRAW_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const STYLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const TABLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const PRESENTATION_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";

struct PageIndex {
    range: Range<usize>,
    namespaces: Vec<(String, String)>,
}

pub(crate) struct OdpContainer {
    config: ParserConfig,
    archive: zip::ZipArchive<std::fs::File>,
    content: Vec<u8>,
    pages: Vec<PageIndex>,
    styles: StyleResolver,
    metadata: PresentationMetadata,
}

impl OdpContainer {
    pub(crate) fn open(path: &Path, config: ParserConfig) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        let mut archive = zip::ZipArchive::new(file)?;
        let content = read_archive_file(&mut archive, "content.xml")?;
        let style_xml = read_archive_file(&mut archive, "styles.xml").unwrap_or_default();
        let meta = read_optional_archive_file(&mut archive, "meta.xml")?;
        let styles = StyleResolver::from_documents(&content, &style_xml)?;
        let pages = index_pages(&content)?;
        let metadata = parse_odp_metadata(meta.as_deref())?;
        Ok(Self {
            config,
            archive,
            content,
            pages,
            styles,
            metadata,
        })
    }

    pub(crate) fn parse_all(&mut self) -> Result<Vec<Slide>> {
        (0..self.pages.len())
            .map(|index| self.load_slide(index))
            .collect()
    }

    pub(crate) fn metadata(&self) -> &PresentationMetadata {
        &self.metadata
    }

    pub(crate) fn convert_to_md(&mut self) -> Result<String> {
        let slides = self.parse_all()?;
        render_presentation_markdown(
            &self.metadata,
            self.config.include_presentation_metadata,
            slides,
        )
    }

    fn load_slide(&mut self, index: usize) -> Result<Slide> {
        let page = self.pages.get(index).ok_or(Error::SlideNotFound)?;
        let fragment = page_fragment(&self.content[page.range.clone()], &page.namespaces);
        let mut parsed = parse_page_fragment(&fragment, &self.styles)?;
        let images: Vec<ImageReference> = parsed
            .elements
            .iter()
            .filter_map(|element| match element {
                SlideElement::Image(image, _) => Some(image.clone()),
                _ => None,
            })
            .collect();
        let mut image_data = HashMap::new();
        if self.config.extract_images {
            for image in &images {
                match read_archive_file(&mut self.archive, &image.target) {
                    Ok(data) => {
                        image_data.insert(image.id.clone(), data);
                    }
                    Err(error) => parsed.diagnostics.push(ParseDiagnostic {
                        severity: crate::DiagnosticSeverity::Warning,
                        message: format!("Image resource could not be loaded: {error}"),
                        source: Some(image.target.clone()),
                    }),
                }
            }
        }
        Ok(Slide::new_semantic(
            format!("content.xml#page{}", index + 1),
            (index + 1) as u32,
            parsed.elements,
            parsed.blocks,
            parsed.speaker_notes,
            parsed.comments,
            images,
            image_data,
            self.config.clone(),
            parsed.diagnostics,
        ))
    }

    pub(crate) fn iter_slides(&mut self) -> OdpSlideIterator<'_> {
        OdpSlideIterator {
            container: self,
            current_index: 0,
        }
    }
}

pub(crate) struct OdpSlideIterator<'a> {
    container: &'a mut OdpContainer,
    current_index: usize,
}

impl Iterator for OdpSlideIterator<'_> {
    type Item = Result<Slide>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.current_index >= self.container.pages.len() {
            return None;
        }
        let index = self.current_index;
        self.current_index += 1;
        Some(self.container.load_slide(index))
    }
}

fn read_archive_file(archive: &mut zip::ZipArchive<std::fs::File>, path: &str) -> Result<Vec<u8>> {
    let mut file = archive.by_name(path)?;
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes)?;
    Ok(bytes)
}

fn read_optional_archive_file(
    archive: &mut zip::ZipArchive<std::fs::File>,
    path: &str,
) -> Result<Option<Vec<u8>>> {
    let mut file = match archive.by_name(path) {
        Ok(file) => file,
        Err(zip::result::ZipError::FileNotFound) => return Ok(None),
        Err(error) => return Err(error.into()),
    };
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes)?;
    Ok(Some(bytes))
}

fn index_pages(content: &[u8]) -> Result<Vec<PageIndex>> {
    let mut xml = reader(content);
    let mut namespace_values: HashMap<String, Vec<String>> = HashMap::new();
    let mut scopes: Vec<Vec<String>> = Vec::new();
    let mut pages = Vec::new();
    let mut page_start = None;
    let mut page_depth = 0usize;
    let mut page_namespaces = Vec::new();
    loop {
        let event_start = xml.buffer_position() as usize;
        match event(&mut xml, "ODP content.xml")? {
            Event::Start(element) => {
                let declared = push_namespace_declarations(&element, &mut namespace_values);
                scopes.push(declared);
                if page_start.is_some() {
                    page_depth += 1;
                } else if element_is(&xml, &element, DRAW_NS, b"page") {
                    page_start = Some(event_start);
                    page_depth = 1;
                    page_namespaces = namespace_values
                        .iter()
                        .filter_map(|(name, values)| {
                            values.last().map(|value| (name.clone(), value.clone()))
                        })
                        .collect();
                }
            }
            Event::Empty(element)
                if page_start.is_none() && element_is(&xml, &element, DRAW_NS, b"page") =>
            {
                let mut namespaces: Vec<_> = namespace_values
                    .iter()
                    .filter_map(|(name, values)| {
                        values.last().map(|value| (name.clone(), value.clone()))
                    })
                    .collect();
                for attribute in element.attributes().with_checks(false).flatten() {
                    let name = String::from_utf8_lossy(attribute.key.as_ref());
                    if name == "xmlns" || name.starts_with("xmlns:") {
                        namespaces.push((
                            name.into_owned(),
                            String::from_utf8_lossy(attribute.value.as_ref()).into_owned(),
                        ));
                    }
                }
                pages.push(PageIndex {
                    range: event_start..xml.buffer_position() as usize,
                    namespaces,
                });
            }
            Event::End(_) => {
                if page_start.is_some() {
                    page_depth -= 1;
                    if page_depth == 0 {
                        pages.push(PageIndex {
                            range: page_start.take().unwrap()..xml.buffer_position() as usize,
                            namespaces: std::mem::take(&mut page_namespaces),
                        });
                    }
                }
                if let Some(declared) = scopes.pop() {
                    pop_namespace_declarations(declared, &mut namespace_values);
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if page_start.is_some() {
        return Err(Error::ParseError("Unexpected end of ODP page"));
    }
    Ok(pages)
}

fn push_namespace_declarations(
    element: &BytesStart<'_>,
    namespaces: &mut HashMap<String, Vec<String>>,
) -> Vec<String> {
    let mut declared = Vec::new();
    for attribute in element.attributes().with_checks(false).flatten() {
        let name = String::from_utf8_lossy(attribute.key.as_ref());
        if name == "xmlns" || name.starts_with("xmlns:") {
            let name = name.into_owned();
            namespaces
                .entry(name.clone())
                .or_default()
                .push(String::from_utf8_lossy(attribute.value.as_ref()).into_owned());
            declared.push(name);
        }
    }
    declared
}

fn pop_namespace_declarations(
    declared: Vec<String>,
    namespaces: &mut HashMap<String, Vec<String>>,
) {
    for name in declared {
        let remove = if let Some(values) = namespaces.get_mut(&name) {
            values.pop();
            values.is_empty()
        } else {
            false
        };
        if remove {
            namespaces.remove(&name);
        }
    }
}

fn page_fragment(page: &[u8], namespaces: &[(String, String)]) -> Vec<u8> {
    let mut fragment = Vec::with_capacity(page.len() + namespaces.len() * 48 + 32);
    fragment.extend_from_slice(b"<odp-fragment");
    for (name, value) in namespaces {
        fragment.push(b' ');
        fragment.extend_from_slice(name.as_bytes());
        fragment.extend_from_slice(b"=\"");
        fragment.extend_from_slice(quick_xml::escape::escape(value).as_bytes());
        fragment.push(b'"');
    }
    fragment.push(b'>');
    fragment.extend_from_slice(page);
    fragment.extend_from_slice(b"</odp-fragment>");
    fragment
}

#[derive(Default, Clone)]
struct PartialFormatting {
    bold: Option<bool>,
    italic: Option<bool>,
    underlined: Option<bool>,
    strikethrough: Option<bool>,
    baseline: Option<crate::Baseline>,
    font_size_points: Option<f32>,
    lang: Option<String>,
}

impl PartialFormatting {
    fn merge_into(&self, formatting: &mut Formatting) {
        if let Some(value) = self.bold {
            formatting.bold = value;
        }
        if let Some(value) = self.italic {
            formatting.italic = value;
        }
        if let Some(value) = self.underlined {
            formatting.underlined = value;
        }
        if let Some(value) = self.strikethrough {
            formatting.strikethrough = value;
        }
        if let Some(value) = self.baseline {
            formatting.baseline = value;
        }
        if let Some(value) = self.font_size_points {
            formatting.font_size_points = Some(value);
        }
        if let Some(value) = &self.lang {
            formatting.lang = value.clone();
        }
    }
}

struct StyleDefinition {
    parent: Option<String>,
    formatting: PartialFormatting,
}

#[derive(Default)]
struct StyleResolver {
    styles: HashMap<String, StyleDefinition>,
    list_levels: HashMap<(String, u32), bool>,
}

impl StyleResolver {
    fn from_documents(content: &[u8], styles: &[u8]) -> Result<Self> {
        let mut resolver = Self::default();
        for (part, data) in [("ODP content.xml", content), ("ODP styles.xml", styles)] {
            if !data.is_empty() {
                resolver.collect(data, part)?;
            }
        }
        Ok(resolver)
    }

    fn collect(&mut self, data: &[u8], part: &str) -> Result<()> {
        let mut xml = reader(data);
        loop {
            match event(&mut xml, part)? {
                Event::Start(element) if element_is(&xml, &element, STYLE_NS, b"style") => {
                    self.parse_style(&mut xml, &element, part)?;
                }
                Event::Start(element) if element_is(&xml, &element, TEXT_NS, b"list-style") => {
                    self.parse_list_style(&mut xml, &element, part)?;
                }
                Event::Eof => break,
                _ => {}
            }
        }
        Ok(())
    }

    fn parse_style(
        &mut self,
        xml: &mut XmlReader<'_>,
        start: &BytesStart<'_>,
        part: &str,
    ) -> Result<()> {
        let name = attr(start, b"name");
        let parent = attr(start, b"parent-style-name");
        let mut formatting = PartialFormatting::default();
        loop {
            match event(xml, part)? {
                Event::Start(element) | Event::Empty(element)
                    if element_is(xml, &element, STYLE_NS, b"text-properties") =>
                {
                    formatting = parse_formatting_properties(&element);
                }
                Event::End(element) if end_is(element.name().as_ref(), b"style") => break,
                Event::Eof => return Err(Error::ParseError("Unexpected end of ODP style")),
                _ => {}
            }
        }
        if let Some(name) = name {
            self.styles
                .insert(name, StyleDefinition { parent, formatting });
        }
        Ok(())
    }

    fn parse_list_style(
        &mut self,
        xml: &mut XmlReader<'_>,
        start: &BytesStart<'_>,
        part: &str,
    ) -> Result<()> {
        let name = attr(start, b"name");
        loop {
            match event(xml, part)? {
                Event::Start(element) | Event::Empty(element) => {
                    let ordered = element_is(xml, &element, TEXT_NS, b"list-level-style-number");
                    let bullet = element_is(xml, &element, TEXT_NS, b"list-level-style-bullet");
                    if (ordered || bullet)
                        && let Some(name) = &name
                    {
                        let level = attr(&element, b"level")
                            .and_then(|value| value.parse().ok())
                            .unwrap_or(1);
                        self.list_levels.insert((name.clone(), level), ordered);
                    }
                }
                Event::End(element) if end_is(element.name().as_ref(), b"list-style") => break,
                Event::Eof => return Err(Error::ParseError("Unexpected end of ODP list style")),
                _ => {}
            }
        }
        Ok(())
    }

    fn formatting(&self, style_name: Option<&str>, base: Formatting) -> Formatting {
        let mut chain = Vec::new();
        let mut current = style_name;
        while let Some(name) = current {
            let Some(style) = self.styles.get(name) else {
                break;
            };
            chain.push(&style.formatting);
            current = style.parent.as_deref();
        }
        let mut result = base;
        for formatting in chain.into_iter().rev() {
            formatting.merge_into(&mut result);
        }
        result
    }

    fn is_ordered_list(&self, style_name: Option<&str>, level: u32) -> bool {
        style_name
            .and_then(|name| self.list_levels.get(&(name.to_string(), level + 1)))
            .copied()
            .unwrap_or(false)
    }
}

fn parse_formatting_properties(element: &BytesStart<'_>) -> PartialFormatting {
    PartialFormatting {
        bold: attr(element, b"font-weight").map(|value| is_enabled(&value)),
        italic: attr(element, b"font-style").map(|value| value.eq_ignore_ascii_case("italic")),
        underlined: attr(element, b"text-underline-style")
            .map(|value| !value.eq_ignore_ascii_case("none")),
        strikethrough: attr(element, b"text-line-through-style")
            .map(|value| !value.eq_ignore_ascii_case("none")),
        baseline: attr(element, b"text-position").and_then(|value| {
            let value = value.to_ascii_lowercase();
            if value.starts_with("super") || value.starts_with('+') {
                Some(crate::Baseline::Superscript)
            } else if value.starts_with("sub") || value.starts_with('-') {
                Some(crate::Baseline::Subscript)
            } else {
                None
            }
        }),
        font_size_points: attr(element, b"font-size")
            .and_then(|value| value.strip_suffix("pt").map(str::to_string))
            .and_then(|value| value.parse().ok()),
        lang: attr(element, b"language"),
    }
}

fn is_enabled(value: &str) -> bool {
    matches!(value, "bold" | "700" | "800" | "900")
}

#[derive(Default)]
struct ParsedPage {
    elements: Vec<SlideElement>,
    blocks: Vec<SlideBlock>,
    speaker_notes: Vec<TextElement>,
    comments: Vec<TextElement>,
    diagnostics: Vec<ParseDiagnostic>,
}

enum PageSection {
    Main,
    Notes,
    Comment,
}

struct TextContainerContext {
    position: ElementPosition,
    bounds: crate::Bounds,
    section: PageSection,
    role: TextRole,
}

fn parse_page_fragment(data: &[u8], styles: &StyleResolver) -> Result<ParsedPage> {
    let mut xml = reader(data);
    loop {
        match event(&mut xml, "ODP page")? {
            Event::Start(element) if element_is(&xml, &element, DRAW_NS, b"page") => {
                let mut page = ParsedPage::default();
                parse_container(
                    &mut xml,
                    b"page",
                    ElementPosition::default(),
                    PageSection::Main,
                    styles,
                    &mut page,
                )?;
                return Ok(page);
            }
            Event::Empty(element) if element_is(&xml, &element, DRAW_NS, b"page") => {
                return Ok(ParsedPage::default());
            }
            Event::Eof => return Err(Error::ParseError("ODP page not found")),
            _ => {}
        }
    }
}

fn parse_container(
    xml: &mut XmlReader<'_>,
    end: &[u8],
    parent_position: ElementPosition,
    section: PageSection,
    styles: &StyleResolver,
    page: &mut ParsedPage,
) -> Result<()> {
    loop {
        match event(xml, "ODP page")? {
            Event::Start(element) => {
                if element_is(xml, &element, PRESENTATION_NS, b"notes") {
                    parse_container(
                        xml,
                        b"notes",
                        ElementPosition::default(),
                        PageSection::Notes,
                        styles,
                        page,
                    )?;
                } else if crate::xml::local(element.name().as_ref()) == b"annotation" {
                    let position = add_position(parent_position, node_position(&element));
                    let bounds = node_bounds(&element, position);
                    parse_text_container(
                        xml,
                        b"annotation",
                        TextContainerContext {
                            position,
                            bounds,
                            section: PageSection::Comment,
                            role: TextRole::Other,
                        },
                        styles,
                        page,
                    )?;
                } else {
                    parse_node(xml, &element, parent_position, &section, styles, page)?;
                }
            }
            Event::Empty(element) => {
                parse_empty_node(&element, parent_position, &section, page);
            }
            Event::End(element) if end_is(element.name().as_ref(), end) => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of ODP container")),
            _ => {}
        }
    }
    Ok(())
}

fn parse_node(
    xml: &mut XmlReader<'_>,
    start: &BytesStart<'_>,
    parent_position: ElementPosition,
    section: &PageSection,
    styles: &StyleResolver,
    page: &mut ParsedPage,
) -> Result<()> {
    let position = add_position(parent_position, node_position(start));
    let bounds = node_bounds(start, position);
    if element_is(xml, start, DRAW_NS, b"g") {
        parse_container(xml, b"g", position, clone_section(section), styles, page)
    } else if element_is(xml, start, DRAW_NS, b"frame") {
        parse_frame(xml, start, position, bounds, section, styles, page)
    } else if element_is(xml, start, DRAW_NS, b"custom-shape") {
        parse_text_container(
            xml,
            b"custom-shape",
            TextContainerContext {
                position,
                bounds,
                section: clone_section(section),
                role: TextRole::Other,
            },
            styles,
            page,
        )
    } else if element_is(xml, start, TABLE_NS, b"table") {
        let table = parse_table(xml, styles)?;
        push_element(SlideElement::Table(table, position), section, page);
        set_last_bounds(page, section, bounds);
        Ok(())
    } else {
        let end = start.name().as_ref().to_vec();
        skip_element(xml, &end, "ODP page")?;
        if matches!(section, PageSection::Main) {
            push_unsupported_odp(page, crate::xml::local(start.name().as_ref()));
        }
        Ok(())
    }
}

fn clone_section(section: &PageSection) -> PageSection {
    match section {
        PageSection::Main => PageSection::Main,
        PageSection::Notes => PageSection::Notes,
        PageSection::Comment => PageSection::Comment,
    }
}

fn parse_empty_node(
    element: &BytesStart<'_>,
    parent_position: ElementPosition,
    section: &PageSection,
    page: &mut ParsedPage,
) {
    if crate::xml::local(element.name().as_ref()) == b"image"
        && let Some(reference) = parse_image(element)
    {
        push_element(
            SlideElement::Image(
                reference,
                add_position(parent_position, node_position(element)),
            ),
            section,
            page,
        );
        set_last_bounds(
            page,
            section,
            node_bounds(
                element,
                add_position(parent_position, node_position(element)),
            ),
        );
    }
}

fn parse_frame(
    xml: &mut XmlReader<'_>,
    start: &BytesStart<'_>,
    position: ElementPosition,
    bounds: crate::Bounds,
    section: &PageSection,
    styles: &StyleResolver,
    page: &mut ParsedPage,
) -> Result<()> {
    let role = odp_text_role(attr(start, b"class").as_deref());
    let mut alt_text = attr(start, b"name");
    loop {
        match event(xml, "ODP frame")? {
            Event::Start(element) if element_is(xml, &element, DRAW_NS, b"image") => {
                if let Some(reference) = parse_image(&element) {
                    push_element(SlideElement::Image(reference, position), section, page);
                    set_last_bounds(page, section, bounds);
                    set_last_image_alt(page, section, alt_text.clone());
                }
                skip_element(xml, b"image", "ODP image")?;
            }
            Event::Empty(element) if element_is(xml, &element, DRAW_NS, b"image") => {
                if let Some(reference) = parse_image(&element) {
                    push_element(SlideElement::Image(reference, position), section, page);
                    set_last_bounds(page, section, bounds);
                    set_last_image_alt(page, section, alt_text.clone());
                }
            }
            Event::Start(element)
                if matches!(
                    crate::xml::local(element.name().as_ref()),
                    b"title" | b"desc"
                ) =>
            {
                let end = crate::xml::local(element.name().as_ref()).to_vec();
                let value = read_odp_simple_text(xml, &end)?;
                if !value.trim().is_empty() {
                    alt_text = Some(value.trim().to_string());
                    set_last_image_alt(page, section, alt_text.clone());
                }
            }
            Event::Start(element) if element_is(xml, &element, DRAW_NS, b"text-box") => {
                parse_text_container(
                    xml,
                    b"text-box",
                    TextContainerContext {
                        position,
                        bounds,
                        section: clone_section(section),
                        role,
                    },
                    styles,
                    page,
                )?;
            }
            Event::Start(element) if element_is(xml, &element, TABLE_NS, b"table") => {
                let table = parse_table(xml, styles)?;
                push_element(SlideElement::Table(table, position), section, page);
                set_last_bounds(page, section, bounds);
            }
            Event::End(element) if end_is(element.name().as_ref(), b"frame") => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of ODP frame")),
            _ => {}
        }
    }
    Ok(())
}

fn set_last_image_alt(page: &mut ParsedPage, section: &PageSection, alt_text: Option<String>) {
    if !matches!(section, PageSection::Main) {
        return;
    }
    if let Some(SlideBlock {
        content: SlideBlockContent::Image(image),
        ..
    }) = page.blocks.last_mut()
    {
        image.alt_text = alt_text;
    }
}

fn read_odp_simple_text(xml: &mut XmlReader<'_>, end: &[u8]) -> Result<String> {
    let mut value = String::new();
    loop {
        match event(xml, "ODP accessible text")? {
            Event::Text(content) => value.push_str(&text(&content, "ODP accessible text")?),
            Event::GeneralRef(content) => {
                value.push_str(&reference(&content, "ODP accessible text")?)
            }
            Event::End(element) if end_is(element.name().as_ref(), end) => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of ODP accessible text")),
            _ => {}
        }
    }
    Ok(value)
}

fn parse_text_container(
    xml: &mut XmlReader<'_>,
    end: &[u8],
    context: TextContainerContext,
    styles: &StyleResolver,
    page: &mut ParsedPage,
) -> Result<()> {
    let TextContainerContext {
        position,
        bounds,
        section,
        role,
    } = context;
    let mut paragraphs = Vec::new();
    loop {
        match event(xml, "ODP text container")? {
            Event::Start(element)
                if element_is(xml, &element, TEXT_NS, b"p")
                    || element_is(xml, &element, TEXT_NS, b"h") =>
            {
                let paragraph_role =
                    if element_is(xml, &element, TEXT_NS, b"h") && role == TextRole::Other {
                        TextRole::Heading
                    } else {
                        role
                    };
                if paragraph_role != role {
                    flush_odp_text(&mut paragraphs, position, bounds, role, &section, page);
                }
                let mut runs = parse_paragraph(xml, &element, styles)?;
                if let Some(last) = runs.last_mut() {
                    last.text.push('\n');
                }
                if paragraph_role != role {
                    let mut heading = vec![Paragraph::plain(runs)];
                    flush_odp_text(
                        &mut heading,
                        position,
                        bounds,
                        paragraph_role,
                        &section,
                        page,
                    );
                } else {
                    paragraphs.push(Paragraph::plain(runs));
                }
            }
            Event::Start(element) if element_is(xml, &element, TEXT_NS, b"list") => {
                flush_odp_text(&mut paragraphs, position, bounds, role, &section, page);
                let list = parse_list(xml, &element, styles, 0)?;
                push_element(SlideElement::List(list, position), &section, page);
            }
            Event::End(element) if end_is(element.name().as_ref(), end) => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of ODP text container")),
            _ => {}
        }
    }
    flush_odp_text(&mut paragraphs, position, bounds, role, &section, page);
    Ok(())
}

fn flush_odp_text(
    paragraphs: &mut Vec<Paragraph>,
    position: ElementPosition,
    bounds: crate::Bounds,
    role: TextRole,
    section: &PageSection,
    page: &mut ParsedPage,
) {
    if paragraphs.is_empty() {
        return;
    }
    let runs = paragraphs
        .iter()
        .flat_map(|paragraph| paragraph.runs.clone())
        .collect();
    push_element(
        SlideElement::Text(TextElement { runs }, position),
        section,
        page,
    );
    set_last_bounds(page, section, bounds);
    if matches!(section, PageSection::Main)
        && let Some(SlideBlock {
            content: SlideBlockContent::Text(text),
            ..
        }) = page.blocks.last_mut()
    {
        *text = TextBlock {
            role,
            paragraphs: std::mem::take(paragraphs),
        };
    } else {
        paragraphs.clear();
    }
}

fn push_element(element: SlideElement, section: &PageSection, page: &mut ParsedPage) {
    match section {
        PageSection::Main => {
            let source_order = page.blocks.len();
            page.blocks
                .push(crate::slide::legacy_block(&element, source_order));
            page.elements.push(element);
        }
        PageSection::Notes => {
            if let SlideElement::Text(text, _) = element {
                page.speaker_notes.push(text);
            }
        }
        PageSection::Comment => {
            if let SlideElement::Text(text, _) = element {
                page.comments.push(text);
            }
        }
    }
}

fn push_unsupported_odp(page: &mut ParsedPage, kind: &[u8]) {
    let kind = String::from_utf8_lossy(kind).into_owned();
    page.diagnostics.push(ParseDiagnostic {
        severity: crate::DiagnosticSeverity::Warning,
        message: format!("Unsupported ODP slide element: {kind}"),
        source: None,
    });
    let source_order = page.blocks.len();
    page.blocks.push(SlideBlock {
        bounds: crate::Bounds::default(),
        source_order,
        content: SlideBlockContent::Unsupported(crate::UnsupportedBlock {
            kind,
            fallback_text: None,
        }),
    });
    page.elements.push(SlideElement::Unknown);
}

fn set_last_bounds(page: &mut ParsedPage, section: &PageSection, bounds: crate::Bounds) {
    if matches!(section, PageSection::Main)
        && let Some(block) = page.blocks.last_mut()
    {
        block.bounds = bounds;
    }
}

fn odp_text_role(value: Option<&str>) -> TextRole {
    match value {
        Some("title") => TextRole::Title,
        Some("subtitle") => TextRole::Subtitle,
        Some("outline") | Some("body") => TextRole::Body,
        Some("notes") => TextRole::Body,
        _ => TextRole::Other,
    }
}

fn parse_image(element: &BytesStart<'_>) -> Option<ImageReference> {
    let target = attr(element, b"href")?;
    Some(ImageReference {
        id: target.clone(),
        target,
    })
}

fn parse_paragraph(
    xml: &mut XmlReader<'_>,
    start: &BytesStart<'_>,
    styles: &StyleResolver,
) -> Result<Vec<Run>> {
    let formatting =
        styles.formatting(attr(start, b"style-name").as_deref(), Formatting::default());
    collect_runs(
        xml,
        crate::xml::local(start.name().as_ref()),
        formatting,
        styles,
        None,
    )
}

fn collect_runs(
    xml: &mut XmlReader<'_>,
    end: &[u8],
    formatting: Formatting,
    styles: &StyleResolver,
    link_target: Option<String>,
) -> Result<Vec<Run>> {
    let mut runs = Vec::new();
    loop {
        match event(xml, "ODP text")? {
            Event::Text(content) => {
                let value = text(&content, "ODP text")?;
                if !value.is_empty() {
                    runs.push(Run {
                        text: value,
                        formatting: formatting.clone(),
                        link_target: link_target.clone(),
                    });
                }
            }
            Event::GeneralRef(content) => {
                let value = reference(&content, "ODP text")?;
                if !value.is_empty() {
                    runs.push(Run {
                        text: value,
                        formatting: formatting.clone(),
                        link_target: link_target.clone(),
                    });
                }
            }
            Event::Start(element) if element_is(xml, &element, TEXT_NS, b"span") => {
                let next =
                    styles.formatting(attr(&element, b"style-name").as_deref(), formatting.clone());
                runs.extend(collect_runs(
                    xml,
                    b"span",
                    next,
                    styles,
                    link_target.clone(),
                )?);
            }
            Event::Start(element) if element_is(xml, &element, TEXT_NS, b"a") => {
                let target = attr(&element, b"href").or_else(|| link_target.clone());
                runs.extend(collect_runs(xml, b"a", formatting.clone(), styles, target)?);
            }
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, TEXT_NS, b"line-break") =>
            {
                runs.push(special_run("\n", &formatting, link_target.as_deref()));
            }
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, TEXT_NS, b"tab") =>
            {
                runs.push(special_run("\t", &formatting, link_target.as_deref()));
            }
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, TEXT_NS, b"s") =>
            {
                let count = attr(&element, b"c")
                    .and_then(|value| value.parse().ok())
                    .unwrap_or(1);
                runs.push(special_run(
                    &" ".repeat(count),
                    &formatting,
                    link_target.as_deref(),
                ));
            }
            Event::End(element) if end_is(element.name().as_ref(), end) => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of ODP text")),
            _ => {}
        }
    }
    Ok(runs)
}

fn special_run(value: &str, formatting: &Formatting, link_target: Option<&str>) -> Run {
    Run {
        text: value.to_string(),
        formatting: formatting.clone(),
        link_target: link_target.map(str::to_string),
    }
}

fn parse_list(
    xml: &mut XmlReader<'_>,
    start: &BytesStart<'_>,
    styles: &StyleResolver,
    level: u32,
) -> Result<ListElement> {
    let style_name = attr(start, b"style-name");
    let mut items = Vec::new();
    loop {
        match event(xml, "ODP list")? {
            Event::Start(element) if element_is(xml, &element, TEXT_NS, b"list-item") => {
                parse_list_item(xml, styles, level, style_name.as_deref(), &mut items)?;
            }
            Event::End(element) if end_is(element.name().as_ref(), b"list") => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of ODP list")),
            _ => {}
        }
    }
    Ok(ListElement { items })
}

fn parse_list_item(
    xml: &mut XmlReader<'_>,
    styles: &StyleResolver,
    level: u32,
    style_name: Option<&str>,
    items: &mut Vec<ListItem>,
) -> Result<()> {
    let mut runs = Vec::new();
    let mut nested = Vec::new();
    loop {
        match event(xml, "ODP list item")? {
            Event::Start(element) if element_is(xml, &element, TEXT_NS, b"p") => {
                let mut paragraph = parse_paragraph(xml, &element, styles)?;
                if let Some(last) = paragraph.last_mut() {
                    last.text.push('\n');
                }
                runs.append(&mut paragraph);
            }
            Event::Start(element) if element_is(xml, &element, TEXT_NS, b"list") => {
                nested.extend(parse_list(xml, &element, styles, level + 1)?.items);
            }
            Event::End(element) if end_is(element.name().as_ref(), b"list-item") => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of ODP list item")),
            _ => {}
        }
    }
    if !runs.is_empty() {
        items.push(ListItem {
            level,
            is_ordered: styles.is_ordered_list(style_name, level),
            runs,
        });
    }
    items.extend(nested);
    Ok(())
}

fn parse_table(xml: &mut XmlReader<'_>, styles: &StyleResolver) -> Result<TableElement> {
    let mut rows = Vec::new();
    loop {
        match event(xml, "ODP table")? {
            Event::Start(element) if element_is(xml, &element, TABLE_NS, b"table-row") => {
                let repeats = usize_attr(&element, b"number-rows-repeated").unwrap_or(1);
                let row = parse_table_row(xml, styles)?;
                rows.extend(std::iter::repeat_n(row, repeats));
            }
            Event::End(element) if end_is(element.name().as_ref(), b"table") => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of ODP table")),
            _ => {}
        }
    }
    let width = rows.iter().map(|row| row.cells.len()).max().unwrap_or(0);
    for row in &mut rows {
        row.cells.resize_with(width, TableCell::default);
    }
    Ok(TableElement { rows })
}

fn parse_table_row(xml: &mut XmlReader<'_>, styles: &StyleResolver) -> Result<TableRow> {
    let mut cells = Vec::new();
    loop {
        match event(xml, "ODP table row")? {
            Event::Start(element)
                if element_is(xml, &element, TABLE_NS, b"table-cell")
                    || element_is(xml, &element, TABLE_NS, b"covered-table-cell") =>
            {
                let covered = element_is(xml, &element, TABLE_NS, b"covered-table-cell");
                let repeats = usize_attr(&element, b"number-columns-repeated").unwrap_or(1);
                let spans = usize_attr(&element, b"number-columns-spanned").unwrap_or(1);
                let cell = if covered {
                    skip_element(xml, b"covered-table-cell", "ODP covered cell")?;
                    TableCell {
                        covered: true,
                        ..TableCell::default()
                    }
                } else {
                    parse_table_cell(xml, &element, styles)?
                };
                append_cells(&mut cells, cell, repeats, spans);
            }
            Event::Empty(element)
                if element_is(xml, &element, TABLE_NS, b"table-cell")
                    || element_is(xml, &element, TABLE_NS, b"covered-table-cell") =>
            {
                append_cells(
                    &mut cells,
                    TableCell::default(),
                    usize_attr(&element, b"number-columns-repeated").unwrap_or(1),
                    usize_attr(&element, b"number-columns-spanned").unwrap_or(1),
                );
            }
            Event::End(element) if end_is(element.name().as_ref(), b"table-row") => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of ODP table row")),
            _ => {}
        }
    }
    Ok(TableRow { cells })
}

fn append_cells(cells: &mut Vec<TableCell>, mut cell: TableCell, repeats: usize, spans: usize) {
    cell.column_span = spans.max(1);
    for _ in 0..repeats {
        cells.push(cell.clone());
        cells.extend((1..spans).map(|_| TableCell {
            covered: true,
            ..TableCell::default()
        }));
    }
}

fn parse_table_cell(
    xml: &mut XmlReader<'_>,
    start: &BytesStart<'_>,
    styles: &StyleResolver,
) -> Result<TableCell> {
    let mut runs = Vec::new();
    let mut paragraphs = Vec::new();
    loop {
        match event(xml, "ODP table cell")? {
            Event::Start(element) if element_is(xml, &element, TEXT_NS, b"p") => {
                let mut paragraph = parse_paragraph(xml, &element, styles)?;
                if let Some(last) = paragraph.last_mut() {
                    last.text.push('\n');
                }
                runs.append(&mut paragraph.clone());
                paragraphs.push(crate::Paragraph::plain(paragraph));
            }
            Event::End(element) if end_is(element.name().as_ref(), b"table-cell") => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of ODP table cell")),
            _ => {}
        }
    }
    Ok(TableCell {
        runs,
        paragraphs,
        row_span: usize_attr(start, b"number-rows-spanned").unwrap_or(1),
        column_span: usize_attr(start, b"number-columns-spanned").unwrap_or(1),
        covered: false,
    })
}

fn node_position(element: &BytesStart<'_>) -> ElementPosition {
    let x = attr(element, b"x")
        .as_deref()
        .and_then(parse_length)
        .unwrap_or(0);
    let y = attr(element, b"y")
        .as_deref()
        .and_then(parse_length)
        .unwrap_or(0);
    let transform = attr(element, b"transform")
        .as_deref()
        .map(parse_translate)
        .unwrap_or_default();
    ElementPosition {
        x: x + transform.x,
        y: y + transform.y,
    }
}

fn node_bounds(element: &BytesStart<'_>, position: ElementPosition) -> crate::Bounds {
    crate::Bounds {
        x: position.x,
        y: position.y,
        width: attr(element, b"width")
            .as_deref()
            .and_then(parse_length)
            .unwrap_or(0),
        height: attr(element, b"height")
            .as_deref()
            .and_then(parse_length)
            .unwrap_or(0),
    }
}

fn parse_translate(value: &str) -> ElementPosition {
    let Some(start) = value.find("translate") else {
        return ElementPosition::default();
    };
    let Some(open) = value[start..].find('(') else {
        return ElementPosition::default();
    };
    let content = &value[start + open + 1..];
    let Some(close) = content.find(')') else {
        return ElementPosition::default();
    };
    let mut values = content[..close].split_whitespace();
    ElementPosition {
        x: values.next().and_then(parse_length).unwrap_or(0),
        y: values.next().and_then(parse_length).unwrap_or(0),
    }
}

fn parse_length(value: &str) -> Option<i64> {
    for (suffix, multiplier) in [
        ("cm", 360_000.0),
        ("mm", 36_000.0),
        ("in", 914_400.0),
        ("pt", 12_700.0),
        ("px", 9_525.0),
    ] {
        if let Some(number) = value.strip_suffix(suffix) {
            return number
                .trim()
                .parse::<f64>()
                .ok()
                .map(|number| (number * multiplier).round() as i64);
        }
    }
    value
        .parse::<f64>()
        .ok()
        .map(|number| number.round() as i64)
}

fn add_position(left: ElementPosition, right: ElementPosition) -> ElementPosition {
    ElementPosition {
        x: left.x + right.x,
        y: left.y + right.y,
    }
}

fn usize_attr(element: &BytesStart<'_>, name: &[u8]) -> Option<usize> {
    attr(element, name).and_then(|value| value.parse().ok())
}

#[cfg(test)]
#[path = "../tests/unit/odp.rs"]
mod tests;
