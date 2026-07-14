use crate::{
    ElementPosition, Error, Formatting, ImageReference, ListElement, ListItem, ParserConfig,
    PresentationMetadata, Result, Run, Slide, SlideElement, TableCell, TableElement, TableRow, TextElement,
};
use crate::metadata::{parse_odp_metadata, render_presentation_markdown};
use roxmltree::{Document, Node};
use std::collections::HashMap;
use std::io::Read;
use std::path::Path;

const DRAW_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const TABLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const SVG_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const XLINK_NS: &str = "http://www.w3.org/1999/xlink";
const FO_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const PRESENTATION_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";

pub(crate) struct OdpContainer {
    config: ParserConfig,
    archive: zip::ZipArchive<std::fs::File>,
    content: Vec<u8>,
    styles: Vec<u8>,
    page_count: usize,
    metadata: PresentationMetadata,
}

impl OdpContainer {
    pub(crate) fn open(path: &Path, config: ParserConfig) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        let mut archive = zip::ZipArchive::new(file)?;
        let content = read_archive_file(&mut archive, "content.xml")?;
        let styles = read_archive_file(&mut archive, "styles.xml").unwrap_or_default();
        let meta = read_optional_archive_file(&mut archive, "meta.xml")?;
        let page_count = presentation_page_count(&content)?;
        let metadata = parse_odp_metadata(meta.as_deref())?;

        Ok(Self {
            config,
            archive,
            content,
            styles,
            page_count,
            metadata,
        })
    }

    pub(crate) fn parse_all(&mut self) -> Result<Vec<Slide>> {
        (0..self.page_count)
            .map(|index| self.load_slide(index))
            .collect()
    }

    pub(crate) fn metadata(&self) -> &PresentationMetadata { &self.metadata }

    pub(crate) fn convert_to_md(&mut self) -> Result<String> {
        let slides = self.parse_all()?;
        render_presentation_markdown(&self.metadata, self.config.include_presentation_metadata, slides)
    }

    fn load_slide(&mut self, index: usize) -> Result<Slide> {
        let content = std::str::from_utf8(&self.content)
            .map_err(|_| Error::ParseError("ODP XML is not UTF-8"))?;
        let document = Document::parse(content)?;
        let page = presentation_pages(&document)
            .nth(index)
            .ok_or(Error::SlideNotFound)?;
        let styles = StyleResolver::from_documents(&self.content, &self.styles)?;
        let elements = parse_page(page, &styles)?;
        let speaker_notes = parse_speaker_notes(page, &styles)?;
        let comments = parse_comments(page, &styles)?;
        let images: Vec<ImageReference> = elements
            .iter()
            .filter_map(|element| match element {
                SlideElement::Image(image, _) => Some(image.clone()),
                _ => None,
            })
            .collect();
        let mut image_data = HashMap::new();
        if self.config.extract_images {
            for image in &images {
                if let Ok(data) = read_archive_file(&mut self.archive, &image.target) {
                    image_data.insert(image.id.clone(), data);
                }
            }
        }

        Ok(Slide::new(
            format!("content.xml#page{}", index + 1),
            (index + 1) as u32,
            elements,
            speaker_notes,
            comments,
            images,
            image_data,
            self.config.clone(),
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
        if self.current_index >= self.container.page_count {
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

fn read_optional_archive_file(archive: &mut zip::ZipArchive<std::fs::File>, path: &str) -> Result<Option<Vec<u8>>> {
    let mut file = match archive.by_name(path) {
        Ok(file) => file,
        Err(zip::result::ZipError::FileNotFound) => return Ok(None),
        Err(error) => return Err(error.into()),
    };
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes)?;
    Ok(Some(bytes))
}

fn presentation_page_count(xml: &[u8]) -> Result<usize> {
    let xml = std::str::from_utf8(xml).map_err(|_| Error::ParseError("ODP XML is not UTF-8"))?;
    let document = Document::parse(xml)?;
    Ok(presentation_pages(&document).count())
}

fn presentation_pages<'a>(document: &'a Document<'a>) -> impl Iterator<Item = Node<'a, 'a>> + 'a {
    document
        .descendants()
        .find(|node| is_element(*node, OFFICE_NS, "presentation"))
        .into_iter()
        .flat_map(|presentation| presentation.children())
        .filter(|node| is_element(*node, DRAW_NS, "page"))
}

#[derive(Default, Clone)]
struct PartialFormatting {
    bold: Option<bool>,
    italic: Option<bool>,
    underlined: Option<bool>,
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
        for xml in [content, styles] {
            if xml.is_empty() {
                continue;
            }
            let xml = std::str::from_utf8(xml)
                .map_err(|_| Error::ParseError("ODP style XML is not UTF-8"))?;
            let document = Document::parse(xml)?;
            resolver.collect(&document);
        }
        Ok(resolver)
    }

    fn collect(&mut self, document: &Document<'_>) {
        for node in document
            .descendants()
            .filter(|node| is_element(*node, STYLE_NS, "style"))
        {
            let Some(name) = node.attribute((STYLE_NS, "name")) else {
                continue;
            };
            let formatting = node
                .children()
                .find(|child| is_element(*child, STYLE_NS, "text-properties"))
                .map(parse_formatting_properties)
                .unwrap_or_default();
            self.styles.insert(
                name.to_string(),
                StyleDefinition {
                    parent: node
                        .attribute((STYLE_NS, "parent-style-name"))
                        .map(str::to_string),
                    formatting,
                },
            );
        }

        for list_style in document
            .descendants()
            .filter(|node| is_element(*node, TEXT_NS, "list-style"))
        {
            let Some(name) = list_style.attribute((STYLE_NS, "name")) else {
                continue;
            };
            for level in list_style.children().filter(|child| child.is_element()) {
                let ordered = is_element(level, TEXT_NS, "list-level-style-number");
                if !ordered && !is_element(level, TEXT_NS, "list-level-style-bullet") {
                    continue;
                }
                let level_number = level
                    .attribute((TEXT_NS, "level"))
                    .and_then(|value| value.parse().ok())
                    .unwrap_or(1);
                self.list_levels
                    .insert((name.to_string(), level_number), ordered);
            }
        }
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

fn parse_formatting_properties(node: Node<'_, '_>) -> PartialFormatting {
    let mut formatting = PartialFormatting::default();
    formatting.bold = node.attribute((FO_NS, "font-weight")).map(is_enabled);
    formatting.italic = node
        .attribute((FO_NS, "font-style"))
        .map(|value| value.eq_ignore_ascii_case("italic"));
    formatting.underlined = node
        .attribute((STYLE_NS, "text-underline-style"))
        .map(|value| !value.eq_ignore_ascii_case("none"));
    formatting.lang = node.attribute((FO_NS, "language")).map(str::to_string);
    formatting
}

fn is_enabled(value: &str) -> bool {
    matches!(value, "bold" | "700" | "800" | "900")
}

fn parse_page(page: Node<'_, '_>, styles: &StyleResolver) -> Result<Vec<SlideElement>> {
    let mut elements = Vec::new();
    for child in page.children().filter(|node| node.is_element()) {
        parse_node(child, ElementPosition::default(), styles, &mut elements)?;
    }
    Ok(elements)
}

fn parse_speaker_notes(page: Node<'_, '_>, styles: &StyleResolver) -> Result<Vec<TextElement>> {
    let mut elements = Vec::new();
    for notes in page
        .children()
        .filter(|node| is_element(*node, PRESENTATION_NS, "notes"))
    {
        for child in notes.children().filter(|node| node.is_element()) {
            parse_node(child, ElementPosition::default(), styles, &mut elements)?;
        }
    }
    Ok(elements
        .into_iter()
        .filter_map(|element| match element {
            SlideElement::Text(text, _) => Some(text),
            _ => None,
        })
        .collect())
}

fn parse_comments(page: Node<'_, '_>, styles: &StyleResolver) -> Result<Vec<TextElement>> {
    let mut elements = Vec::new();
    for annotation in page
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "annotation")
    {
        parse_text_container(
            annotation,
            ElementPosition::default(),
            styles,
            &mut elements,
        )?;
    }
    Ok(elements
        .into_iter()
        .filter_map(|element| match element {
            SlideElement::Text(text, _) => Some(text),
            _ => None,
        })
        .collect())
}

fn parse_node(
    node: Node<'_, '_>,
    parent_position: ElementPosition,
    styles: &StyleResolver,
    elements: &mut Vec<SlideElement>,
) -> Result<()> {
    let position = add_position(parent_position, node_position(node));
    if is_element(node, DRAW_NS, "g") {
        for child in node.children().filter(|child| child.is_element()) {
            parse_node(child, position, styles, elements)?;
        }
    } else if is_element(node, DRAW_NS, "frame") {
        for child in node.children().filter(|child| child.is_element()) {
            if is_element(child, DRAW_NS, "image") {
                if let Some(reference) = parse_image(child) {
                    elements.push(SlideElement::Image(reference, position));
                }
            } else if is_element(child, DRAW_NS, "text-box") {
                parse_text_container(child, position, styles, elements)?;
            } else if is_element(child, TABLE_NS, "table") {
                elements.push(SlideElement::Table(parse_table(child, styles)?, position));
            }
        }
    } else if is_element(node, DRAW_NS, "custom-shape") {
        parse_text_container(node, position, styles, elements)?;
    } else if is_element(node, TABLE_NS, "table") {
        elements.push(SlideElement::Table(parse_table(node, styles)?, position));
    }
    Ok(())
}

fn parse_text_container(
    node: Node<'_, '_>,
    position: ElementPosition,
    styles: &StyleResolver,
    elements: &mut Vec<SlideElement>,
) -> Result<()> {
    let paragraphs: Vec<_> = node
        .children()
        .filter(|child| is_element(*child, TEXT_NS, "p") || is_element(*child, TEXT_NS, "h"))
        .collect();
    if !paragraphs.is_empty() {
        let mut runs = Vec::new();
        for paragraph in paragraphs {
            let mut paragraph_runs = parse_paragraph(paragraph, styles);
            if let Some(last) = paragraph_runs.last_mut() {
                last.text.push('\n');
            }
            runs.append(&mut paragraph_runs);
        }
        if !runs.is_empty() {
            elements.push(SlideElement::Text(TextElement { runs }, position));
        }
    }
    for list in node
        .children()
        .filter(|child| is_element(*child, TEXT_NS, "list"))
    {
        elements.push(SlideElement::List(parse_list(list, styles, 0), position));
    }
    Ok(())
}

fn parse_image(node: Node<'_, '_>) -> Option<ImageReference> {
    let target = node.attribute((XLINK_NS, "href"))?.to_string();
    Some(ImageReference {
        id: target.clone(),
        target,
    })
}

fn parse_paragraph(node: Node<'_, '_>, styles: &StyleResolver) -> Vec<Run> {
    let formatting = styles.formatting(
        node.attribute((TEXT_NS, "style-name")),
        Formatting::default(),
    );
    let mut runs = Vec::new();
    collect_runs(node, formatting, styles, None, &mut runs);
    runs
}

fn collect_runs(
    node: Node<'_, '_>,
    formatting: Formatting,
    styles: &StyleResolver,
    link_target: Option<&str>,
    runs: &mut Vec<Run>,
) {
    for child in node.children() {
        if child.is_text() {
            if let Some(text) = child.text().filter(|text| !text.is_empty()) {
                runs.push(Run {
                    text: text.to_string(),
                    formatting: formatting.clone(),
                    link_target: link_target.map(str::to_string),
                });
            }
        } else if is_element(child, TEXT_NS, "span") {
            let next =
                styles.formatting(child.attribute((TEXT_NS, "style-name")), formatting.clone());
            collect_runs(child, next, styles, link_target, runs);
        } else if is_element(child, TEXT_NS, "a") {
            let target = child.attribute((XLINK_NS, "href")).or(link_target);
            collect_runs(child, formatting.clone(), styles, target, runs);
        } else if is_element(child, TEXT_NS, "line-break") {
            runs.push(Run {
                text: "\n".to_string(),
                formatting: formatting.clone(),
                link_target: link_target.map(str::to_string),
            });
        } else if is_element(child, TEXT_NS, "tab") {
            runs.push(Run {
                text: "\t".to_string(),
                formatting: formatting.clone(),
                link_target: link_target.map(str::to_string),
            });
        } else if is_element(child, TEXT_NS, "s") {
            let count = child
                .attribute((TEXT_NS, "c"))
                .and_then(|value| value.parse().ok())
                .unwrap_or(1);
            runs.push(Run {
                text: " ".repeat(count),
                formatting: formatting.clone(),
                link_target: link_target.map(str::to_string),
            });
        }
    }
}

fn parse_list(node: Node<'_, '_>, styles: &StyleResolver, level: u32) -> ListElement {
    let style_name = node.attribute((TEXT_NS, "style-name"));
    let mut items = Vec::new();
    for item in node
        .children()
        .filter(|child| is_element(*child, TEXT_NS, "list-item"))
    {
        let mut runs = Vec::new();
        for paragraph in item
            .children()
            .filter(|child| is_element(*child, TEXT_NS, "p"))
        {
            let mut paragraph_runs = parse_paragraph(paragraph, styles);
            if let Some(last) = paragraph_runs.last_mut() {
                last.text.push('\n');
            }
            runs.append(&mut paragraph_runs);
        }
        if !runs.is_empty() {
            items.push(ListItem {
                level,
                is_ordered: styles.is_ordered_list(style_name, level),
                runs,
            });
        }
        for nested in item
            .children()
            .filter(|child| is_element(*child, TEXT_NS, "list"))
        {
            items.extend(parse_list(nested, styles, level + 1).items);
        }
    }
    ListElement { items }
}

fn parse_table(node: Node<'_, '_>, styles: &StyleResolver) -> Result<TableElement> {
    let mut rows = Vec::new();
    for row in node
        .children()
        .filter(|child| is_element(*child, TABLE_NS, "table-row"))
    {
        let parsed = parse_table_row(row, styles);
        let repeats = attribute_usize(row, TABLE_NS, "number-rows-repeated").unwrap_or(1);
        for _ in 0..repeats {
            rows.push(parsed.clone());
        }
    }
    let width = rows
        .iter()
        .map(|row: &TableRow| row.cells.len())
        .max()
        .unwrap_or(0);
    for row in &mut rows {
        while row.cells.len() < width {
            row.cells.push(TableCell { runs: Vec::new() });
        }
    }
    Ok(TableElement { rows })
}

fn parse_table_row(node: Node<'_, '_>, styles: &StyleResolver) -> TableRow {
    let mut cells = Vec::new();
    for cell in node.children().filter(|child| {
        is_element(*child, TABLE_NS, "table-cell")
            || is_element(*child, TABLE_NS, "covered-table-cell")
    }) {
        let parsed = if is_element(cell, TABLE_NS, "covered-table-cell") {
            TableCell { runs: Vec::new() }
        } else {
            parse_table_cell(cell, styles)
        };
        let repeats = attribute_usize(cell, TABLE_NS, "number-columns-repeated").unwrap_or(1);
        let spans = attribute_usize(cell, TABLE_NS, "number-columns-spanned").unwrap_or(1);
        for _ in 0..repeats {
            cells.push(parsed.clone());
            for _ in 1..spans {
                cells.push(TableCell { runs: Vec::new() });
            }
        }
    }
    TableRow { cells }
}

fn parse_table_cell(node: Node<'_, '_>, styles: &StyleResolver) -> TableCell {
    let mut runs = Vec::new();
    for paragraph in node
        .children()
        .filter(|child| is_element(*child, TEXT_NS, "p"))
    {
        let mut paragraph_runs = parse_paragraph(paragraph, styles);
        if let Some(last) = paragraph_runs.last_mut() {
            last.text.push('\n');
        }
        runs.append(&mut paragraph_runs);
    }
    TableCell { runs }
}

fn node_position(node: Node<'_, '_>) -> ElementPosition {
    let x = node
        .attribute((SVG_NS, "x"))
        .and_then(parse_length)
        .unwrap_or(0);
    let y = node
        .attribute((SVG_NS, "y"))
        .and_then(parse_length)
        .unwrap_or(0);
    let transform = node
        .attribute((DRAW_NS, "transform"))
        .map(parse_translate)
        .unwrap_or_default();
    ElementPosition {
        x: x + transform.x,
        y: y + transform.y,
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
    let units = [
        ("cm", 360_000.0),
        ("mm", 36_000.0),
        ("in", 914_400.0),
        ("pt", 12_700.0),
        ("px", 9_525.0),
    ];
    for (suffix, multiplier) in units {
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

fn attribute_usize(node: Node<'_, '_>, namespace: &str, name: &str) -> Option<usize> {
    node.attribute((namespace, name))
        .and_then(|value| value.parse().ok())
}

fn is_element(node: Node<'_, '_>, namespace: &str, name: &str) -> bool {
    node.is_element()
        && node.tag_name().namespace() == Some(namespace)
        && node.tag_name().name() == name
}

#[cfg(test)]
#[path = "../tests/unit/odp.rs"]
mod tests;
