use crate::constants::{A_NAMESPACE, P_NAMESPACE};
use crate::types::{SlideElement, TableCell, TableElement, TableRow, TextElement};
use crate::xml::{
    XmlReader, attr, element_is, end_is, event, reader, reference, skip_element, text,
};
use crate::{
    Bounds, DiagnosticSeverity, ElementPosition, Error, Formatting, ImageBlock, ImageReference,
    ListElement, ListInfo, ListItem, ListKind, Paragraph, ParagraphAlignment, ParseDiagnostic,
    Result, Run, SemanticTable, SemanticTableCell, SemanticTableRow, SlideBlock, SlideBlockContent,
    TextBlock, TextRole, UnsupportedBlock,
};
use quick_xml::events::{BytesStart, Event};
use std::collections::HashMap;

type ParsedContent = TextBlock;

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

    fn apply_bounds(self, bounds: Bounds) -> Bounds {
        let position = self.apply(ElementPosition {
            x: bounds.x,
            y: bounds.y,
        });
        Bounds {
            x: position.x,
            y: position.y,
            width: (bounds.width as f64 * self.scale_x.abs()).round() as i64,
            height: (bounds.height as f64 * self.scale_y.abs()).round() as i64,
        }
    }

    fn then(self, next: Self) -> Self {
        Self {
            scale_x: self.scale_x * next.scale_x,
            scale_y: self.scale_y * next.scale_y,
            translate_x: self.scale_x * next.translate_x + self.translate_x,
            translate_y: self.scale_y * next.translate_y + self.translate_y,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct PlaceholderKey {
    kind: Option<String>,
    idx: Option<String>,
}

#[derive(Debug, Clone, Default)]
pub struct InheritedPositions {
    positions: HashMap<PlaceholderKey, ElementPosition>,
    list_styles: HashMap<PlaceholderKey, HashMap<u32, ListKind>>,
}

impl InheritedPositions {
    fn resolve(&self, key: &PlaceholderKey) -> Option<ElementPosition> {
        self.positions.get(key).copied().or_else(|| {
            key.idx
                .as_ref()
                .and_then(|idx| {
                    self.positions
                        .iter()
                        .find(|(candidate, _)| candidate.idx.as_deref() == Some(idx))
                })
                .or_else(|| {
                    key.kind.as_ref().and_then(|kind| {
                        self.positions
                            .iter()
                            .find(|(candidate, _)| candidate.kind.as_deref() == Some(kind))
                    })
                })
                .map(|(_, position)| *position)
        })
    }

    fn resolve_list_kind(&self, key: &PlaceholderKey, level: u32) -> Option<ListKind> {
        self.list_styles
            .get(key)
            .or_else(|| {
                key.idx.as_ref().and_then(|idx| {
                    self.list_styles
                        .iter()
                        .find(|(candidate, _)| candidate.idx.as_ref() == Some(idx))
                        .map(|(_, styles)| styles)
                })
            })
            .or_else(|| {
                key.kind.as_ref().and_then(|kind| {
                    self.list_styles
                        .iter()
                        .find(|(candidate, _)| candidate.kind.as_ref() == Some(kind))
                        .map(|(_, styles)| styles)
                })
            })
            .and_then(|styles| styles.get(&level).or_else(|| styles.get(&0)))
            .cloned()
    }
}

#[derive(Default)]
struct PositionData {
    x: Option<i64>,
    y: Option<i64>,
    width: Option<i64>,
    height: Option<i64>,
    placeholder: Option<PlaceholderKey>,
    fallback_text: String,
}

impl PositionData {
    fn observe_off(&mut self, element: &BytesStart<'_>) {
        if self.x.is_none() {
            self.x = attr(element, b"x").and_then(|value| value.parse().ok());
            self.y = attr(element, b"y").and_then(|value| value.parse().ok());
        }
    }

    fn observe_placeholder(&mut self, element: &BytesStart<'_>) {
        self.placeholder = Some(PlaceholderKey {
            kind: attr(element, b"type"),
            idx: attr(element, b"idx"),
        });
    }

    fn observe_ext(&mut self, element: &BytesStart<'_>) {
        if self.width.is_none() {
            self.width = attr(element, b"cx").and_then(|value| value.parse().ok());
            self.height = attr(element, b"cy").and_then(|value| value.parse().ok());
        }
    }

    fn raw(&self) -> Option<ElementPosition> {
        Some(ElementPosition {
            x: self.x?,
            y: self.y?,
        })
    }

    fn effective(
        &self,
        transform: CoordinateTransform,
        inherited: &InheritedPositions,
    ) -> ElementPosition {
        self.raw()
            .map(|position| transform.apply(position))
            .or_else(|| {
                self.placeholder
                    .as_ref()
                    .and_then(|key| inherited.resolve(key))
            })
            .unwrap_or_default()
    }

    fn effective_bounds(
        &self,
        transform: CoordinateTransform,
        inherited: &InheritedPositions,
    ) -> Bounds {
        if let Some(position) = self.raw() {
            transform.apply_bounds(Bounds {
                x: position.x,
                y: position.y,
                width: self.width.unwrap_or(0),
                height: self.height.unwrap_or(0),
            })
        } else {
            let position = self.effective(transform, inherited);
            Bounds {
                x: position.x,
                y: position.y,
                width: (self.width.unwrap_or(0) as f64 * transform.scale_x.abs()).round() as i64,
                height: (self.height.unwrap_or(0) as f64 * transform.scale_y.abs()).round() as i64,
            }
        }
    }
}

struct ShapeData {
    content: Option<ParsedContent>,
    position: PositionData,
}

pub(crate) struct ParsedSlideDocument {
    pub elements: Vec<SlideElement>,
    pub blocks: Vec<SlideBlock>,
    pub diagnostics: Vec<ParseDiagnostic>,
}

pub fn parse_slide_xml(xml_data: &[u8]) -> Result<Vec<SlideElement>> {
    parse_slide_xml_with_hyperlinks(xml_data, &InheritedPositions::default(), &HashMap::new())
}

#[cfg(test)]
pub(crate) fn parse_speaker_notes_xml(xml_data: &[u8]) -> Result<Vec<TextElement>> {
    parse_speaker_notes_xml_with_hyperlinks(xml_data, &HashMap::new())
}

pub(crate) fn parse_speaker_notes_xml_with_hyperlinks(
    xml_data: &[u8],
    hyperlinks: &HashMap<String, String>,
) -> Result<Vec<TextElement>> {
    let mut xml = reader(xml_data);
    loop {
        match event(&mut xml, "PPTX notes")? {
            Event::Start(element) if element_is(&xml, &element, P_NAMESPACE, b"spTree") => {
                return parse_notes_tree(&mut xml, hyperlinks);
            }
            Event::Eof => return Err(Error::ParseError("PPTX notes shape tree not found")),
            _ => {}
        }
    }
}

fn parse_notes_tree(
    xml: &mut XmlReader<'_>,
    hyperlinks: &HashMap<String, String>,
) -> Result<Vec<TextElement>> {
    let mut notes = Vec::new();
    loop {
        match event(xml, "PPTX notes")? {
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"sp") => {
                let shape = parse_shape(xml, hyperlinks)?;
                if shape
                    .position
                    .placeholder
                    .as_ref()
                    .and_then(|key| key.kind.as_deref())
                    == Some("body")
                    && let Some(content) = shape.content
                {
                    let text = content_to_text(content);
                    if !text.runs.is_empty() {
                        notes.push(text);
                    }
                }
            }
            Event::Start(element) => {
                let end = element.name().as_ref().to_vec();
                skip_element(xml, &end, "PPTX notes")?;
            }
            Event::End(element) if end_is(element.name().as_ref(), b"spTree") => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of PPTX notes")),
            _ => {}
        }
    }
    Ok(notes)
}

pub(crate) fn parse_comments_xml_with_hyperlinks(
    xml_data: &[u8],
    hyperlinks: &HashMap<String, String>,
) -> Result<Vec<TextElement>> {
    let mut xml = reader(xml_data);
    let mut comments = Vec::new();
    let mut legacy = Vec::new();
    loop {
        match event(&mut xml, "PPTX comments")? {
            Event::Start(element) if crate::xml::local(element.name().as_ref()) == b"txBody" => {
                let content = parse_text_body(&mut xml, true, hyperlinks)?;
                let text = content_to_text(content);
                if !text.runs.is_empty() {
                    comments.push(text);
                }
            }
            Event::Start(element) if element_is(&xml, &element, P_NAMESPACE, b"text") => {
                let value = read_simple_text(&mut xml, b"text", "PPTX comments")?;
                if !value.is_empty() {
                    legacy.push(TextElement {
                        runs: vec![Run {
                            text: format!("{value}\n"),
                            formatting: Formatting::default(),
                            link_target: None,
                        }],
                    });
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(if comments.is_empty() {
        legacy
    } else {
        comments
    })
}

pub fn parse_slide_xml_with_inherited_positions(
    xml_data: &[u8],
    inherited_positions: &InheritedPositions,
) -> Result<Vec<SlideElement>> {
    parse_slide_xml_with_hyperlinks(xml_data, inherited_positions, &HashMap::new())
}

pub fn parse_slide_xml_with_hyperlinks(
    xml_data: &[u8],
    inherited_positions: &InheritedPositions,
    hyperlinks: &HashMap<String, String>,
) -> Result<Vec<SlideElement>> {
    let mut xml = reader(xml_data);
    let mut in_common_slide = false;
    loop {
        match event(&mut xml, "PPTX slide")? {
            Event::Start(element) if element_is(&xml, &element, P_NAMESPACE, b"cSld") => {
                in_common_slide = true;
            }
            Event::Start(element)
                if in_common_slide && element_is(&xml, &element, P_NAMESPACE, b"spTree") =>
            {
                return parse_shape_tree(
                    &mut xml,
                    CoordinateTransform::identity(),
                    inherited_positions,
                    hyperlinks,
                    b"spTree",
                );
            }
            Event::Empty(element)
                if in_common_slide && element_is(&xml, &element, P_NAMESPACE, b"spTree") =>
            {
                return Ok(Vec::new());
            }
            Event::End(element) if end_is(element.name().as_ref(), b"cSld") => {
                in_common_slide = false;
            }
            Event::Eof => return Err(Error::ParseError("PPTX slide shape tree not found")),
            _ => {}
        }
    }
}

pub(crate) fn parse_slide_document_with_hyperlinks(
    xml_data: &[u8],
    inherited_positions: &InheritedPositions,
    hyperlinks: &HashMap<String, String>,
) -> Result<ParsedSlideDocument> {
    let mut xml = reader(xml_data);
    let mut in_common_slide = false;
    let mut source_order = 0usize;
    loop {
        match event(&mut xml, "PPTX slide")? {
            Event::Start(element) if element_is(&xml, &element, P_NAMESPACE, b"cSld") => {
                in_common_slide = true;
            }
            Event::Start(element)
                if in_common_slide && element_is(&xml, &element, P_NAMESPACE, b"spTree") =>
            {
                return parse_semantic_shape_tree(
                    &mut xml,
                    CoordinateTransform::identity(),
                    inherited_positions,
                    hyperlinks,
                    b"spTree",
                    &mut source_order,
                );
            }
            Event::Empty(element)
                if in_common_slide && element_is(&xml, &element, P_NAMESPACE, b"spTree") =>
            {
                return Ok(ParsedSlideDocument {
                    elements: Vec::new(),
                    blocks: Vec::new(),
                    diagnostics: Vec::new(),
                });
            }
            Event::End(element) if end_is(element.name().as_ref(), b"cSld") => {
                in_common_slide = false;
            }
            Event::Eof => return Err(Error::ParseError("PPTX slide shape tree not found")),
            _ => {}
        }
    }
}

fn parse_semantic_shape_tree(
    xml: &mut XmlReader<'_>,
    transform: CoordinateTransform,
    inherited: &InheritedPositions,
    hyperlinks: &HashMap<String, String>,
    end: &[u8],
    source_order: &mut usize,
) -> Result<ParsedSlideDocument> {
    let mut parsed = ParsedSlideDocument {
        elements: Vec::new(),
        blocks: Vec::new(),
        diagnostics: Vec::new(),
    };
    loop {
        match event(xml, "PPTX slide")? {
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"sp") => {
                let mut shape = parse_shape(xml, hyperlinks)?;
                let position = shape.position.effective(transform, inherited);
                let bounds = shape.position.effective_bounds(transform, inherited);
                if let Some(mut content) = shape.content.take() {
                    apply_inherited_list_styles(&mut content, &shape.position, inherited);
                    content.role = placeholder_role(shape.position.placeholder.as_ref());
                    parsed
                        .elements
                        .extend(content_to_elements(content.clone(), position));
                    push_semantic_block(
                        &mut parsed,
                        source_order,
                        bounds,
                        SlideBlockContent::Text(content),
                    );
                } else {
                    push_unsupported(&mut parsed, source_order, bounds, "shape", None);
                }
            }
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"graphicFrame") => {
                let (table, position) = parse_graphic_frame(xml, hyperlinks)?;
                let bounds = position.effective_bounds(transform, inherited);
                if let Some(table) = table {
                    parsed.elements.push(SlideElement::Table(
                        table.clone(),
                        position.effective(transform, inherited),
                    ));
                    push_semantic_block(
                        &mut parsed,
                        source_order,
                        bounds,
                        SlideBlockContent::Table(legacy_table_to_semantic(&table)),
                    );
                } else {
                    let fallback_text = (!position.fallback_text.trim().is_empty())
                        .then(|| position.fallback_text.trim().to_string());
                    push_unsupported(
                        &mut parsed,
                        source_order,
                        bounds,
                        "graphicFrame",
                        fallback_text,
                    );
                }
            }
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"pic") => {
                let (image, position, alt_text) = parse_picture(xml)?;
                parsed.elements.push(SlideElement::Image(
                    image.clone(),
                    position.effective(transform, inherited),
                ));
                push_semantic_block(
                    &mut parsed,
                    source_order,
                    position.effective_bounds(transform, inherited),
                    SlideBlockContent::Image(ImageBlock {
                        reference: image,
                        alt_text,
                        mime_type: None,
                    }),
                );
            }
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"grpSp") => {
                merge_parsed(
                    &mut parsed,
                    parse_semantic_group(xml, transform, inherited, hyperlinks, source_order)?,
                );
            }
            Event::Start(element) => {
                let local_name = crate::xml::local(element.name().as_ref()).to_vec();
                let is_presentation = {
                    let (namespace, _) = xml.resolver().resolve_element(element.name());
                    matches!(namespace, quick_xml::name::ResolveResult::Bound(value) if value.as_ref() == P_NAMESPACE.as_bytes())
                };
                let element_end = element.name().as_ref().to_vec();
                skip_element(xml, &element_end, "PPTX slide")?;
                if is_presentation && !matches!(local_name.as_slice(), b"nvGrpSpPr" | b"grpSpPr") {
                    push_unsupported(
                        &mut parsed,
                        source_order,
                        Bounds::default(),
                        &String::from_utf8_lossy(&local_name),
                        None,
                    );
                }
            }
            Event::End(element) if end_is(element.name().as_ref(), end) => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of PPTX shape tree")),
            _ => {}
        }
    }
    Ok(parsed)
}

fn parse_semantic_group(
    xml: &mut XmlReader<'_>,
    parent: CoordinateTransform,
    inherited: &InheritedPositions,
    hyperlinks: &HashMap<String, String>,
    source_order: &mut usize,
) -> Result<ParsedSlideDocument> {
    let mut parsed = ParsedSlideDocument {
        elements: Vec::new(),
        blocks: Vec::new(),
        diagnostics: Vec::new(),
    };
    let mut transform = GroupTransformData::default();
    loop {
        match event(xml, "PPTX group")? {
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"grpSpPr") => {
                parse_group_properties(xml, &mut transform)?;
            }
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"sp") => {
                let combined = parent.then(transform.finish());
                let mut shape = parse_shape(xml, hyperlinks)?;
                let position = shape.position.effective(combined, inherited);
                let bounds = shape.position.effective_bounds(combined, inherited);
                if let Some(mut content) = shape.content.take() {
                    apply_inherited_list_styles(&mut content, &shape.position, inherited);
                    content.role = placeholder_role(shape.position.placeholder.as_ref());
                    parsed
                        .elements
                        .extend(content_to_elements(content.clone(), position));
                    push_semantic_block(
                        &mut parsed,
                        source_order,
                        bounds,
                        SlideBlockContent::Text(content),
                    );
                } else {
                    push_unsupported(&mut parsed, source_order, bounds, "shape", None);
                }
            }
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"graphicFrame") => {
                let combined = parent.then(transform.finish());
                let (table, position) = parse_graphic_frame(xml, hyperlinks)?;
                let bounds = position.effective_bounds(combined, inherited);
                if let Some(table) = table {
                    parsed.elements.push(SlideElement::Table(
                        table.clone(),
                        position.effective(combined, inherited),
                    ));
                    push_semantic_block(
                        &mut parsed,
                        source_order,
                        bounds,
                        SlideBlockContent::Table(legacy_table_to_semantic(&table)),
                    );
                } else {
                    let fallback_text = (!position.fallback_text.trim().is_empty())
                        .then(|| position.fallback_text.trim().to_string());
                    push_unsupported(
                        &mut parsed,
                        source_order,
                        bounds,
                        "graphicFrame",
                        fallback_text,
                    );
                }
            }
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"pic") => {
                let combined = parent.then(transform.finish());
                let (image, position, alt_text) = parse_picture(xml)?;
                parsed.elements.push(SlideElement::Image(
                    image.clone(),
                    position.effective(combined, inherited),
                ));
                push_semantic_block(
                    &mut parsed,
                    source_order,
                    position.effective_bounds(combined, inherited),
                    SlideBlockContent::Image(ImageBlock {
                        reference: image,
                        alt_text,
                        mime_type: None,
                    }),
                );
            }
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"grpSp") => {
                merge_parsed(
                    &mut parsed,
                    parse_semantic_group(
                        xml,
                        parent.then(transform.finish()),
                        inherited,
                        hyperlinks,
                        source_order,
                    )?,
                );
            }
            Event::Start(element) => {
                let name = crate::xml::local(element.name().as_ref()).to_vec();
                let end = element.name().as_ref().to_vec();
                skip_element(xml, &end, "PPTX group")?;
                if !matches!(name.as_slice(), b"nvGrpSpPr") {
                    push_unsupported(
                        &mut parsed,
                        source_order,
                        Bounds::default(),
                        &String::from_utf8_lossy(&name),
                        None,
                    );
                }
            }
            Event::End(element) if end_is(element.name().as_ref(), b"grpSp") => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of PPTX group")),
            _ => {}
        }
    }
    Ok(parsed)
}

fn merge_parsed(target: &mut ParsedSlideDocument, mut source: ParsedSlideDocument) {
    target.elements.append(&mut source.elements);
    target.blocks.append(&mut source.blocks);
    target.diagnostics.append(&mut source.diagnostics);
}

fn push_semantic_block(
    parsed: &mut ParsedSlideDocument,
    source_order: &mut usize,
    bounds: Bounds,
    content: SlideBlockContent,
) {
    parsed.blocks.push(SlideBlock {
        bounds,
        source_order: *source_order,
        content,
    });
    *source_order += 1;
}

fn push_unsupported(
    parsed: &mut ParsedSlideDocument,
    source_order: &mut usize,
    bounds: Bounds,
    kind: &str,
    fallback_text: Option<String>,
) {
    parsed.elements.push(SlideElement::Unknown);
    parsed.diagnostics.push(ParseDiagnostic {
        severity: DiagnosticSeverity::Warning,
        message: format!("Unsupported PPTX slide element: {kind}"),
        source: None,
    });
    push_semantic_block(
        parsed,
        source_order,
        bounds,
        SlideBlockContent::Unsupported(UnsupportedBlock {
            kind: kind.to_string(),
            fallback_text,
        }),
    );
}

fn placeholder_role(placeholder: Option<&PlaceholderKey>) -> TextRole {
    match placeholder.and_then(|placeholder| placeholder.kind.as_deref()) {
        Some("title") | Some("ctrTitle") => TextRole::Title,
        Some("subTitle") => TextRole::Subtitle,
        Some("body") | Some("obj") => TextRole::Body,
        Some("caption") => TextRole::Caption,
        None if placeholder.is_some_and(|placeholder| placeholder.idx.is_some()) => TextRole::Body,
        _ => TextRole::Other,
    }
}

fn apply_inherited_list_styles(
    content: &mut TextBlock,
    position: &PositionData,
    inherited: &InheritedPositions,
) {
    let Some(placeholder) = position.placeholder.as_ref() else {
        return;
    };
    for paragraph in &mut content.paragraphs {
        if paragraph.list_explicit {
            continue;
        }
        let level = paragraph.list.as_ref().map(|list| list.level).unwrap_or(0);
        if let Some(kind) = inherited.resolve_list_kind(placeholder, level) {
            paragraph.list = Some(ListInfo { level, kind });
        }
    }
}

fn legacy_table_to_semantic(table: &TableElement) -> SemanticTable {
    SemanticTable {
        rows: table
            .rows
            .iter()
            .map(|row| SemanticTableRow {
                cells: row
                    .cells
                    .iter()
                    .map(|cell| SemanticTableCell {
                        paragraphs: if cell.paragraphs.is_empty() {
                            vec![Paragraph::plain(cell.runs.clone())]
                        } else {
                            cell.paragraphs.clone()
                        },
                        row_span: cell.row_span.max(1),
                        column_span: cell.column_span.max(1),
                        covered: cell.covered,
                    })
                    .collect(),
            })
            .collect(),
    }
}

fn parse_shape_tree(
    xml: &mut XmlReader<'_>,
    transform: CoordinateTransform,
    inherited: &InheritedPositions,
    hyperlinks: &HashMap<String, String>,
    end: &[u8],
) -> Result<Vec<SlideElement>> {
    let mut elements = Vec::new();
    loop {
        match event(xml, "PPTX slide")? {
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"sp") => {
                let mut shape = parse_shape(xml, hyperlinks)?;
                let position = shape.position.effective(transform, inherited);
                if let Some(content) = shape.content.as_mut() {
                    apply_inherited_list_styles(content, &shape.position, inherited);
                }
                elements.extend(content_to_elements(
                    shape
                        .content
                        .ok_or(Error::ParseError("PPTX shape has no text body"))?,
                    position,
                ));
            }
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"graphicFrame") => {
                let (table, position) = parse_graphic_frame(xml, hyperlinks)?;
                if let Some(table) = table {
                    elements.push(SlideElement::Table(
                        table,
                        position.effective(transform, inherited),
                    ));
                }
            }
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"pic") => {
                let (image, position, _) = parse_picture(xml)?;
                elements.push(SlideElement::Image(
                    image,
                    position.effective(transform, inherited),
                ));
            }
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"grpSp") => {
                elements.extend(parse_group(xml, transform, inherited, hyperlinks)?);
            }
            Event::Start(element) => {
                let is_presentation = {
                    let (namespace, _) = xml.resolver().resolve_element(element.name());
                    matches!(namespace, quick_xml::name::ResolveResult::Bound(value) if value.as_ref() == P_NAMESPACE.as_bytes())
                };
                let element_end = element.name().as_ref().to_vec();
                skip_element(xml, &element_end, "PPTX slide")?;
                if is_presentation {
                    elements.push(SlideElement::Unknown);
                }
            }
            Event::Empty(element) => {
                let (namespace, _) = xml.resolver().resolve_element(element.name());
                if matches!(namespace, quick_xml::name::ResolveResult::Bound(value) if value.as_ref() == P_NAMESPACE.as_bytes())
                {
                    elements.push(SlideElement::Unknown);
                }
            }
            Event::End(element) if end_is(element.name().as_ref(), end) => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of PPTX shape tree")),
            _ => {}
        }
    }
    Ok(elements)
}

fn parse_group(
    xml: &mut XmlReader<'_>,
    parent: CoordinateTransform,
    inherited: &InheritedPositions,
    hyperlinks: &HashMap<String, String>,
) -> Result<Vec<SlideElement>> {
    let mut elements = Vec::new();
    let mut transform = GroupTransformData::default();
    loop {
        match event(xml, "PPTX group")? {
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"grpSpPr") => {
                parse_group_properties(xml, &mut transform)?;
                elements.push(SlideElement::Unknown);
            }
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"sp") => {
                let combined = parent.then(transform.finish());
                let mut shape = parse_shape(xml, hyperlinks)?;
                let position = shape.position.effective(combined, inherited);
                if let Some(content) = shape.content.as_mut() {
                    apply_inherited_list_styles(content, &shape.position, inherited);
                }
                elements.extend(content_to_elements(
                    shape
                        .content
                        .ok_or(Error::ParseError("PPTX shape has no text body"))?,
                    position,
                ));
            }
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"graphicFrame") => {
                let combined = parent.then(transform.finish());
                let (table, position) = parse_graphic_frame(xml, hyperlinks)?;
                if let Some(table) = table {
                    elements.push(SlideElement::Table(
                        table,
                        position.effective(combined, inherited),
                    ));
                }
            }
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"pic") => {
                let combined = parent.then(transform.finish());
                let (image, position, _) = parse_picture(xml)?;
                elements.push(SlideElement::Image(
                    image,
                    position.effective(combined, inherited),
                ));
            }
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"grpSp") => {
                elements.extend(parse_group(
                    xml,
                    parent.then(transform.finish()),
                    inherited,
                    hyperlinks,
                )?);
            }
            Event::Start(element) => {
                let is_presentation = {
                    let (namespace, _) = xml.resolver().resolve_element(element.name());
                    matches!(namespace, quick_xml::name::ResolveResult::Bound(value) if value.as_ref() == P_NAMESPACE.as_bytes())
                };
                let end = element.name().as_ref().to_vec();
                skip_element(xml, &end, "PPTX group")?;
                if is_presentation {
                    elements.push(SlideElement::Unknown);
                }
            }
            Event::End(element) if end_is(element.name().as_ref(), b"grpSp") => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of PPTX group")),
            _ => {}
        }
    }
    Ok(elements)
}

fn parse_shape(xml: &mut XmlReader<'_>, hyperlinks: &HashMap<String, String>) -> Result<ShapeData> {
    let mut position = PositionData::default();
    let mut content = None;
    loop {
        match event(xml, "PPTX shape")? {
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, A_NAMESPACE, b"off") =>
            {
                position.observe_off(&element);
            }
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, A_NAMESPACE, b"ext") =>
            {
                position.observe_ext(&element);
            }
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, P_NAMESPACE, b"ph") =>
            {
                position.observe_placeholder(&element);
            }
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"txBody") => {
                content = Some(parse_text_body(xml, true, hyperlinks)?);
            }
            Event::End(element) if end_is(element.name().as_ref(), b"sp") => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of PPTX shape")),
            _ => {}
        }
    }
    Ok(ShapeData { content, position })
}

#[derive(Default)]
struct ParagraphData {
    runs: Vec<Run>,
    level: u32,
    list: Option<ListKind>,
    list_explicit: bool,
    alignment: ParagraphAlignment,
    default_formatting: Formatting,
}

fn parse_text_body(
    xml: &mut XmlReader<'_>,
    add_newline: bool,
    hyperlinks: &HashMap<String, String>,
) -> Result<ParsedContent> {
    let mut paragraphs = Vec::new();
    loop {
        match event(xml, "DrawingML text body")? {
            Event::Start(element) if element_is(xml, &element, A_NAMESPACE, b"p") => {
                paragraphs.push(parse_paragraph_events(xml, add_newline, hyperlinks)?);
            }
            Event::End(element) if end_is(element.name().as_ref(), b"txBody") => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of DrawingML text body")),
            _ => {}
        }
    }
    Ok(TextBlock {
        role: TextRole::Other,
        paragraphs: paragraphs
            .into_iter()
            .map(|paragraph| Paragraph {
                runs: paragraph.runs,
                alignment: paragraph.alignment,
                list: paragraph.list.map(|kind| ListInfo {
                    level: paragraph.level,
                    kind,
                }),
                list_explicit: paragraph.list_explicit,
            })
            .collect(),
    })
}

fn parse_paragraph_events(
    xml: &mut XmlReader<'_>,
    add_newline: bool,
    hyperlinks: &HashMap<String, String>,
) -> Result<ParagraphData> {
    let mut paragraph = ParagraphData::default();
    loop {
        match event(xml, "DrawingML paragraph")? {
            Event::Start(element) if element_is(xml, &element, A_NAMESPACE, b"pPr") => {
                parse_paragraph_properties(xml, &element, &mut paragraph)?;
            }
            Event::Empty(element) if element_is(xml, &element, A_NAMESPACE, b"pPr") => {
                paragraph.level = attr(&element, b"lvl")
                    .and_then(|value| value.parse().ok())
                    .unwrap_or(0);
                if attr(&element, b"lvl").is_some() {
                    paragraph.list = Some(ListKind::Bullet { character: None });
                }
                paragraph.alignment = paragraph_alignment(attr(&element, b"algn").as_deref());
            }
            Event::Start(element) if element_is(xml, &element, A_NAMESPACE, b"r") => {
                paragraph.runs.push(parse_run_events_with_base(
                    xml,
                    hyperlinks,
                    b"r",
                    paragraph.default_formatting.clone(),
                )?);
            }
            Event::Start(element) if element_is(xml, &element, A_NAMESPACE, b"fld") => {
                paragraph.runs.push(parse_run_events_with_base(
                    xml,
                    hyperlinks,
                    b"fld",
                    paragraph.default_formatting.clone(),
                )?);
            }
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, A_NAMESPACE, b"br") =>
            {
                paragraph.runs.push(Run {
                    text: "\n".to_string(),
                    formatting: paragraph.default_formatting.clone(),
                    link_target: None,
                });
            }
            Event::End(element) if end_is(element.name().as_ref(), b"p") => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of DrawingML paragraph")),
            _ => {}
        }
    }
    if add_newline && let Some(last) = paragraph.runs.last_mut() {
        last.text.push('\n');
    }
    Ok(paragraph)
}

fn parse_paragraph_properties(
    xml: &mut XmlReader<'_>,
    start: &BytesStart<'_>,
    paragraph: &mut ParagraphData,
) -> Result<()> {
    paragraph.level = attr(start, b"lvl")
        .and_then(|value| value.parse().ok())
        .unwrap_or(0);
    if attr(start, b"lvl").is_some() {
        paragraph.list = Some(ListKind::Bullet { character: None });
    }
    paragraph.alignment = paragraph_alignment(attr(start, b"algn").as_deref());
    loop {
        match event(xml, "DrawingML paragraph properties")? {
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, A_NAMESPACE, b"buAutoNum") =>
            {
                paragraph.list = Some(ListKind::Ordered {
                    style: attr(&element, b"type"),
                    start: attr(&element, b"startAt")
                        .and_then(|value| value.parse().ok())
                        .unwrap_or(1),
                });
                paragraph.list_explicit = true;
            }
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, A_NAMESPACE, b"buChar") =>
            {
                paragraph.list = Some(ListKind::Bullet {
                    character: attr(&element, b"char"),
                });
                paragraph.list_explicit = true;
            }
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, A_NAMESPACE, b"buNone") =>
            {
                paragraph.list = None;
                paragraph.list_explicit = true;
            }
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, A_NAMESPACE, b"defRPr") =>
            {
                apply_run_attributes(&element, &mut paragraph.default_formatting);
            }
            Event::End(element) if end_is(element.name().as_ref(), b"pPr") => break,
            Event::Eof => {
                return Err(Error::ParseError(
                    "Unexpected end of DrawingML paragraph properties",
                ));
            }
            _ => {}
        }
    }
    Ok(())
}

#[cfg(test)]
fn parse_run_events(xml: &mut XmlReader<'_>, hyperlinks: &HashMap<String, String>) -> Result<Run> {
    parse_run_events_with_base(xml, hyperlinks, b"r", Formatting::default())
}

fn parse_run_events_with_base(
    xml: &mut XmlReader<'_>,
    hyperlinks: &HashMap<String, String>,
    end: &[u8],
    mut formatting: Formatting,
) -> Result<Run> {
    let mut value = String::new();
    let mut link_id = None;
    loop {
        match event(xml, "DrawingML run")? {
            Event::Start(element) if element_is(xml, &element, A_NAMESPACE, b"rPr") => {
                apply_run_attributes(&element, &mut formatting);
                link_id = parse_run_properties(xml, link_id)?;
            }
            Event::Empty(element) if element_is(xml, &element, A_NAMESPACE, b"rPr") => {
                apply_run_attributes(&element, &mut formatting);
            }
            Event::Start(element) if element_is(xml, &element, A_NAMESPACE, b"t") => {
                value.push_str(&read_simple_text(xml, b"t", "DrawingML run")?);
            }
            Event::End(element) if end_is(element.name().as_ref(), end) => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of DrawingML run")),
            _ => {}
        }
    }
    Ok(Run {
        text: value,
        formatting,
        link_target: link_id.and_then(|id| hyperlinks.get(&id).cloned()),
    })
}

fn apply_run_attributes(element: &BytesStart<'_>, formatting: &mut Formatting) {
    if let Some(value) = attr(element, b"b") {
        formatting.bold = value == "1" || value.eq_ignore_ascii_case("true");
    }
    if let Some(value) = attr(element, b"i") {
        formatting.italic = value == "1" || value.eq_ignore_ascii_case("true");
    }
    if let Some(value) = attr(element, b"u") {
        formatting.underlined = value != "none";
    }
    if let Some(value) = attr(element, b"strike") {
        formatting.strikethrough = value != "noStrike" && value != "0" && value != "false";
    }
    if let Some(value) = attr(element, b"baseline").and_then(|value| value.parse::<i32>().ok()) {
        formatting.baseline = if value > 0 {
            crate::Baseline::Superscript
        } else if value < 0 {
            crate::Baseline::Subscript
        } else {
            crate::Baseline::Normal
        };
    }
    if let Some(value) = attr(element, b"sz").and_then(|value| value.parse::<f32>().ok()) {
        formatting.font_size_points = Some(value / 100.0);
    }
    if let Some(value) = attr(element, b"lang") {
        formatting.lang = value;
    }
}

fn parse_run_properties(
    xml: &mut XmlReader<'_>,
    mut link_id: Option<String>,
) -> Result<Option<String>> {
    loop {
        match event(xml, "DrawingML run properties")? {
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, A_NAMESPACE, b"hlinkClick") =>
            {
                link_id = attr(&element, b"id");
            }
            Event::End(element) if end_is(element.name().as_ref(), b"rPr") => break,
            Event::Eof => {
                return Err(Error::ParseError(
                    "Unexpected end of DrawingML run properties",
                ));
            }
            _ => {}
        }
    }
    Ok(link_id)
}

fn parse_graphic_frame(
    xml: &mut XmlReader<'_>,
    hyperlinks: &HashMap<String, String>,
) -> Result<(Option<TableElement>, PositionData)> {
    let mut position = PositionData::default();
    let mut in_table_data = false;
    let mut table = None;
    loop {
        match event(xml, "PPTX graphic frame")? {
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, A_NAMESPACE, b"off") =>
            {
                position.observe_off(&element);
            }
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, A_NAMESPACE, b"ext") =>
            {
                position.observe_ext(&element);
            }
            Event::Start(element) if element_is(xml, &element, A_NAMESPACE, b"graphicData") => {
                in_table_data = attr(&element, b"uri").as_deref()
                    == Some("http://schemas.openxmlformats.org/drawingml/2006/table");
            }
            Event::Start(element)
                if !in_table_data && element_is(xml, &element, A_NAMESPACE, b"t") =>
            {
                let value = read_simple_text(xml, b"t", "PPTX graphic frame")?;
                if !value.is_empty() {
                    if !position.fallback_text.is_empty() {
                        position.fallback_text.push(' ');
                    }
                    position.fallback_text.push_str(&value);
                }
            }
            Event::Start(element)
                if in_table_data && element_is(xml, &element, A_NAMESPACE, b"tbl") =>
            {
                table = Some(parse_table_events(xml, hyperlinks)?);
            }
            Event::End(element) if end_is(element.name().as_ref(), b"graphicData") => {
                in_table_data = false;
            }
            Event::End(element) if end_is(element.name().as_ref(), b"graphicFrame") => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of PPTX graphic frame")),
            _ => {}
        }
    }
    Ok((table, position))
}

fn parse_table_events(
    xml: &mut XmlReader<'_>,
    hyperlinks: &HashMap<String, String>,
) -> Result<TableElement> {
    let mut rows = Vec::new();
    loop {
        match event(xml, "DrawingML table")? {
            Event::Start(element) if element_is(xml, &element, A_NAMESPACE, b"tr") => {
                rows.push(parse_table_row_events(xml, hyperlinks)?);
            }
            Event::End(element) if end_is(element.name().as_ref(), b"tbl") => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of DrawingML table")),
            _ => {}
        }
    }
    Ok(TableElement { rows })
}

fn parse_table_row_events(
    xml: &mut XmlReader<'_>,
    hyperlinks: &HashMap<String, String>,
) -> Result<TableRow> {
    let mut cells = Vec::new();
    loop {
        match event(xml, "DrawingML table row")? {
            Event::Start(element) if element_is(xml, &element, A_NAMESPACE, b"tc") => {
                cells.push(parse_table_cell_events(xml, &element, hyperlinks)?);
            }
            Event::End(element) if end_is(element.name().as_ref(), b"tr") => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of DrawingML table row")),
            _ => {}
        }
    }
    Ok(TableRow { cells })
}

fn parse_table_cell_events(
    xml: &mut XmlReader<'_>,
    start: &BytesStart<'_>,
    hyperlinks: &HashMap<String, String>,
) -> Result<TableCell> {
    let mut runs = Vec::new();
    let mut paragraphs = Vec::new();
    loop {
        match event(xml, "DrawingML table cell")? {
            Event::Start(element) if element_is(xml, &element, A_NAMESPACE, b"txBody") => {
                let content = parse_text_body(xml, false, hyperlinks)?;
                paragraphs = content.paragraphs.clone();
                runs = content_to_text(content).runs;
            }
            Event::End(element) if end_is(element.name().as_ref(), b"tc") => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of DrawingML table cell")),
            _ => {}
        }
    }
    Ok(TableCell {
        runs,
        paragraphs,
        row_span: attr(start, b"rowSpan")
            .and_then(|value| value.parse().ok())
            .unwrap_or(1),
        column_span: attr(start, b"gridSpan")
            .and_then(|value| value.parse().ok())
            .unwrap_or(1),
        covered: attr(start, b"hMerge").as_deref() == Some("1")
            || attr(start, b"vMerge").as_deref() == Some("1"),
    })
}

fn content_to_text(content: ParsedContent) -> TextElement {
    TextElement {
        runs: content
            .paragraphs
            .into_iter()
            .flat_map(|paragraph| paragraph.runs)
            .collect(),
    }
}

fn content_to_elements(content: ParsedContent, position: ElementPosition) -> Vec<SlideElement> {
    let all_text = content
        .paragraphs
        .iter()
        .all(|paragraph| paragraph.list.is_none());
    let all_list = content
        .paragraphs
        .iter()
        .all(|paragraph| paragraph.list.is_some());
    if all_text {
        return vec![SlideElement::Text(content_to_text(content), position)];
    }
    if all_list {
        return vec![SlideElement::List(
            ListElement {
                items: content
                    .paragraphs
                    .into_iter()
                    .map(paragraph_to_list_item)
                    .collect(),
            },
            position,
        )];
    }
    content
        .paragraphs
        .into_iter()
        .map(|paragraph| {
            if paragraph.list.is_some() {
                SlideElement::List(
                    ListElement {
                        items: vec![paragraph_to_list_item(paragraph)],
                    },
                    position,
                )
            } else {
                SlideElement::Text(
                    TextElement {
                        runs: paragraph.runs,
                    },
                    position,
                )
            }
        })
        .collect()
}

fn paragraph_to_list_item(paragraph: Paragraph) -> ListItem {
    let list = paragraph.list.expect("list paragraph");
    ListItem {
        level: list.level,
        is_ordered: matches!(list.kind, ListKind::Ordered { .. }),
        runs: paragraph.runs,
    }
}

fn paragraph_alignment(value: Option<&str>) -> ParagraphAlignment {
    match value {
        Some("ctr") | Some("center") => ParagraphAlignment::Center,
        Some("r") | Some("right") => ParagraphAlignment::End,
        Some("just") | Some("dist") => ParagraphAlignment::Justify,
        _ => ParagraphAlignment::Start,
    }
}

fn parse_picture(
    xml: &mut XmlReader<'_>,
) -> Result<(ImageReference, PositionData, Option<String>)> {
    let mut position = PositionData::default();
    let mut image_id = None;
    let mut alt_text = None;
    loop {
        match event(xml, "PPTX picture")? {
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, A_NAMESPACE, b"off") =>
            {
                position.observe_off(&element);
            }
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, A_NAMESPACE, b"ext") =>
            {
                position.observe_ext(&element);
            }
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, A_NAMESPACE, b"blip") =>
            {
                image_id = attr(&element, b"embed");
            }
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, P_NAMESPACE, b"cNvPr") =>
            {
                alt_text = attr(&element, b"descr")
                    .or_else(|| attr(&element, b"title"))
                    .or_else(|| attr(&element, b"name"));
            }
            Event::End(element) if end_is(element.name().as_ref(), b"pic") => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of PPTX picture")),
            _ => {}
        }
    }
    Ok((
        ImageReference {
            id: image_id.ok_or(Error::ImageNotFound)?,
            target: String::new(),
        },
        position,
        alt_text,
    ))
}

#[derive(Default)]
struct GroupTransformData {
    off_x: i64,
    off_y: i64,
    child_off_x: i64,
    child_off_y: i64,
    extent_x: i64,
    extent_y: i64,
    child_extent_x: i64,
    child_extent_y: i64,
}

impl GroupTransformData {
    fn finish(&self) -> CoordinateTransform {
        let scale_x = if self.child_extent_x == 0 {
            1.0
        } else {
            self.extent_x as f64 / self.child_extent_x as f64
        };
        let scale_y = if self.child_extent_y == 0 {
            1.0
        } else {
            self.extent_y as f64 / self.child_extent_y as f64
        };
        CoordinateTransform {
            scale_x,
            scale_y,
            translate_x: self.off_x as f64 - self.child_off_x as f64 * scale_x,
            translate_y: self.off_y as f64 - self.child_off_y as f64 * scale_y,
        }
    }
}

fn parse_group_properties(xml: &mut XmlReader<'_>, data: &mut GroupTransformData) -> Result<()> {
    loop {
        match event(xml, "PPTX group properties")? {
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, A_NAMESPACE, b"off") =>
            {
                data.off_x = integer_attr(&element, b"x");
                data.off_y = integer_attr(&element, b"y");
            }
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, A_NAMESPACE, b"chOff") =>
            {
                data.child_off_x = integer_attr(&element, b"x");
                data.child_off_y = integer_attr(&element, b"y");
            }
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, A_NAMESPACE, b"ext") =>
            {
                data.extent_x = integer_attr(&element, b"cx");
                data.extent_y = integer_attr(&element, b"cy");
            }
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, A_NAMESPACE, b"chExt") =>
            {
                data.child_extent_x = integer_attr(&element, b"cx");
                data.child_extent_y = integer_attr(&element, b"cy");
            }
            Event::End(element) if end_is(element.name().as_ref(), b"grpSpPr") => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of PPTX group properties")),
            _ => {}
        }
    }
    Ok(())
}

fn integer_attr(element: &BytesStart<'_>, name: &[u8]) -> i64 {
    attr(element, name)
        .and_then(|value| value.parse().ok())
        .unwrap_or(0)
}

fn read_simple_text(xml: &mut XmlReader<'_>, end: &[u8], part: &str) -> Result<String> {
    let mut value = String::new();
    loop {
        match event(xml, part)? {
            Event::Text(content) => value.push_str(&text(&content, part)?),
            Event::GeneralRef(content) => value.push_str(&reference(&content, part)?),
            Event::CData(content) => value.push_str(&String::from_utf8_lossy(content.as_ref())),
            Event::End(element) if end_is(element.name().as_ref(), end) => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of XML text")),
            _ => {}
        }
    }
    Ok(value)
}

pub fn extract_inherited_positions(
    xml_data: &[u8],
    inherited_positions: &InheritedPositions,
) -> Result<InheritedPositions> {
    let mut xml = reader(xml_data);
    let mut in_common_slide = false;
    loop {
        match event(&mut xml, "PPTX layout or master")? {
            Event::Start(element) if element_is(&xml, &element, P_NAMESPACE, b"cSld") => {
                in_common_slide = true;
            }
            Event::Start(element)
                if in_common_slide && element_is(&xml, &element, P_NAMESPACE, b"spTree") =>
            {
                let mut positions = inherited_positions.positions.clone();
                let mut list_styles = inherited_positions.list_styles.clone();
                collect_placeholder_positions(
                    &mut xml,
                    b"spTree",
                    CoordinateTransform::identity(),
                    inherited_positions,
                    &mut positions,
                    &mut list_styles,
                )?;
                return Ok(InheritedPositions {
                    positions,
                    list_styles,
                });
            }
            Event::Empty(element)
                if in_common_slide && element_is(&xml, &element, P_NAMESPACE, b"spTree") =>
            {
                return Ok(inherited_positions.clone());
            }
            Event::Eof => return Err(Error::ParseError("PPTX placeholder shape tree not found")),
            _ => {}
        }
    }
}

fn collect_placeholder_positions(
    xml: &mut XmlReader<'_>,
    end: &[u8],
    transform: CoordinateTransform,
    inherited: &InheritedPositions,
    positions: &mut HashMap<PlaceholderKey, ElementPosition>,
    list_styles: &mut HashMap<PlaceholderKey, HashMap<u32, ListKind>>,
) -> Result<()> {
    let mut group_transform = GroupTransformData::default();
    loop {
        match event(xml, "PPTX placeholders")? {
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"grpSpPr") => {
                parse_group_properties(xml, &mut group_transform)?;
            }
            Event::Start(element) if element_is(xml, &element, P_NAMESPACE, b"grpSp") => {
                collect_placeholder_positions(
                    xml,
                    b"grpSp",
                    transform.then(group_transform.finish()),
                    inherited,
                    positions,
                    list_styles,
                )?;
            }
            Event::Start(element)
                if element_is(xml, &element, P_NAMESPACE, b"sp")
                    || element_is(xml, &element, P_NAMESPACE, b"pic")
                    || element_is(xml, &element, P_NAMESPACE, b"graphicFrame") =>
            {
                let shape_end = crate::xml::local(element.name().as_ref()).to_vec();
                let (data, styles) = scan_position(xml, &shape_end)?;
                if let Some(key) = data.placeholder.as_ref()
                    && let Some(position) = data
                        .raw()
                        .map(|position| transform.apply(position))
                        .or_else(|| inherited.resolve(key))
                {
                    insert_placeholder_position(positions, key.clone(), position);
                }
                if let Some(key) = data.placeholder.as_ref()
                    && !styles.is_empty()
                {
                    insert_placeholder_list_styles(list_styles, key.clone(), styles);
                }
            }
            Event::End(element) if end_is(element.name().as_ref(), end) => break,
            Event::Eof => return Err(Error::ParseError("Unexpected end of PPTX placeholders")),
            _ => {}
        }
    }
    Ok(())
}

fn scan_position(
    xml: &mut XmlReader<'_>,
    end: &[u8],
) -> Result<(PositionData, HashMap<u32, ListKind>)> {
    let mut data = PositionData::default();
    let mut list_styles = HashMap::new();
    loop {
        match event(xml, "PPTX placeholder shape")? {
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, A_NAMESPACE, b"off") =>
            {
                data.observe_off(&element);
            }
            Event::Start(element) | Event::Empty(element)
                if element_is(xml, &element, P_NAMESPACE, b"ph") =>
            {
                data.observe_placeholder(&element);
            }
            Event::Start(element) if element_is(xml, &element, A_NAMESPACE, b"pPr") => {
                let mut paragraph = ParagraphData::default();
                parse_paragraph_properties(xml, &element, &mut paragraph)?;
                if let Some(kind) = paragraph.list {
                    list_styles.insert(paragraph.level, kind);
                }
            }
            Event::Empty(element) if element_is(xml, &element, A_NAMESPACE, b"pPr") => {
                if let Some(level) = attr(&element, b"lvl").and_then(|value| value.parse().ok()) {
                    list_styles.insert(level, ListKind::Bullet { character: None });
                }
            }
            Event::End(element) if end_is(element.name().as_ref(), end) => break,
            Event::Eof => {
                return Err(Error::ParseError(
                    "Unexpected end of PPTX placeholder shape",
                ));
            }
            _ => {}
        }
    }
    Ok((data, list_styles))
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

fn insert_placeholder_list_styles(
    list_styles: &mut HashMap<PlaceholderKey, HashMap<u32, ListKind>>,
    key: PlaceholderKey,
    styles: HashMap<u32, ListKind>,
) {
    if let Some(idx) = key.idx.as_deref() {
        list_styles.retain(|candidate, _| candidate.idx.as_deref() != Some(idx));
    }
    list_styles.insert(key, styles);
}

#[cfg(test)]
#[path = "../tests/unit/parse_xml.rs"]
mod tests;
