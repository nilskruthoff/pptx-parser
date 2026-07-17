#[derive(Debug)]
pub struct Presentation {
    pub metadata: crate::PresentationMetadata,
    pub slides: Vec<crate::Slide>,
    pub diagnostics: Vec<ParseDiagnostic>,
}

#[derive(Debug, Clone)]
pub enum SlideElement {
    Text(TextElement, ElementPosition),
    Table(TableElement, ElementPosition),
    Image(ImageReference, ElementPosition),
    List(ListElement, ElementPosition),
    Unknown,
}

impl SlideElement {
    pub fn position(&self) -> ElementPosition {
        match self {
            SlideElement::Text(_, pos)
            | SlideElement::Image(_, pos)
            | SlideElement::List(_, pos)
            | SlideElement::Table(_, pos) => *pos,
            SlideElement::Unknown => ElementPosition::default(),
        }
    }
}

#[derive(Debug, Clone)]
pub struct ImageReference {
    pub id: String,
    pub target: String,
}

#[derive(Debug, Clone)]
pub struct TextElement {
    pub runs: Vec<Run>,
}

#[derive(Debug, Default, Clone)]
pub struct Formatting {
    pub bold: bool,
    pub italic: bool,
    pub underlined: bool,
    pub strikethrough: bool,
    pub baseline: Baseline,
    pub font_size_points: Option<f32>,
    pub lang: String,
}

#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
pub enum Baseline {
    #[default]
    Normal,
    Superscript,
    Subscript,
}

#[derive(Debug, Clone)]
pub struct Run {
    pub text: String,
    pub formatting: Formatting,
    pub link_target: Option<String>,
}

impl Run {
    pub fn extract(&self) -> String {
        self.text.to_string()
    }

    pub fn render_as_md(&self) -> String {
        crate::markdown::render_run(self)
    }
}

#[derive(Debug, Clone)]
pub struct TableElement {
    pub rows: Vec<TableRow>,
}

#[derive(Debug, Clone)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
}

#[derive(Debug, Clone, Default)]
pub struct TableCell {
    pub runs: Vec<Run>,
    pub paragraphs: Vec<Paragraph>,
    pub row_span: usize,
    pub column_span: usize,
    pub covered: bool,
}

#[derive(Debug, Clone)]
pub struct ListElement {
    pub items: Vec<ListItem>,
}

#[derive(Debug, Clone)]
pub struct ListItem {
    pub level: u32,
    pub is_ordered: bool,
    pub runs: Vec<Run>,
}

#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub struct ElementPosition {
    pub x: i64,
    pub y: i64,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Bounds {
    pub x: i64,
    pub y: i64,
    pub width: i64,
    pub height: i64,
}

impl From<ElementPosition> for Bounds {
    fn from(position: ElementPosition) -> Self {
        Self {
            x: position.x,
            y: position.y,
            width: 0,
            height: 0,
        }
    }
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum TextRole {
    Title,
    Subtitle,
    Heading,
    Body,
    Caption,
    #[default]
    Other,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum ParagraphAlignment {
    #[default]
    Start,
    Center,
    End,
    Justify,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ListKind {
    Bullet { character: Option<String> },
    Ordered { style: Option<String>, start: u32 },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ListInfo {
    pub level: u32,
    pub kind: ListKind,
}

#[derive(Debug, Clone, Default)]
pub struct Paragraph {
    pub runs: Vec<Run>,
    pub alignment: ParagraphAlignment,
    pub list: Option<ListInfo>,
    pub list_explicit: bool,
}

impl Paragraph {
    pub fn plain(runs: Vec<Run>) -> Self {
        Self {
            runs,
            ..Self::default()
        }
    }

    pub fn text(&self) -> String {
        self.runs.iter().map(|run| run.text.as_str()).collect()
    }
}

#[derive(Debug, Clone, Default)]
pub struct TextBlock {
    pub role: TextRole,
    pub paragraphs: Vec<Paragraph>,
}

#[derive(Debug, Clone, Default)]
pub struct SemanticTableCell {
    pub paragraphs: Vec<Paragraph>,
    pub row_span: usize,
    pub column_span: usize,
    pub covered: bool,
}

#[derive(Debug, Clone, Default)]
pub struct SemanticTableRow {
    pub cells: Vec<SemanticTableCell>,
}

#[derive(Debug, Clone, Default)]
pub struct SemanticTable {
    pub rows: Vec<SemanticTableRow>,
}

#[derive(Debug, Clone)]
pub struct ImageBlock {
    pub reference: ImageReference,
    pub alt_text: Option<String>,
    pub mime_type: Option<String>,
}

#[derive(Debug, Clone)]
pub struct UnsupportedBlock {
    pub kind: String,
    pub fallback_text: Option<String>,
}

#[derive(Debug, Clone)]
pub enum SlideBlockContent {
    Text(TextBlock),
    Table(SemanticTable),
    Image(ImageBlock),
    Unsupported(UnsupportedBlock),
}

#[derive(Debug, Clone)]
pub struct SlideBlock {
    pub bounds: Bounds,
    pub source_order: usize,
    pub content: SlideBlockContent,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum ReadingOrder {
    Source,
    #[default]
    Spatial,
}

#[derive(Debug, Clone)]
pub struct MarkdownOptions {
    pub reading_order: ReadingOrder,
    pub include_slide_number_as_comment: bool,
    pub include_speaker_notes: bool,
    pub include_comments: bool,
    pub render_unsupported_comments: bool,
}

impl Default for MarkdownOptions {
    fn default() -> Self {
        Self {
            reading_order: ReadingOrder::Spatial,
            include_slide_number_as_comment: true,
            include_speaker_notes: false,
            include_comments: false,
            render_unsupported_comments: true,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DiagnosticSeverity {
    Warning,
    Error,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParseDiagnostic {
    pub severity: DiagnosticSeverity,
    pub message: String,
    pub source: Option<String>,
}

#[cfg(test)]
#[path = "../tests/unit/types.rs"]
mod tests;
