#[derive(Debug)]
pub struct Presentation {
    pub slides: Vec<Slide>,
}

#[derive(Debug)]
pub struct Slide {
    pub elements: Vec<SlideElement>,
}

#[derive(Debug)]
pub enum SlideElement {
    Text(TextElement),
    Table(TableElement),
    Image(ImageReference),
    List(ListElement),
    Unknown,
}

#[derive(Debug)]
pub struct ImageReference {
    pub id: String,
    pub target: String,
}

#[derive(Debug)]
pub struct TextElement {
    pub runs: Vec<Run>,
}

#[derive(Debug, Default)]
pub struct Formatting {
    pub bold: bool,
    pub italic: bool,
    pub underlined: bool,
    pub lang: String,
}

#[derive(Debug)]
pub struct Run {
    pub text: String,
    pub formatting: Formatting,
}

impl Run {
    pub fn extract(&self) -> String {
        self.text.to_string()
    }
}

#[derive(Debug)]
pub struct TableElement {
    pub rows: Vec<TableRow>,
}

#[derive(Debug)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
}

#[derive(Debug)]
pub struct TableCell {
    pub runs: Vec<Run>,
}

#[derive(Debug)]
pub struct ListElement {
    pub items: Vec<ListItem>,
}

#[derive(Debug)]
pub struct ListItem {
    pub level: u32,
    pub is_ordered: bool,
    pub runs: Vec<Run>,
}