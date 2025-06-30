#[derive(Debug)]
pub struct Presentation {
    pub slides: Vec<Slide>,
}

#[derive(Debug)]
pub struct Slide {
    pub elements: Vec<SlideElement>,
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
    pub lang: String,
}

#[derive(Debug, Clone)]
pub struct Run {
    pub text: String,
    pub formatting: Formatting,
}

impl Run {
    pub fn extract(&self) -> String {
        self.text.to_string()
    }

    pub fn render_as_md(&self) -> String {
        let mut has_new_line = false;

        let mut result = self.extract();
        if result.ends_with("\n") {
            has_new_line = true;
            result = result.replace('\n', "");
        }

        if self.formatting.bold && self.formatting.italic {
            result = format!("***{}***", result);
        } else {
            if self.formatting.bold {
                result = format!("**{}**", result);
            }
            if self.formatting.italic {
                result = format!("_{}_", result);
            }
        }

        if self.formatting.underlined {
            result = format!("<u>{}</u>", result);
        }

        if has_new_line {
            return format!("{}\n", result)
        }
        
        result
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

#[derive(Debug, Clone)]
pub struct TableCell {
    pub runs: Vec<Run>,
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