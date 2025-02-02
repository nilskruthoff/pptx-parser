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
    Unknown,
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