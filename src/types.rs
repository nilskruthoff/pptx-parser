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
}
#[derive(Debug)]
pub struct TextElement {
    pub text: String,
    pub formatting: Formatting,
}
#[derive(Debug, Default)]
pub struct Formatting {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
}