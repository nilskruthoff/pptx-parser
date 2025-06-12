#[derive(Debug, Clone)]
pub struct ParserConfig {
    pub extract_images: bool,
}

impl Default for ParserConfig {
    fn default() -> Self {
        Self { 
            extract_images: true 
        }
    }
}

impl ParserConfig {
    pub fn builder() -> ParserConfigBuilder {
        ParserConfigBuilder::default()
    }
}

#[derive(Debug, Default)]
pub struct ParserConfigBuilder {
    extract_images: Option<bool>,
}

impl ParserConfigBuilder {
    pub fn extract_images(mut self, value: bool) -> Self {
        self.extract_images = Some(value);
        self
    }
    
    pub fn build(self) -> ParserConfig {
        ParserConfig {
            extract_images: self.extract_images.unwrap_or(true),
        }
    }
}