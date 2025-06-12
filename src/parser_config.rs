/// Configuration options for the PPTX parser.
///
/// Use [`ParserConfig::builder()`] to create a configuration instance.
/// This allows you to customize only the desired fields while falling back to sensible defaults for the rest.
///
/// # Configuration Options
/// 
/// | Parameter | Type | Default | Description |
/// |-----------|------|---------|-------------|
/// | `extract_images` | `bool` | `true` | Whether images are extracted from slides or not |
///
/// # Example
///
/// ```
/// use pptx_to_md::ParserConfig;
///
/// let config = ParserConfig::builder()
///     .extract_images(true)
///     .build();
/// ```
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

/// Builder for [`ParserConfig`].
///
/// Allows setting individual configuration fields while falling back to defaults for any unspecified values
#[derive(Debug, Default)]
pub struct ParserConfigBuilder {
    extract_images: Option<bool>,
}

impl ParserConfigBuilder {
    /// Sets weather images should be extracted from the slides.
    pub fn extract_images(mut self, value: bool) -> Self {
        self.extract_images = Some(value);
        self
    }

    /// Builds the final [`ParserConfig`] instance, applying default values for any fields that were not set.
    pub fn build(self) -> ParserConfig {
        ParserConfig {
            extract_images: self.extract_images.unwrap_or(true),
        }
    }
}