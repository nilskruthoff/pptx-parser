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
/// | `compress_images` | `bool` | `true` | Whether images are compressed before encoding or not |
/// | `image_quality` | `f32` | `80.0` | Compression level (0.0-100.0);<br/> higher values retain more detail but increase file size |
/// | `image_size_ratio` | `f32` | `0.8` | Scaling factor for image dimensions (0-1.0);<br/> smaller values reduce resolution and file size |
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
    pub compress_images: bool,
    pub quality: f32,
    pub size_ratio: f32,
}

impl Default for ParserConfig {
    fn default() -> Self {
        Self { 
            extract_images: true,
            compress_images: true,
            quality: 80.0,
            size_ratio: 0.8,
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
    compress_images: Option<bool>,
    image_quality: Option<f32>,
    image_size_ratio: Option<f32>,
}

impl ParserConfigBuilder {
    /// Sets weather images should be extracted from the slides.
    pub fn extract_images(mut self, value: bool) -> Self {
        self.extract_images = Some(value);
        self
    }
    
    /// Sets weather images should be compressed before encoded to base64 or not
    pub fn compress_images(mut self, value: bool) -> Self {
        self.compress_images = Some(value);
        self
    }
    
    /// Specifies the desired image quality where `100` is the original quality and `50` means half the quality
    /// The lower the quality, the smaller the file size of the output image will be
    pub fn quality(mut self, value: f32) -> Self {
        self.image_quality = Some(value);
        self
    }
    
    /// Specifies the ratio of the new size compared to the original image, where `100` is the size of the origional
    /// image and `50` halves the _width_ and _height_.
    /// The lower the image size, the smaller the file size of the output image will be
    pub fn size_ratio(mut self, value: f32) -> Self {
        self.image_size_ratio = Some(value);
        self
    }

    /// Builds the final [`ParserConfig`] instance, applying default values for any fields that were not set.
    pub fn build(self) -> ParserConfig {
        ParserConfig {
            extract_images: self.extract_images.unwrap_or(true),
            compress_images: true,
            quality: 80.0,
            size_ratio: 0.8,
        }
    }
}