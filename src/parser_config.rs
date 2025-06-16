use std::path::PathBuf;

/// Determines how images are handled during content export.
///
/// # Members
///
/// | Member                | Description                                                                                                                       |
/// |-----------------------|-----------------------------------------------------------------------------------------------------------------------------------|            
/// | `InMarkdown`          | Images are embedded directly in the Markdown output using standard syntax as `base64` data (`![]()`)                              |            
/// | `Manually`            | Image handling is delegated to the user, requiring manual copying or referencing (as `base64` encoded string)                     |            
/// | `Save`                | Images will be saved in a provided output directory and integrated using `<a>` tag syntax (`<a href="file:///<abs_path>"></a>`)   |            
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ImageHandlingMode {
    InMarkdown,
    Manually,
    Save,
}

/// Configuration options for the PPTX parser.
///
/// Use [`ParserConfig::builder()`] to create a configuration instance.
/// This allows you to customize only the desired fields while falling back to sensible defaults for the rest.
///
/// # Configuration Options
///
/// | Parameter                 | Type                  | Default       | Description                                                                                               |
/// |---------------------------|-----------------------|---------------|-----------------------------------------------------------------------------------------------------------|
/// | `extract_images`          | `bool`                | `true`        | Whether images are extracted from slides or not. If false, images can not be extracted manually either    |
/// | `compress_images`         | `bool`                | `true`        | Whether images are compressed before encoding or not. Effects manually extracted images too               |
/// | `image_quality`           | `u8`                  | `80`          | Compression level (0-100);<br/> higher values retain more detail but increase file size                   |
/// | `image_handling_mode`     | `ImageHandlingMode`   | `InMarkdown`  | Determines how images are handled during content export                                                   |
/// | `image_output_path`       | `Option<PathBuf>`     | `None`        | Output directory path for `ImageHandlingMode::Save` (mandatory for the saving mode)                       |
///
/// # Example
///
/// ```
/// use std::path::PathBuf;
/// use pptx_to_md::{ImageHandlingMode, ParserConfig};
///
/// let config = ParserConfig::builder()
///     .extract_images(true)
///     .compress_images(true)
///     .quality(75)
///     .image_handling_mode(ImageHandlingMode::Save)
///     .image_output_path(PathBuf::from("/path/to/output/dir/"))
///     .build();
/// ```
#[derive(Debug, Clone)]
pub struct ParserConfig {
    pub extract_images: bool,
    pub compress_images: bool,
    pub quality: u8,
    pub image_handling_mode: ImageHandlingMode,
    pub image_output_path: Option<PathBuf>,
}

impl Default for ParserConfig {
    fn default() -> Self {
        Self {
            extract_images: true,
            compress_images: true,
            quality: 80,
            image_handling_mode: ImageHandlingMode::InMarkdown,
            image_output_path: None,
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
    image_quality: Option<u8>,
    image_handling_mode: Option<ImageHandlingMode>,
    image_output_path: Option<PathBuf>,
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
    pub fn quality(mut self, value: u8) -> Self {
        self.image_quality = Some(value);
        self
    }

    /// Specifies the mode for processing the image after its extracted
    pub fn image_handling_mode(mut self, value: ImageHandlingMode) -> Self {
        self.image_handling_mode = Some(value);
        self
    }

    /// Specifies the output directory for the [`ImageHandlingMode::Save`]
    pub fn image_output_path<P>(mut self, path: P) -> Self
    where
        P: Into<PathBuf>,
    {
        self.image_output_path = Some(path.into());
        self
    }

    /// Builds the final [`ParserConfig`] instance, applying default values for any fields that were not set.
    pub fn build(self) -> ParserConfig {
        ParserConfig {
            extract_images: self.extract_images.unwrap_or(true),
            compress_images: self.compress_images.unwrap_or(true),
            quality: self.image_quality.unwrap_or(80),
            image_handling_mode: self.image_handling_mode.unwrap_or(ImageHandlingMode::InMarkdown),
            image_output_path: self.image_output_path,
        }
    }
}