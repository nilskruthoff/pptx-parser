mod constants;
mod container;
mod markdown;
mod metadata;
mod odp;
pub mod parse_rels;
pub mod parse_xml;
mod parser_config;
mod presentation;
mod slide;
mod types;
mod xml;

pub use container::PptxContainer;
pub use metadata::PresentationMetadata;
pub use parser_config::{ImageHandlingMode, ParserConfig};
pub use presentation::{PresentationContainer, PresentationFormat, PresentationSlideIterator};
pub use slide::Slide;
pub use types::*;

/// Error raised while reading the ZIP package of a presentation.
///
/// The error of the underlying ZIP library is available through
/// [`std::error::Error::source`], but its concrete type is not part of the
/// public API, so updating that library does not change this crate's API.
#[derive(Debug)]
pub struct ZipError(zip::result::ZipError);

impl std::fmt::Display for ZipError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        self.0.fmt(f)
    }
}

impl std::error::Error for ZipError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        Some(&self.0)
    }
}

/// Error raised while parsing an XML part of a presentation.
///
/// The error of the underlying XML library is available through
/// [`std::error::Error::source`], but its concrete type is not part of the
/// public API, so updating that library does not change this crate's API.
#[derive(Debug)]
pub struct XmlError(quick_xml::Error);

impl std::fmt::Display for XmlError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        self.0.fmt(f)
    }
}

impl std::error::Error for XmlError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        Some(&self.0)
    }
}

#[derive(Debug, thiserror::Error)]
pub enum Error {
    #[error("Zip error: {0}")]
    Zip(#[from] ZipError),

    #[error("XML parse error in {part}: {source}")]
    Xml {
        part: String,
        #[source]
        source: XmlError,
    },

    #[error("UTF-8 conversion error: {0}")]
    Utf8(#[from] std::str::Utf8Error),

    #[error("Slide not found")]
    SlideNotFound,

    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    #[error("Parse error: {0}")]
    ParseError(&'static str),

    #[error("Image not found")]
    ImageNotFound,

    #[error("Relationship not found")]
    RelationshipNotFound,

    #[error("Conversion was not possible")]
    ConversionFailed,

    #[error("Conversion was not possible")]
    MultiThreadedConversionFailed,

    #[error("Unknown Error")]
    Unknown,
}

impl Error {
    /// Wraps an error of the ZIP library without exposing it through a public `From` impl.
    pub(crate) fn zip(error: zip::result::ZipError) -> Self {
        Self::Zip(ZipError(error))
    }

    /// Wraps an error of the XML library for the given package part.
    pub(crate) fn xml(part: &str, error: impl Into<quick_xml::Error>) -> Self {
        Self::Xml {
            part: part.to_string(),
            source: XmlError(error.into()),
        }
    }
}

pub type Result<T> = std::result::Result<T, Error>;
