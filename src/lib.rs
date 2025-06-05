mod container;
mod slide;
mod types;
mod constants;
mod parse_xml;
mod parse_rels;

pub use container::PptxContainer;
pub use slide::Slide;
pub use types::*;

#[derive(Debug, thiserror::Error)]
pub enum Error {
    #[error("Zip error: {0}")]
    Zip(#[from] zip::result::ZipError),

    #[error("XML parse error: {0}")]
    Xml(#[from] roxmltree::Error),

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

    #[error("Unbekannter Fehler")]
    Unknown,
}

pub type Result<T> = std::result::Result<T, Error>;