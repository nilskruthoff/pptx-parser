mod container;
mod slide;
mod types;
mod constants;

pub use container::PptxContainer;
pub use slide::parse_slide_xml;
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

    #[error("Unbekannter Fehler")]
    Unknown,
}

pub type Result<T> = std::result::Result<T, Error>;