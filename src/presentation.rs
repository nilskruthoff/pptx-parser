use crate::container::SlideIterator;
use crate::odp::{OdpContainer, OdpSlideIterator};
use crate::{ParserConfig, PptxContainer, Presentation, PresentationMetadata, Result, Slide};
use std::io::Read;
use std::path::Path;

/// The presentation format detected by [`PresentationContainer`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum PresentationFormat {
    Pptx,
    Odp,
}

enum ContainerInner {
    Pptx(PptxContainer),
    Odp(OdpContainer),
}

/// Opens either a PowerPoint or an OpenDocument presentation.
///
/// This is additive to [`PptxContainer`], which continues to provide the
/// existing PPTX-only API unchanged.
pub struct PresentationContainer {
    format: PresentationFormat,
    inner: ContainerInner,
}

impl PresentationContainer {
    /// Opens a `.pptx` or `.odp` presentation by detecting the file format from
    /// the package contents.
    ///
    /// Use this as the default entry point when the input may be either format.
    pub fn open(path: &Path, config: ParserConfig) -> Result<Self> {
        let format = detect_format(path)?;
        Self::open_as(path, config, format)
    }

    /// Opens a presentation as the explicitly provided format.
    ///
    /// Use this when the caller already knows the format and wants to skip
    /// auto-detection, or when a file extension is not a reliable signal.
    pub fn open_as(path: &Path, config: ParserConfig, format: PresentationFormat) -> Result<Self> {
        let inner = match format {
            PresentationFormat::Pptx => ContainerInner::Pptx(PptxContainer::open(path, config)?),
            PresentationFormat::Odp => ContainerInner::Odp(OdpContainer::open(path, config)?),
        };

        Ok(Self { format, inner })
    }

    pub fn format(&self) -> PresentationFormat {
        self.format
    }

    pub fn metadata(&self) -> &PresentationMetadata {
        match &self.inner {
            ContainerInner::Pptx(container) => container.metadata(),
            ContainerInner::Odp(container) => container.metadata(),
        }
    }

    pub fn parse_all(&mut self) -> Result<Vec<Slide>> {
        match &mut self.inner {
            ContainerInner::Pptx(container) => container.parse_all(),
            ContainerInner::Odp(container) => container.parse_all(),
        }
    }

    /// Parses the complete presentation into the semantic document model.
    pub fn parse_document(&mut self) -> Result<Presentation> {
        let metadata = self.metadata().clone();
        let slides = self.parse_all()?;
        let diagnostics = slides
            .iter()
            .flat_map(|slide| slide.diagnostics.iter().cloned())
            .collect();
        Ok(Presentation {
            metadata,
            slides,
            diagnostics,
        })
    }

    pub fn parse_all_multi_threaded(&mut self) -> Result<Vec<Slide>> {
        match &mut self.inner {
            ContainerInner::Pptx(container) => container.parse_all_multi_threaded(),
            // ODP stores all pages in one content.xml, so there is no independent
            // slide XML to preload as there is for PPTX.
            ContainerInner::Odp(container) => container.parse_all(),
        }
    }

    pub fn convert_to_md(&mut self) -> Result<String> {
        match &mut self.inner {
            ContainerInner::Pptx(container) => container.convert_to_md(),
            ContainerInner::Odp(container) => container.convert_to_md(),
        }
    }

    pub fn convert_to_md_multi_threaded(&mut self) -> Result<String> {
        match &mut self.inner {
            ContainerInner::Pptx(container) => container.convert_to_md_multi_threaded(),
            ContainerInner::Odp(container) => container.convert_to_md(),
        }
    }

    pub fn iter_slides(&mut self) -> PresentationSlideIterator<'_> {
        let inner = match &mut self.inner {
            ContainerInner::Pptx(container) => {
                PresentationIteratorInner::Pptx(container.iter_slides())
            }
            ContainerInner::Odp(container) => {
                PresentationIteratorInner::Odp(container.iter_slides())
            }
        };
        PresentationSlideIterator { inner }
    }
}

fn detect_format(path: &Path) -> Result<PresentationFormat> {
    let file = std::fs::File::open(path)?;
    let mut archive = zip::ZipArchive::new(file)?;

    if let Ok(mut mimetype) = archive.by_name("mimetype") {
        let mut value = String::new();
        mimetype.read_to_string(&mut value)?;
        if value.trim() == "application/vnd.oasis.opendocument.presentation" {
            return Ok(PresentationFormat::Odp);
        }
    }

    if archive.by_name("[Content_Types].xml").is_ok()
        && archive.by_name("ppt/presentation.xml").is_ok()
    {
        return Ok(PresentationFormat::Pptx);
    }

    Err(crate::Error::ParseError("Unsupported presentation format"))
}

/// Iterator returned by [`PresentationContainer::iter_slides`].
pub struct PresentationSlideIterator<'a> {
    inner: PresentationIteratorInner<'a>,
}

enum PresentationIteratorInner<'a> {
    Pptx(SlideIterator<'a>),
    Odp(OdpSlideIterator<'a>),
}

impl Iterator for PresentationSlideIterator<'_> {
    type Item = Result<Slide>;

    fn next(&mut self) -> Option<Self::Item> {
        match &mut self.inner {
            PresentationIteratorInner::Pptx(iterator) => iterator.next(),
            PresentationIteratorInner::Odp(iterator) => iterator.next(),
        }
    }
}
