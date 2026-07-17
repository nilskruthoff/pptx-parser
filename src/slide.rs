use crate::markdown::{MarkdownContext, render_runs};
use crate::parser_config::ImageHandlingMode;
use crate::{
    Bounds, ImageBlock, ImageReference, ListInfo, ListKind, MarkdownOptions, Paragraph,
    ParseDiagnostic, ParserConfig, ReadingOrder, Result, SemanticTable, SemanticTableCell,
    SemanticTableRow, SlideBlock, SlideBlockContent, SlideElement, TextBlock, TextRole,
    UnsupportedBlock,
};
use base64::{Engine as _, engine::general_purpose};
use image::codecs::jpeg::JpegEncoder;
use std::collections::HashMap;
use std::fs;
use std::path::{Path, PathBuf};

/// Encapsulates images for manual extraction of images from slides
#[derive(Debug)]
pub struct ManualImage {
    pub base64_content: String,
    pub img_ref: ImageReference,
}
impl ManualImage {
    pub fn new(base64_content: String, img_ref: ImageReference) -> ManualImage {
        Self {
            base64_content,
            img_ref,
        }
    }
}
/// Represents a single slide extracted from a PowerPoint (pptx) file.
///
/// Contains structured slide data including slide number, parsed content elements
/// (text, tables, images, lists), speaker notes, and associated image references.
///
/// A `Slide` can be converted into other formats, such as Markdown, or its
/// contained images can be extracted in base64 representation.
///
/// Typically, you retrieve instances of `Slide` through [`PptxContainer::parse()`].
#[derive(Debug)]
pub struct Slide {
    pub rel_path: String,
    pub slide_number: u32,
    pub elements: Vec<SlideElement>,
    pub speaker_notes: Vec<crate::TextElement>,
    pub comments: Vec<crate::TextElement>,
    pub images: Vec<ImageReference>,
    pub image_data: HashMap<String, Vec<u8>>,
    pub config: ParserConfig,
    pub blocks: Vec<SlideBlock>,
    pub diagnostics: Vec<ParseDiagnostic>,
}

impl Slide {
    #[allow(clippy::too_many_arguments)]
    pub fn new(
        rel_path: String,
        slide_number: u32,
        elements: Vec<SlideElement>,
        speaker_notes: Vec<crate::TextElement>,
        comments: Vec<crate::TextElement>,
        images: Vec<ImageReference>,
        image_data: HashMap<String, Vec<u8>>,
        config: ParserConfig,
    ) -> Self {
        let blocks = legacy_blocks(&elements);
        Self {
            rel_path,
            slide_number,
            elements,
            speaker_notes,
            comments,
            images,
            image_data,
            config,
            blocks,
            diagnostics: Vec::new(),
        }
    }

    #[allow(clippy::too_many_arguments)]
    pub fn new_semantic(
        rel_path: String,
        slide_number: u32,
        elements: Vec<SlideElement>,
        blocks: Vec<SlideBlock>,
        speaker_notes: Vec<crate::TextElement>,
        comments: Vec<crate::TextElement>,
        images: Vec<ImageReference>,
        image_data: HashMap<String, Vec<u8>>,
        config: ParserConfig,
        diagnostics: Vec<ParseDiagnostic>,
    ) -> Self {
        Self {
            rel_path,
            slide_number,
            elements,
            speaker_notes,
            comments,
            images,
            image_data,
            config,
            blocks,
            diagnostics,
        }
    }

    /// Converts slide contents into a Markdown formatted string.
    ///
    /// Translates internal slide elements (text, tables, lists, images) to valid
    /// and readable Markdown. Embedded images will be encoded as base64 inline images.
    ///
    /// # Returns
    ///
    /// Returns an `Option<String>`:
    /// - `Some(String)`: Markdown representation of slide if conversion succeeds.
    /// - `None`: If a conversion error occurs during image encoding.
    pub fn convert_to_md(&self) -> Result<String> {
        let options = MarkdownOptions {
            include_slide_number_as_comment: self.config.include_slide_number_as_comment,
            include_speaker_notes: self.config.include_speaker_notes,
            include_comments: self.config.include_comments,
            ..MarkdownOptions::default()
        };
        self.to_markdown(&options)
    }

    pub fn to_markdown(&self, options: &MarkdownOptions) -> Result<String> {
        let mut slide_txt = String::new();
        if options.include_slide_number_as_comment {
            slide_txt.push_str(format!("<!-- Slide {} -->\n\n", self.slide_number).as_str());
        }
        let mut image_count = 0;
        let fallback_blocks;
        let blocks = if self.blocks.is_empty() {
            fallback_blocks = legacy_blocks(&self.elements);
            &fallback_blocks
        } else {
            &self.blocks
        };

        for block in ordered_blocks(blocks, options.reading_order) {
            match &block.content {
                SlideBlockContent::Text(text) => {
                    render_text_block(&mut slide_txt, text);
                    if !slide_txt.ends_with("\n\n") {
                        slide_txt.push('\n');
                    }
                }
                SlideBlockContent::Table(table) => render_table(&mut slide_txt, table),
                SlideBlockContent::Image(image) => {
                    let image_ref = &image.reference;
                    match self.config.image_handling_mode {
                        ImageHandlingMode::InMarkdown => {
                            if let Some(image_data) = self.image_data.get(&image_ref.id) {
                                let image_data = if self.config.compress_images {
                                    self.compress_image(image_data)
                                } else {
                                    Some(image_data.clone())
                                };

                                let Some(image_data) = image_data else {
                                    slide_txt.push_str(&missing_image_markdown(image));
                                    continue;
                                };
                                let base64_string = general_purpose::STANDARD.encode(image_data);
                                let image_name =
                                    image_ref.target.split('/').next_back().unwrap_or("image");
                                let file_ext = image
                                    .mime_type
                                    .as_deref()
                                    .and_then(|mime| mime.split('/').next_back())
                                    .or_else(|| image_name.rsplit('.').next())
                                    .unwrap_or("bin");
                                let alt = image.alt_text.as_deref().unwrap_or(image_name);

                                slide_txt.push_str(
                                    format!(
                                        "![{}](data:image/{};base64,{})",
                                        alt, file_ext, base64_string
                                    )
                                    .as_str(),
                                );
                            } else {
                                slide_txt.push_str(&missing_image_markdown(image));
                            }
                        }
                        ImageHandlingMode::Save => {
                            if let Some(image_data) = self.image_data.get(&image_ref.id) {
                                let image_data = if self.config.compress_images {
                                    self.compress_image(image_data)
                                } else {
                                    Some(image_data.clone())
                                };

                                let ext = if self.config.compress_images {
                                    "jpg".to_string()
                                } else {
                                    self.get_image_extension(&image_ref.target)
                                };

                                let output_dir = self
                                    .config
                                    .image_output_path
                                    .clone()
                                    .unwrap_or_else(|| PathBuf::from("."));

                                fs::create_dir_all(&output_dir)?;

                                let mut image_path = output_dir.clone();
                                let file_name = format!(
                                    "slide{}_image{}_{}.{}",
                                    self.slide_number,
                                    image_count + 1,
                                    &image_ref.id,
                                    ext
                                );
                                image_path.push(&file_name);

                                let Some(image_data) = image_data else {
                                    slide_txt.push_str(&missing_image_markdown(image));
                                    continue;
                                };
                                fs::write(&image_path, image_data)?;

                                let abs_file_url = self.path_to_file_url(&image_path);
                                let Some(abs_file_url) = abs_file_url else {
                                    slide_txt.push_str(&missing_image_markdown(image));
                                    continue;
                                };
                                let alt = image.alt_text.as_deref().unwrap_or(&file_name);
                                let html_link = format!("![{alt}]({abs_file_url})");
                                image_count += 1;
                                slide_txt.push_str(&html_link);
                                slide_txt.push('\n');
                            } else {
                                slide_txt.push_str(&missing_image_markdown(image));
                            }
                        }
                        ImageHandlingMode::Manually => {
                            slide_txt.push('\n');
                            continue;
                        }
                    }
                    slide_txt.push('\n');
                }
                SlideBlockContent::Unsupported(unsupported) => {
                    if let Some(text) = &unsupported.fallback_text {
                        slide_txt.push_str(text);
                        slide_txt.push_str("\n\n");
                    }
                    if options.render_unsupported_comments {
                        slide_txt.push_str(&format!(
                            "<!-- Unsupported slide element: {} -->\n\n",
                            unsupported.kind.replace("--", "—")
                        ));
                    }
                }
            }
        }
        if options.include_speaker_notes && !self.speaker_notes.is_empty() {
            append_quoted_section(&mut slide_txt, "Speaker Notes", &self.speaker_notes);
        }
        if options.include_comments && !self.comments.is_empty() {
            append_quoted_section(&mut slide_txt, "Comments", &self.comments);
        }
        Ok(slide_txt)
    }

    /// Extracts the numeric slide identifier from a slide path.
    ///
    /// Helper method to parse slide numbers from internal pptx
    /// slide paths (e.g., "ppt/slides/slide1.xml" → `1`).
    pub fn extract_slide_number(path: &str) -> Option<u32> {
        path.split('/')
            .next_back()
            .and_then(|filename| {
                filename
                    .strip_prefix("slide")
                    .and_then(|s| s.strip_suffix(".xml"))
            })
            .and_then(|num_str| num_str.parse::<u32>().ok())
    }

    /// Links slide images references with their corresponding targets.
    ///
    /// Ensures that each image referenced by its ID is correctly
    /// linked to the actual internal resource paths stored in the slide.
    /// This method is typically used internally after parsing a slide
    ///
    /// # Notes
    ///
    /// Internally those are the values image references are holding
    ///
    /// | Parameter | Example value         |
    /// |---------- |---------------------- |
    /// | `id`      | *rId2*                |
    /// | `target`  | *../media/image2.png* |
    ///
    pub fn link_images(&mut self) {
        let id_to_target: HashMap<String, String> = self
            .images
            .iter()
            .map(|img_ref| (img_ref.id.clone(), img_ref.target.clone()))
            .collect();

        for element in &mut self.elements {
            if let SlideElement::Image(img_ref, _pos) = element
                && let Some(target) = id_to_target.get(&img_ref.id)
            {
                img_ref.target = target.clone();
            }
        }
        for block in &mut self.blocks {
            if let SlideBlockContent::Image(image) = &mut block.content
                && let Some(target) = id_to_target.get(&image.reference.id)
            {
                image.reference.target = target.clone();
                image.mime_type = mime_type_from_path(target).map(str::to_string);
            } else if let SlideBlockContent::Image(image) = &mut block.content {
                image.mime_type = mime_type_from_path(&image.reference.target).map(str::to_string);
            }
        }
    }

    /// Extracts the file extension from image paths
    pub fn get_image_extension(&self, path: &str) -> String {
        Path::new(path)
            .extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or("bin")
            .to_string()
    }

    /// Compresses the image data and returning it as a `jpg` byte slice
    ///
    /// # Parameter
    ///
    /// - `image_data`: The raw image data as a byte array
    ///
    /// # Returns
    ///
    /// - `Vec<u8>`: Returns the compressed and converted jpg byte array
    ///
    /// # Notes
    ///
    /// All images will be converted to `jpg`
    pub fn compress_image(&self, image_data: &[u8]) -> Option<Vec<u8>> {
        let img = match image::load_from_memory(image_data) {
            Ok(image) => image,
            Err(_) => return None,
        };

        let mut output = Vec::new();
        let quality = self.config.quality;

        if JpegEncoder::new_with_quality(&mut output, quality)
            .encode_image(&img)
            .is_ok()
        {
            Some(output)
        } else {
            None
        }
    }

    pub fn load_images_manually(&self) -> Option<Vec<ManualImage>> {
        let mut images: Vec<ManualImage> = Vec::new();

        let image_refs: Vec<&ImageReference> = self
            .elements
            .iter()
            .filter_map(|element| match element {
                SlideElement::Image(img, _pos) => Some(img),
                _ => None,
            })
            .collect();

        for image_ref in image_refs {
            if let Some(image_data) = self.image_data.get(&image_ref.id) {
                let image_data = if self.config.compress_images {
                    self.compress_image(image_data)
                } else {
                    Some(image_data.clone())
                };

                let base64_str = general_purpose::STANDARD.encode(image_data?);

                let image = ManualImage::new(base64_str, image_ref.clone());
                images.push(image);
            }
        }

        Some(images)
    }

    fn path_to_file_url(&self, path: &Path) -> Option<String> {
        let abs_path = path.canonicalize().ok()?;
        let mut path_str = abs_path.to_string_lossy().replace('\\', "/");

        // remove windows unc prefix
        if cfg!(windows) {
            if let Some(stripped) = path_str.strip_prefix("//?/") {
                path_str = stripped.to_string();
            }
            Some(format!("file:///{}", path_str))
        } else {
            Some(format!("file://{}", path_str))
        }
    }
}

pub(crate) fn legacy_blocks(elements: &[SlideElement]) -> Vec<SlideBlock> {
    elements
        .iter()
        .enumerate()
        .map(|(source_order, element)| legacy_block(element, source_order))
        .collect()
}

pub(crate) fn legacy_block(element: &SlideElement, source_order: usize) -> SlideBlock {
    let (bounds, content) = match element {
        SlideElement::Text(text, position) => (
            (*position).into(),
            SlideBlockContent::Text(TextBlock {
                role: TextRole::Other,
                paragraphs: vec![Paragraph::plain(text.runs.clone())],
            }),
        ),
        SlideElement::List(list, position) => (
            (*position).into(),
            SlideBlockContent::Text(TextBlock {
                role: TextRole::Body,
                paragraphs: list
                    .items
                    .iter()
                    .map(|item| Paragraph {
                        runs: item.runs.clone(),
                        alignment: Default::default(),
                        list: Some(ListInfo {
                            level: item.level,
                            kind: if item.is_ordered {
                                ListKind::Ordered {
                                    style: None,
                                    start: 1,
                                }
                            } else {
                                ListKind::Bullet { character: None }
                            },
                        }),
                        list_explicit: true,
                    })
                    .collect(),
            }),
        ),
        SlideElement::Table(table, position) => (
            (*position).into(),
            SlideBlockContent::Table(SemanticTable {
                rows: table
                    .rows
                    .iter()
                    .map(|row| SemanticTableRow {
                        cells: row
                            .cells
                            .iter()
                            .map(|cell| SemanticTableCell {
                                paragraphs: vec![Paragraph::plain(cell.runs.clone())],
                                row_span: 1,
                                column_span: 1,
                                covered: false,
                            })
                            .collect(),
                    })
                    .collect(),
            }),
        ),
        SlideElement::Image(image, position) => (
            (*position).into(),
            SlideBlockContent::Image(ImageBlock {
                reference: image.clone(),
                alt_text: None,
                mime_type: None,
            }),
        ),
        SlideElement::Unknown => (
            Bounds::default(),
            SlideBlockContent::Unsupported(UnsupportedBlock {
                kind: "unknown".to_string(),
                fallback_text: None,
            }),
        ),
    };
    SlideBlock {
        bounds,
        source_order,
        content,
    }
}

fn ordered_blocks(blocks: &[SlideBlock], reading_order: ReadingOrder) -> Vec<&SlideBlock> {
    if reading_order == ReadingOrder::Source {
        let mut ordered: Vec<_> = blocks
            .iter()
            .filter(|block| !block_is_semantically_empty(block))
            .collect();
        ordered.sort_by_key(|block| block.source_order);
        return ordered;
    }

    let has_dimensions = blocks
        .iter()
        .any(|block| block.bounds.width > 0 || block.bounds.height > 0);
    if !has_dimensions {
        let mut ordered: Vec<_> = blocks
            .iter()
            .filter(|block| !block_is_semantically_empty(block))
            .collect();
        ordered.sort_by_key(|block| {
            (
                role_priority(block),
                block.bounds.y,
                block.bounds.x,
                block.source_order,
            )
        });
        return ordered;
    }

    let mut ordered = Vec::with_capacity(blocks.len());
    let mut remaining: Vec<_> = blocks
        .iter()
        .filter(|block| !block_is_semantically_empty(block))
        .collect();
    remaining.sort_by_key(|block| block.source_order);
    for priority in [0, 1] {
        let mut index = 0;
        while index < remaining.len() {
            if role_priority(remaining[index]) == priority {
                ordered.push(remaining.remove(index));
            } else {
                index += 1;
            }
        }
    }

    let left = remaining
        .iter()
        .map(|block| block.bounds.x)
        .min()
        .unwrap_or(0);
    let right = remaining
        .iter()
        .map(|block| block.bounds.x + block.bounds.width)
        .max()
        .unwrap_or(left);
    let page_width = (right - left).max(1);
    let mut separators: Vec<_> = remaining
        .iter()
        .copied()
        .filter(|block| block.bounds.width * 100 >= page_width * 65)
        .collect();
    separators.sort_by_key(|block| (block.bounds.y, block.source_order));

    let mut last_y = i64::MIN;
    for separator in separators {
        let mut band: Vec<_> = remaining
            .iter()
            .copied()
            .filter(|block| {
                !std::ptr::eq(*block, separator)
                    && block.bounds.y >= last_y
                    && block.bounds.y < separator.bounds.y
            })
            .collect();
        sort_spatial_band(&mut band);
        ordered.extend(band);
        ordered.push(separator);
        last_y = separator
            .bounds
            .y
            .saturating_add(separator.bounds.height.max(1));
    }
    let mut tail: Vec<_> = remaining
        .into_iter()
        .filter(|block| {
            block.bounds.y >= last_y && !ordered.iter().any(|item| std::ptr::eq(*item, *block))
        })
        .collect();
    sort_spatial_band(&mut tail);
    ordered.extend(tail);
    ordered
}

fn role_priority(block: &SlideBlock) -> u8 {
    match &block.content {
        SlideBlockContent::Text(TextBlock {
            role: TextRole::Title,
            ..
        }) => 0,
        SlideBlockContent::Text(TextBlock {
            role: TextRole::Subtitle,
            ..
        }) => 1,
        _ => 2,
    }
}

fn block_is_semantically_empty(block: &SlideBlock) -> bool {
    matches!(
        &block.content,
        SlideBlockContent::Text(text)
            if text
                .paragraphs
                .iter()
                .all(|paragraph| paragraph.runs.iter().all(|run| run.text.is_empty()))
    )
}

fn sort_spatial_band(blocks: &mut Vec<&SlideBlock>) {
    blocks.sort_by_key(|block| (block.bounds.x, block.bounds.y, block.source_order));
}

fn render_text_block(output: &mut String, text: &TextBlock) {
    let mut counters: HashMap<u32, u32> = HashMap::new();
    for (index, paragraph) in text.paragraphs.iter().enumerate() {
        let context = if paragraph.list.is_some() {
            MarkdownContext::ListItem
        } else {
            MarkdownContext::Flow
        };
        let mut rendered = render_runs(&paragraph.runs, context);
        if context == MarkdownContext::Flow || context == MarkdownContext::Quote {
            if rendered.ends_with('\n') {
                rendered.pop();
            }
        } else if rendered.ends_with("<br>") {
            rendered.truncate(rendered.len() - "<br>".len());
        }
        if let Some(list) = &paragraph.list {
            counters.retain(|level, _| *level <= list.level);
            let indent = "\t".repeat(list.level as usize);
            let marker = match &list.kind {
                ListKind::Bullet { .. } => "- ".to_string(),
                ListKind::Ordered { start, .. } => {
                    let counter = counters.entry(list.level).or_insert(*start);
                    let marker = format!("{}. ", *counter);
                    *counter += 1;
                    marker
                }
            };
            output.push_str(&indent);
            output.push_str(&marker);
            output.push_str(&rendered);
            output.push('\n');
            continue;
        }

        counters.clear();
        let prefix = match text.role {
            TextRole::Title => "## ",
            TextRole::Heading => "### ",
            _ => "",
        };
        if text.role == TextRole::Subtitle && !rendered.is_empty() {
            output.push('_');
            output.push_str(&rendered);
            output.push('_');
        } else {
            output.push_str(prefix);
            output.push_str(&rendered);
        }
        if index + 1 < text.paragraphs.len() {
            output.push_str("\n\n");
        } else {
            output.push('\n');
        }
    }
}

fn render_table(output: &mut String, table: &SemanticTable) {
    let complex = table.rows.iter().flat_map(|row| &row.cells).any(|cell| {
        cell.row_span > 1 || cell.column_span > 1 || cell.covered || cell.paragraphs.len() > 1
    });
    if complex {
        output.push_str("<table>\n");
        for row in &table.rows {
            output.push_str("  <tr>");
            for cell in &row.cells {
                if cell.covered {
                    continue;
                }
                let mut attributes = String::new();
                if cell.row_span > 1 {
                    attributes.push_str(&format!(" rowspan=\"{}\"", cell.row_span));
                }
                if cell.column_span > 1 {
                    attributes.push_str(&format!(" colspan=\"{}\"", cell.column_span));
                }
                let value = cell
                    .paragraphs
                    .iter()
                    .map(|paragraph| render_runs(&paragraph.runs, MarkdownContext::TableCell))
                    .collect::<Vec<_>>()
                    .join("<br>");
                output.push_str(&format!("<td{attributes}>{value}</td>"));
            }
            output.push_str("</tr>\n");
        }
        output.push_str("</table>\n\n");
        return;
    }

    for (row_index, row) in table.rows.iter().enumerate() {
        let cells = row
            .cells
            .iter()
            .map(|cell| {
                cell.paragraphs
                    .iter()
                    .map(|paragraph| render_runs(&paragraph.runs, MarkdownContext::TableCell))
                    .collect::<Vec<_>>()
                    .join("<br>")
            })
            .collect::<Vec<_>>();
        output.push_str(&format!("| {} |\n", cells.join(" | ")));
        if row_index == 0 {
            output.push_str(&format!("|{}|\n", vec![" --- "; cells.len()].join("|")));
        }
    }
    output.push('\n');
}

fn missing_image_markdown(image: &ImageBlock) -> String {
    let label = image
        .alt_text
        .as_deref()
        .or_else(|| image.reference.target.split('/').next_back())
        .unwrap_or("image");
    format!("[Image unavailable: {label}]")
}

fn mime_type_from_path(path: &str) -> Option<&'static str> {
    match Path::new(path)
        .extension()
        .and_then(|extension| extension.to_str())?
        .to_ascii_lowercase()
        .as_str()
    {
        "jpg" | "jpeg" => Some("image/jpeg"),
        "png" => Some("image/png"),
        "gif" => Some("image/gif"),
        "svg" => Some("image/svg+xml"),
        "webp" => Some("image/webp"),
        "bmp" => Some("image/bmp"),
        "tif" | "tiff" => Some("image/tiff"),
        _ => None,
    }
}

fn append_quoted_section(output: &mut String, title: &str, elements: &[crate::TextElement]) {
    if !output.is_empty() && !output.ends_with("\n\n") {
        output.push('\n');
    }
    output.push_str(&format!("> **{}**\n>\n", title));
    for (index, element) in elements.iter().enumerate() {
        let content = render_runs(&element.runs, MarkdownContext::Quote);
        for line in content.lines() {
            output.push_str("> ");
            output.push_str(line);
            output.push('\n');
        }
        if index + 1 < elements.len() {
            output.push_str(">\n");
        }
    }
}

#[cfg(test)]
#[path = "../tests/unit/slide.rs"]
mod tests;
