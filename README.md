[![Crates.io](https://img.shields.io/crates/v/pptx-to-md.svg)](https://crates.io/crates/pptx-to-md)
[![tests](https://github.com/nilskruthoff/pptx-parser/actions/workflows/rust.yml/badge.svg)](https://github.com/nilskruthoff/pptx-parser/actions/workflows/rust.yml)
[![codecov](https://codecov.io/gh/nilskruthoff/pptx-parser/branch/main/graph/badge.svg)](https://codecov.io/gh/nilskruthoff/pptx-parser)
![License](https://img.shields.io/crates/l/pptx-to-md.svg)
[![dependency status](https://deps.rs/repo/github/nilskruthoff/pptx-parser/status.svg)](https://deps.rs/repo/github/nilskruthoff/pptx-parser)
[![Documentation](https://docs.rs/pptx-to-md/badge.svg)](https://docs.rs/pptx-to-md)
![Crates.io Downloads](https://img.shields.io/crates/d/pptx-to-md)

# pptx-to-md

`pptx-to-md` is a library to parse Microsoft PowerPoint (`.pptx`) slides and convert them into structured Markdown content and data, making it easy to process, use, or integrate slide data programmatically.
It also supports OpenDocument Presentation (`.odp`).

---

## 🚀 Features

- 🖥️ **Compatibility:** Supports `.pptx` and `.odp` files
- 📄 **Extract Slide Text:** Parses and extracts text elements from slides.
- 📋 **Lists & Tables:** Recognizes and formats lists (ordered/unordered) and tables into Markdown.
- 🖼️ **Embedded Images:** Supports embedded images extraction as base64-encoded inline images.
- 💾 **Memory Efficient**: Use the streaming API to iterate over one slide at a time, never overloading memory.
- ⏱️ **Multithreading**: Optional support for multithreaded parsing of PowerPoint slides, with a significant performance increase for larger presentations.
- ⚙️ **Robust & Safe APIs:** Designed according to Rust best practices with explicit error handling.
- 🏷️ **Presentation Metadata:** Extracts common document properties from PPTX and ODP.
- 🪄 **Embedding:** Used to provide pptx content and meta-information in a form that is useful for embeddings
---

## 📌 Quickstart

`PresentationContainer` is the preferred entry point for all new code. It
detects PPTX and ODP automatically. The default configuration produces complete
Markdown with presentation metadata, slide markers, spatial reading order, and
embedded images.

```rust
use pptx_to_md::{ParserConfig, PresentationContainer, Result};
use std::path::Path;

fn main() -> Result<()> {
    let mut presentation = PresentationContainer::open(
        Path::new("presentation.pptx"), // also accepts .odp
        ParserConfig::default(),
    )?;

    let markdown = presentation.convert_to_md()?;
    std::fs::write("output.md", markdown)?;
    Ok(())
}
```

See [`examples/basic_usage.rs`](examples/basic_usage.rs) for the corresponding
command-line example.

## 🎯 Choosing the right API

Start with `PresentationContainer` and select the operation based on what the
application needs:

| Goal | Preferred API | What it does |
| --- | --- | --- |
| Convert one complete presentation | `convert_to_md()` | Parses every slide and returns one Markdown document, including optional presentation metadata |
| Convert a large PPTX faster | `convert_to_md_multi_threaded()` | Uses parallel slide parsing for PPTX; ODP transparently uses its normal document parser |
| Inspect or transform structured content | `parse_document()` | Returns metadata, semantic slides and blocks, and aggregated diagnostics |
| Work with all slides directly | `parse_all()` | Returns `Vec<Slide>` without creating presentation-level Markdown |
| Process one slide at a time | `iter_slides()` | Streams `Result<Slide>` values and avoids retaining every parsed slide |
| Customize one slide's Markdown | `Slide::to_markdown(&MarkdownOptions)` | Controls reading order, slide marker, notes, comments, and unsupported-content comments |
| Read only document properties | `metadata()` | Returns parsed presentation metadata without parsing the slides |

`ParserConfig` controls parsing, image handling, and the defaults used by
`convert_to_md()`. `MarkdownOptions` is only needed when rendering an individual
parsed slide differently.

Important behavior differences:

- Presentation-level `convert_to_md*()` methods parse and render the slides and
  emit presentation metadata once. Per-slide methods never emit presentation
  metadata.
- `parse_document()` and `parse_all*()` retain all returned slides. The former
  additionally packages metadata and aggregates slide diagnostics.
- `Slide::convert_to_md()` uses the rendering flags copied from `ParserConfig`.
  `Slide::to_markdown()` accepts explicit `MarkdownOptions` for that one call;
  image loading and image output mode still come from the slide's
  `ParserConfig`.
- `PresentationContainer::open()` performs format detection. Use `format()` to
  inspect the result. `open_as()` is only necessary when the caller explicitly
  wants to force `PresentationFormat::Pptx` or `PresentationFormat::Odp`.

### Structured content and custom Markdown

Use `parse_document()` when Markdown is not the only desired output or parser
diagnostics need to be inspected:

```rust
use pptx_to_md::{MarkdownOptions, ParserConfig, PresentationContainer, ReadingOrder, Result};
use std::path::Path;

fn main() -> Result<()> {
    let mut presentation = PresentationContainer::open(
        Path::new("presentation.odp"),
        ParserConfig::default(),
    )?;
    let document = presentation.parse_document()?;

    for diagnostic in &document.diagnostics {
        eprintln!("{:?}: {}", diagnostic.severity, diagnostic.message);
    }

    let options = MarkdownOptions {
        reading_order: ReadingOrder::Source,
        include_speaker_notes: true,
        ..MarkdownOptions::default()
    };
    let first_slide_markdown = document.slides[0].to_markdown(&options)?;
    println!("{first_slide_markdown}");
    Ok(())
}
```

`ReadingOrder::Spatial` is the default and orders visual columns heuristically.
`ReadingOrder::Source` preserves the order in the source XML.

### Streaming and parallel parsing

Use `iter_slides()` for bounded-memory processing. Each slide is parsed when the
iterator advances. Use `parse_all_multi_threaded()` or
`convert_to_md_multi_threaded()` for CPU-parallel PPTX parsing. ODP stores its
pages in one `content.xml`, so its implementation remains sequential.

---

## ⚙️ Config Parameters

| Parameter                | Type                  | Default       | Description                                                                                               |
|--------------------------|-----------------------|---------------|-----------------------------------------------------------------------------------------------------------|
| `extract_images`         | `bool`                | `true`        | Whether images are extracted from slides or not. If false, images can not be extracted manually either.   |
| `compress_images`        | `bool`                | `true`        | Whether images are compressed before encoding or not. Effects manually extracted images too.              |
| `quality`               | `u8`                  | `80`          | Defines the image compression quality `(0-100)`. Higher values mean better quality but larger file sizes. |
| `image_handling_mode`    | `ImageHandlingMode`   | `InMarkdown`  | Determines how images are handled during content export                                                   |
| `image_output_path`      | `Option<PathBuf>`     | `None`        | Output directory path for `ImageHandlingMode::Save` (mandatory for saving mode)                           |
| `include_slide_number_as_comment`  | `bool`                | `true`        | Whether the slide number comment is included (`<!-- Slide [n] -->`)                                       |
| `include_speaker_notes`  | `bool`                | `false`       | Whether speaker notes are appended to Markdown as blockquotes                                             |
| `include_comments`       | `bool`                | `false`       | Whether presentation comments are appended to Markdown as blockquotes                                     |
| `include_presentation_metadata` | `bool`       | `true`        | Whether complete-presentation Markdown starts with a metadata HTML comment                                 |
<br/>

#### Member of `ImageHandlingMode`
| Member        | Description                                                                                                                     |
|---------------|---------------------------------------------------------------------------------------------------------------------------------|
| `InMarkdown`  | Images are embedded directly in the Markdown output using standard syntax as `base64` data (`![]()`)                            |            
| `Manually`    | Image handling is delegated to the user, requiring manual copying or referencing (as `base64`)                                  |
| `Save`        | Images are saved in the configured output directory and referenced with Markdown image syntax and a `file://` URL              |

---

## 📦 Installation

Include the following line in your Cargo.toml dependencies section:

```toml
[dependencies]
pptx-to-md = "1.0.0"
```

---

## 📜 License
This project is licensed under the [MIT-License](https://github.com/nilskruthoff/pptx-parser/blob/master/LICENCE-MIT)
and [Apache 2.0-Licence](https://github.com/nilskruthoff/pptx-parser/blob/master/LICENSE-APACHE).

Feel free to contribute or suggest improvements! 😊

---
