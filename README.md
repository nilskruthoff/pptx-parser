# pptx-to-md

[![Crates.io](https://img.shields.io/crates/v/pptx-to-md.svg)](https://crates.io/crates/pptx-to-md)
[![tests](https://github.com/nilskruthoff/pptx-parser/actions/workflows/rust.yml/badge.svg)](https://github.com/nilskruthoff/pptx-parser/actions/workflows/rust.yml)
[![Documentation](https://docs.rs/pptx-to-md/badge.svg)](https://docs.rs/pptx-to-md)
![License](https://img.shields.io/crates/l/pptx-to-md.svg)

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
- 🪄 **Embedding:** Used to provide pptx content and meta information in a form that is useful for embeddings
---

## 👨‍💻 Example Usage

`PresentationContainer` is the recommended entry point for new code. It detects whether the input is a PowerPoint (`.pptx`) or OpenDocument Presentation (`.odp`) file and exposes the same parsing API for both formats.

```rust
use pptx_to_md::{ImageHandlingMode, ParserConfig, PresentationContainer, SlideElement};
use std::path::Path;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Create config instance with the `ParserConfigBuilder` 
    // this example is equivalent to the `ParseConfig::default()`
    let config = ParserConfig::builder()
        .extract_images(true)
        .compress_images(true)
        .quality(80)
        .image_handling_mode(ImageHandlingMode::InMarkdown)
        .image_output_path(None)
        .include_slide_comment(true)
        .build();

    let mut container = PresentationContainer::open(
        Path::new("path/to/your/presentation.pptx"), // or .odp
        config,
    )?;

    println!("Detected format: {:?}", container.format());

    let slides = container.parse_all()?; // or `parse_all_multi_threaded()?`

    for slide in slides {
        if let Some(md_content) = slide.convert_to_md() {
            println!("{}", md_content);
        }

        for element in &slide.elements {
            match element {
                SlideElement::Text(text, position) => println!("{:?} at {:?}\n", text, position),
                SlideElement::Table(table, position) => println!("{:?} at {:?}\n", table, position),
                SlideElement::Image(image_reference, position) => {
                    println!("{:?} at {:?}\n", image_reference, position)
                }
                SlideElement::List(list, position) => println!("{:?} at {:?}\n", list, position),
                SlideElement::Unknown => println!("An unknown element was found.\n"),
            }
        }
    }

    Ok(())
}
```

For ODP-specific usage, see [`examples/basic_odp_usage.rs`](https://github.com/nilskruthoff/pptx-parser/tree/master/examples/basic_odp_usage.rs).
For PPTX-only code, `PptxContainer` remains available for backwards compatibility and for callers that explicitly want the old PowerPoint-only entry point. New code that may handle both formats should prefer `PresentationContainer`.
For more usage examples, refer to the [examples](https://github.com/nilskruthoff/pptx-parser/tree/master/examples) directory.

### PPTX and ODP auto-detection

Use `PresentationContainer` when the input may be either `.pptx` or `.odp`:

```rust
use pptx_to_md::{ParserConfig, PresentationContainer, PresentationFormat};
use std::path::Path;

let mut presentation = PresentationContainer::open(
    Path::new("path/to/presentation.odp"),
    ParserConfig::default(),
)?;

assert_eq!(presentation.format(), PresentationFormat::Odp);
let slides = presentation.parse_all()?;
# Ok::<(), pptx_to_md::Error>(())
```

If the format is already known, use `open_as` to skip auto-detection:

```rust
use pptx_to_md::{ParserConfig, PresentationContainer, PresentationFormat};
use std::path::Path;

let mut presentation = PresentationContainer::open_as(
    Path::new("path/to/presentation.pptx"),
    ParserConfig::default(),
    PresentationFormat::Pptx,
)?;

let slides = presentation.parse_all()?;
# Ok::<(), pptx_to_md::Error>(())
```

---

## Config Parameters

| Parameter                | Type                  | Default       | Description                                                                                               |
|--------------------------|-----------------------|---------------|-----------------------------------------------------------------------------------------------------------|
| `extract_images`         | `bool`                | `true`        | Whether images are extracted from slides or not. If false, images can not be extracted manually either.   |
| `compress_images`        | `bool`                | `true`        | Whether images are compressed before encoding or not. Effects manually extracted images too.              |
| `image_quality`          | `u8`                  | `80`          | Defines the image compression quality `(0-100)`. Higher values mean better quality but larger file sizes. |
| `image_handling_mode`    | `ImageHandlingMode`   | `InMarkdown`  | Determines how images are handled during content export                                                   |
| `image_output_path`      | `Option<PathBuf>`     | `None`        | Output directory path for `ImageHandlingMode::Save` (mandatory for saving mode)                           |
| `include_slide_comment`  | `bool`                | `true`        | Weather the slide number comment is included or not (`<!-- Slide [n] -->`)                                | 
<br/>

#### Member of `ImageHandlingMode`
| Member        | Description                                                                                                                     |
|---------------|---------------------------------------------------------------------------------------------------------------------------------|
| `InMarkdown`  | Images are embedded directly in the Markdown output using standard syntax as `base64` data (`![]()`)                            |            
| `Manually`    | Image handling is delegated to the user, requiring manual copying or referencing (as `base64`)                                  |
| `Save`        | Images will be saved in a provided output directory and integrated using `<a>` tag syntax (`<a href="file:///<abs_path>"></a>`) |            

---

## 🏗 Project Structure
```
pptx-to-md/
├── Cargo.toml
├── README.md
├── CHANGELOG.md
├── LICENSE-MIT
├── LICENSE-APACHE
├── examples/           # Simple examples to present the usage of this crate
│   ├── basic_usage.rs
│   ├── basic_odp_usage.rs
│   ├── manual_image_extraction.rs
│   ├── memory_efficient_streaming.rs
│   ├── performance_tests.rs
│   ├── save_images.rs
│   └── slide_elements.rs
├── src/
│   ├── lib.rs            # Public API
│   ├── container.rs      # Pptx container handling
│   ├── presentation.rs   # Format-detecting presentation container
│   ├── odp.rs            # ODP container handling
│   ├── parser_config.rs  # Config and config builder
│   ├── slide.rs          # Individual slide representation & markdown conversion
│   ├── parse_xml.rs      # XML parsing logic
│   ├── parse_rels.rs     # Relationship parsing logic
│   └── types.rs          # Common data types used
├── tests/
│   ├── test_data/      # XML & MD test data files
└── └── slide_tests.rs  # tests for md conversion logic
```

---

## 📦 Installation

Include the following line in your Cargo.toml dependencies section:

```toml
[dependencies]
pptx-to-md = "0.4.0"
```

---

## 📜 License
This project is licensed under the [MIT-License](https://github.com/nilskruthoff/pptx-parser/blob/master/LICENCE-MIT)
and [Apache 2.0-Licence](https://github.com/nilskruthoff/pptx-parser/blob/master/LICENSE-APACHE).

Feel free to contribute or suggest improvements!

---
