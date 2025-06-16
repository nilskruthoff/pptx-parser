# pptx-to-md

[![Crates.io](https://img.shields.io/crates/v/pptx-to-md.svg)](https://crates.io/crates/pptx-to-md)
[![tests](https://github.com/nilskruthoff/pptx-parser/actions/workflows/rust.yml/badge.svg)](https://github.com/nilskruthoff/pptx-parser/actions/workflows/rust.yml)
[![Documentation](https://docs.rs/pptx-to-md/badge.svg)](https://docs.rs/pptx-to-md)
![License](https://img.shields.io/crates/l/pptx-to-md.svg)

`pptx-to-md` is a library to parse Microsoft PowerPoint (`.pptx`) slides and convert them into structured Markdown content and data, making it easy to process, use, or integrate slide data programmatically.

---

## 🚀 Features

- 📄 **Extract Slide Text:** Parses and extracts text elements from slides.
- 📋 **Lists & Tables:** Recognizes and formats lists (ordered/unordered) and tables into Markdown.
- 🖼️ **Embedded Images:** Supports embedded images extraction as base64-encoded inline images.
- 💾 **Memory Efficient**: Use the streaming API to iterate over one slide at a time, never overloading memory.
- ⏱️ **Multithreading**: Optional support for multithreaded parsing of PowerPoint slides, with a significant performance increase for larger presentations.
- ⚙️ **Robust & Safe APIs:** Designed according to Rust best practices with explicit error handling.
- 🪄 **Embedding:** Used to provide pptx content and meta information in a form that is useful for embeddings
---

## 👨‍💻 Example Usage

Here's an easy example to convert a PowerPoint slide into Markdown*:

```rust
use pptx_to_md::{PptxContainer, ParserConfig};
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
        .build();
    // alternatively use `let config = ParserConfig::default();`
    
    // open the container with the path to your .pptx file
    let pptx_container = PptxContainer::open(Path::new("path/to/your/presentation.pptx"), config)?;
    
    // Parse all slides' xml at once single- or multithreaded
    let slides = container.parse_all()?; // or `parse_all_multi_threaded()?`
    
    for slide in slides {
        // Convert each slide into Markdown
        if let Some(md_content) = slide.convert_to_md() {
            println!("{}", md_content);
        }

        // Or iterate over each slide element and match them to add custom logic
        for element in &slide.elements {
            match element {
                SlideElement::Text(text) => { println!("{:?}\n", text) }
                SlideElement::Table(table) => { println!("{:?}\n", table) }
                SlideElement::Image(image_reference) => { println!("{:?}\n", image_reference) }
                SlideElement::List(list) => { println!("{:?}\n", list) }
                SlideElement::Unknown => { println!("An Unknown element was found.\n") }
            }
        }
    }

    Ok(())
}
```

*for more usage examples refer to the [examples](https://github.com/nilskruthoff/pptx-parser/tree/master/examples) directory

---

## Config Parameters

| Parameter              | Type                  | Default       | Description                                                                                               |
|------------------------|-----------------------|---------------|-----------------------------------------------------------------------------------------------------------|
| `extract_images`       | `bool`                | `true`        | Whether images are extracted from slides or not. If false, images can not be extracted manually either.   |
| `compress_images`      | `bool`                | `true`        | Whether images are compressed before encoding or not. Effects manually extracted images too.              |
| `image_quality`        | `u8`                  | `80`          | Defines the image compression quality `(0-100)`. Higher values mean better quality but larger file sizes. |
| `image_handling_mode`  | `ImageHandlingMode`   | `InMarkdown`  | Determines how images are handled during content export                                                   |
| `image_output_path`    | `Option<PathBuf>`     | `None`        | Output directory path for `ImageHandlingMode::Save` (mandatory for saving mode)                           |

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
│   ├── manual_image_extraction.rs
│   ├── memory_efficient_streaming.rs
│   ├── performance_tests.rs
│   ├── save_images.rs
│   └── slide_elements.rs
├── src/
│   ├── lib.rs            # Public API
│   ├── container.rs      # Pptx container handling
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
pptx-to-md = "0.3.0"
```

---

## 📜 License
This project is licensed under the [MIT-License](https://github.com/nilskruthoff/pptx-parser/blob/master/LICENCE-MIT)
and [Apache 2.0-Licence](https://github.com/nilskruthoff/pptx-parser/blob/master/LICENSE-APACHE).

Feel free to contribute or suggest improvements!

---