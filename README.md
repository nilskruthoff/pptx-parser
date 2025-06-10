# pptx-to-md

[![Crates.io](https://img.shields.io/crates/v/pptx-to-md.svg)](https://crates.io/crates/pptx-to-md)
[![Documentation](https://docs.rs/pptx-to-md/badge.svg)](https://docs.rs/pptx-to-md)
![License](https://img.shields.io/crates/l/pptx-to-md.svg)

`pptx-to-md` is a library to parse Microsoft PowerPoint (`.pptx`) slides and convert them into structured Markdown content and data, making it easy to process, use, or integrate slide data programmatically.

---

## 🚀 Features

- 📄 **Extract Slide Text:** Parses and extracts text elements from slides.
- 📋 **Lists & Tables:** Recognizes and formats lists (ordered/unordered) and tables into Markdown.
- 🖼️ **Embedded Images:** Supports embedded images extraction as base64-encoded inline images.
- ⚙️ **Robust & Safe APIs:** Designed according to Rust best practices with explicit error handling.
- 🧑‍💻 **Developer-Friendly:** Simple API design, extensive documentation, and examples.
- 🪄 **Embedding:** Used to provide pptx content and meta information in a form that is useful for embeddings

---

## 📦 Installation

Include the following line in your Cargo.toml dependencies section:

```toml
[dependencies]
pptx-to-md = "0.1" # replace with the current version
```

---

## 👨‍💻 Example Usage

Here's an easy example to convert a PowerPoint slide into Markdown:

```rust
use pptx_to_md::PptxContainer;
use std::path::Path;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = Path::new("presentation.pptx");
    let pptx_container = PptxContainer::open(path)?;
    let slides = pptx_container.parse()?;
    
    
    /// access each slide present in the pptx container
    for slide in slides {
        /// convert slide content to markdown
        if let Some(md) = slide.convert_to_md() {
            println!("Slide {}: \n{}", slide.slide_number, md);
        }

        /// Access the `SlideElements` containing the parsed xml
        for element in slide_elements {
            match element {
                SlideElement::Text(text) => println!("Text element: {:?}", text),
                SlideElement::Table(table) => println!("Table element: {:?}", table),
                SlideElement::Image(image) => println!("Image element: {:?}", image),
                SlideElement::List(list) => println!("List element: {:?}", list),
                SlideElement::Unknown => println!("Unknown or unsupported element detected"),
            }
        }
    }

    Ok(())
}
```

---

## 🏗 Project Structure
```
pptx-to-md/
├── Cargo.toml
├── README.md
├── src/
│   ├── lib.rs          # Public API
│   ├── container.rs    # Pptx container handling
│   ├── slide.rs        # Individual slide representation & markdown conversion
│   ├── parse_xml.rs    # XML parsing logic
│   └── types.rs        # Common data types used
├── tests/
│   ├── test_data/      # XML & MD test data files
└── └── slide_tests.rs  # tests for md conversion logic
```

---

## 📜 License
This project is licensed under the [MIT-License](https://github.com/nilskruthoff/pptx-parser/blob/master/LICENCE-MIT)
and [Apache 2.0-Licence](https://github.com/nilskruthoff/pptx-parser/blob/master/LICENSE-APACHE).

Feel free to contribute or suggest improvements!

---

