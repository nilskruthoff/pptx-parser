//! Inspect the semantic document model instead of rendering Markdown directly.
//!
//! Run with:
//! cargo run --example slide_elements <presentation.pptx|presentation.odp>

use pptx_to_md::{ParserConfig, PresentationContainer, Result, SlideBlockContent};
use std::env;
use std::path::Path;

fn main() -> Result<()> {
    let args: Vec<String> = env::args().collect();
    let Some(input_path) = args.get(1) else {
        eprintln!("Usage: cargo run --example slide_elements <presentation.pptx|presentation.odp>");
        return Ok(());
    };

    let mut presentation =
        PresentationContainer::open(Path::new(input_path), ParserConfig::default())?;
    let document = presentation.parse_document()?;

    println!(
        "Parsed {:?} presentation with {} slides and {} diagnostics",
        presentation.format(),
        document.slides.len(),
        document.diagnostics.len()
    );

    for slide in &document.slides {
        println!(
            "Slide {} ({} semantic blocks)",
            slide.slide_number,
            slide.blocks.len()
        );
        for block in &slide.blocks {
            match &block.content {
                SlideBlockContent::Text(text) => {
                    println!("  {:?} text at {:?}: {:?}", text.role, block.bounds, text)
                }
                SlideBlockContent::Table(table) => {
                    println!("  Table at {:?}: {:?}", block.bounds, table)
                }
                SlideBlockContent::Image(image) => {
                    println!("  Image at {:?}: {:?}", block.bounds, image)
                }
                SlideBlockContent::Unsupported(unsupported) => {
                    println!("  Unsupported at {:?}: {:?}", block.bounds, unsupported)
                }
            }
        }
    }

    Ok(())
}
