//! Process one PPTX or ODP slide at a time instead of retaining all slides.
//!
//! Run with:
//! cargo run --example memory_efficient_streaming <presentation.pptx|presentation.odp>

use pptx_to_md::{ParserConfig, PresentationContainer, Result};
use std::env;
use std::fs;
use std::path::Path;

fn main() -> Result<()> {
    let args: Vec<String> = env::args().collect();
    let Some(input_path) = args.get(1) else {
        eprintln!(
            "Usage: cargo run --example memory_efficient_streaming <presentation.pptx|presentation.odp>"
        );
        return Ok(());
    };

    let mut presentation =
        PresentationContainer::open(Path::new(input_path), ParserConfig::default())?;
    let output_dir = "output_streaming";
    fs::create_dir_all(output_dir)?;

    // Unlike parse_document(), the iterator only retains the current slide.
    for slide_result in presentation.iter_slides() {
        let slide = slide_result?;
        let output_path = format!("{output_dir}/slide_{}.md", slide.slide_number);
        fs::write(&output_path, slide.convert_to_md()?)?;
        println!(
            "Saved slide {} ({} semantic blocks) to {output_path}",
            slide.slide_number,
            slide.blocks.len()
        );
    }

    Ok(())
}
