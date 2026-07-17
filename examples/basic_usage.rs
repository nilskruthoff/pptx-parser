//! Preferred PPTX/ODP to Markdown conversion.
//!
//! Run with: cargo run --example basic_usage <presentation.pptx|presentation.odp> [output.md]

use pptx_to_md::{ParserConfig, PresentationContainer, Result};
use std::env;
use std::fs;
use std::path::Path;

fn main() -> Result<()> {
    let args: Vec<String> = env::args().collect();
    let pptx_path = if args.len() > 1 {
        &args[1]
    } else {
        eprintln!("Usage: cargo run --example basic_usage <presentation.pptx|presentation.odp> <extract_images>\ncargo run --example basic_usage sample.pptx true");
        return Ok(());
    };

    // Tries to read if the extract_images flag is false else set to true
    let extract_images = if args.len() > 2 {
        !(args[2] == "false" || args[2] == "False" || args[2] == "0")
    } else {
        true
    };

    let config = ParserConfig::builder()
        .extract_images(extract_images)
        .include_presentation_metadata(true)
        .include_comments(true)
        .include_speaker_notes(true)
        .build();
    let mut container = PresentationContainer::open(Path::new(pptx_path), config)?;
    let markdown = container.convert_to_md()?;

    fs::write("output.md", markdown)?;
    println!("Converted {:?} presentation to output.md", container.format());
    Ok(())
}
