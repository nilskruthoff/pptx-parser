//! Basic usage example for the pptx-to-md crate
//!
//! This example demonstrates how to open a PPTX file and convert all slides to Markdown.
//!
//! Run with: cargo run --example save_images <path/to/your/presentation.pptx>

use pptx_to_md::{ImageHandlingMode, ParserConfig, PptxContainer, Result};
use std::fs::File;
use std::io::Write;
use std::path::{Path, PathBuf};
use std::env;

fn main() -> Result<()> {
    // Get the PPTX file path from command line arguments and provide the mandatory output path
    let args: Vec<String> = env::args().collect();
    let pptx_path = if args.len() > 1 {
        &args[1]
    } else {
        eprintln!("Usage: cargo run --example save_images <path/to/presentation.pptx>");
        return Ok(());
    };

    println!("Processing PPTX file: {}", pptx_path);

    // Use the config builder to build your config
    let config = ParserConfig::builder()
        .extract_images(true)
        .compress_images(true)
        .quality(75)
        .image_handling_mode(ImageHandlingMode::Save)
        .image_output_path(PathBuf::from("C:/Users/nilsk/Downloads/extracted_images"))
        .build();

    // Open the PPTX file
    let mut container = PptxContainer::open(Path::new(pptx_path), config)?;

    // Parse all slides
    let slides = container.parse_all()?;

    println!("Found {} slides", slides.len());

    // create a new Markdown file
    let mut md_file = File::create("output.md")?;

    // Convert each slide to Markdown and save the images automatically
    for slide in slides {
        if let Some(md_content) = slide.convert_to_md() {
            writeln!(md_file, "{}", md_content).expect("Couldn't write to file");
        }
    }

    println!("All slides converted successfully!");

    Ok(())
}