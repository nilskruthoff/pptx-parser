//! Basic ODP usage example for the pptx-to-md crate.
//!
//! This example demonstrates how to open an ODP file and convert all slides to Markdown.
//!
//! Run with: cargo run --example basic_odp_usage <path/to/your/presentation.odp> <extract_images>

use pptx_to_md::{ImageHandlingMode, ParserConfig, PresentationContainer, PresentationFormat, Result};
use std::env;
use std::fs::File;
use std::io::Write;
use std::path::Path;

fn main() -> Result<()> {
    let args: Vec<String> = env::args().collect();
    let odp_path = if args.len() > 1 {
        &args[1]
    } else {
        eprintln!("Usage: cargo run --example basic_odp_usage <path/to/presentation.odp> <extract_images>\ncargo run --example basic_odp_usage sample.odp true");
        return Ok(());
    };

    let extract_images = if args.len() > 2 {
        !(args[2] == "false" || args[2] == "False" || args[2] == "0")
    } else {
        true
    };

    println!("Processing ODP file: {}", odp_path);

    let config = ParserConfig::builder()
        .extract_images(extract_images)
        .compress_images(true)
        .quality(75)
        .image_handling_mode(ImageHandlingMode::InMarkdown)
        .include_slide_number_as_comment(true)
        .include_comments(true)
        .include_speaker_notes(true)
        .build();

    let mut container = PresentationContainer::open(Path::new(odp_path), config)?;
    if container.format() != PresentationFormat::Odp {
        eprintln!("Expected an ODP file, detected {:?}", container.format());
        return Ok(());
    }

    let slides = container.parse_all()?;
    println!("Found {} slides", slides.len());

    let mut md_file = File::create("output.md")?;
    for slide in slides {
        if let Some(md_content) = slide.convert_to_md() {
            println!("{}", md_content);
            writeln!(md_file, "{}", md_content).expect("Couldn't write to file");
        }
    }

    println!("All slides converted successfully!");

    Ok(())
}
