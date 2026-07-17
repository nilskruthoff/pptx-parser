//! Working with the parsed slide elements example for the pptx-to-md crate
//!
//! This example demonstrates how to use the slide elements before the elements are parsed to Markdown,
//! to do your conversion logic.
//!
//! Run with: cargo run --example slide_elements <path/to/your/presentation.pptx>

use pptx_to_md::{ParserConfig, PptxContainer, Result, SlideBlockContent};
use std::env;
use std::path::Path;

fn main() -> Result<()> {
    // Get the PPTX file path from command line arguments
    let args: Vec<String> = env::args().collect();
    let pptx_path = if args.len() > 1 {
        &args[1]
    } else {
        eprintln!("Usage: cargo run --example slide_elements <path/to/presentation.pptx>");
        return Ok(());
    };

    println!("Processing PPTX file: {}", pptx_path);

    // Use the config builder to build your config
    let config = ParserConfig::builder().extract_images(true).build();

    // Open the PPTX file with the streaming API
    let mut streamer = PptxContainer::open(Path::new(pptx_path), config)?;

    // Process slides one by one using the iterator
    for slide_result in streamer.iter_slides() {
        match slide_result {
            Ok(slide) => {
                println!(
                    "Processing slide {} ({} semantic blocks)",
                    slide.slide_number,
                    slide.blocks.len()
                );

                for block in &slide.blocks {
                    match &block.content {
                        SlideBlockContent::Text(text) => {
                            println!("{:?}\t{:?}\n", text, block.bounds)
                        }
                        SlideBlockContent::Table(table) => {
                            println!("{:?}\t{:?}\n", table, block.bounds)
                        }
                        SlideBlockContent::Image(image) => {
                            println!("{:?}\t{:?}\n", image, block.bounds)
                        }
                        SlideBlockContent::Unsupported(unsupported) => {
                            println!("Unsupported: {:?}\n", unsupported)
                        }
                    }
                }
            }
            Err(e) => {
                eprintln!("Error processing slide: {:?}", e);
            }
        }
    }

    println!("All slides processed successfully!");

    Ok(())
}
