//! Memory-efficient streaming example for the pptx-to-md crate
//!
//! This example demonstrates how to use the streaming API to process large
//! PPTX files with minimal memory usage.
//!
//! Run with: cargo run --example memory_efficient_streaming <path/to/your/presentation.pptx>

use std::env;
use std::path::Path;
use std::fs;
use pptx_to_md::{PptxContainer, Result};

fn main() -> Result<()> {
    // Get the PPTX file path from command line arguments
    let args: Vec<String> = env::args().collect();
    let pptx_path = if args.len() > 1 {
        &args[1]
    } else {
        eprintln!("Usage: cargo run --example memory_efficient_streaming <path/to/presentation.pptx>");
        return Ok(());
    };

    println!("Processing PPTX file: {}", pptx_path);

    // Open the PPTX file with the streaming API
    let mut streamer = PptxContainer::open(Path::new(pptx_path))?;
    
    // Create output directory
    let output_dir = "output_streaming";
    fs::create_dir_all(output_dir)?;

    // Process slides one by one using the iterator
    for slide_result in streamer.iter_slides() {
        match slide_result {
            Ok(slide) => {
                println!("Processing slide {} ({} elements)", slide.slide_number, slide.elements.len());

                if let Some(md_content) = slide.convert_to_md() {
                    let output_path = format!("{}/slide_{}.md", output_dir, slide.slide_number);
                    fs::write(&output_path, md_content)?;
                    println!("Saved slide {} to {}", slide.slide_number, output_path);
                }
            },
            Err(e) => {
                eprintln!("Error processing slide: {:?}", e);
            }
        }
    }

    println!("All slides processed successfully!");

    Ok(())
}