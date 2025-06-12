//! Working with the parsed slide elements example for the pptx-to-md crate
//!
//! This example demonstrates how to use the slide elements before the elements are parsed to Markdown,
//! to do your conversion logic.
//!
//! Run with: cargo run --example slide_elements <path/to/your/presentation.pptx>

use pptx_to_md::{ParserConfig, PptxContainer, Result, SlideElement};
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
    let config = ParserConfig::builder()
        .extract_images(true)
        .build();
    
    // Open the PPTX file with the streaming API
    let mut streamer = PptxContainer::open(Path::new(pptx_path), config)?;

    // Process slides one by one using the iterator
    for slide_result in streamer.iter_slides() {
        match slide_result {
            Ok(slide) => {
                println!("Processing slide {} ({} elements)", slide.slide_number, slide.elements.len());

                // iterate over each slide element and match them to add custom logic
                for element in &slide.elements {
                    match element {
                        SlideElement::Text(text) => { println!("{:?}\n", text) }
                        SlideElement::Table(table) => { println!("{:?}\n", table) }
                        SlideElement::Image(image) => { println!("{:?}\n", image) }
                        SlideElement::List(list) => { println!("{:?}\n", list) }
                        SlideElement::Unknown => { println!("An Unknown element was found.\n") }
                    }
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