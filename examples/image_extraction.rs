//! Image extraction example for the pptx-to-md crate
//!
//! This example demonstrates how to extract images from a PPTX file.
//!
//! Run with: cargo run --example image_extraction <path/to/your/presentation.pptx>

use std::env;
use std::path::Path;
use std::fs;
use pptx_to_md::{PptxContainer, SlideElement, Result, ParserConfig};

fn main() -> Result<()> {
    // Get the PPTX file path from command line arguments
    let args: Vec<String> = env::args().collect();
    let pptx_path = if args.len() > 1 {
        &args[1]
    } else {
        eprintln!("Usage: cargo run --example image_extraction <path/to/presentation.pptx>");
        return Ok(());
    };

    println!("Extracting images from PPTX file: {}", pptx_path);

    // Use the config builder to build your config
    let config = ParserConfig::builder()
        .extract_images(true)
        .build();
    
    // Open the PPTX file with the streaming API
    let mut streamer = PptxContainer::open(Path::new(pptx_path), config)?;

    // Create output directory
    let output_dir = "extracted_images";
    fs::create_dir_all(output_dir)?;

    // Process slides one by one using the iterator
    let mut image_count = 0;

    for slide_result in streamer.iter_slides() {
        match slide_result {
            Ok(slide) => {
                // Find image elements in the slide
                for (element_idx, element) in slide.elements.iter().enumerate() {
                    if let SlideElement::Image(img_ref) = element {
                        // Get image data from the slide's image_data HashMap
                        if let Some(image_data) = slide.image_data.get(&img_ref.id) {
                            // Determine image extension from target
                            let extension = img_ref.target
                                .split('.')
                                .next_back()
                                .unwrap_or("bin");

                            // Save the image
                            let output_path = format!(
                                "{}/slide{}_image{}_{}.{}",
                                output_dir,
                                slide.slide_number,
                                element_idx,
                                &img_ref.id,
                                extension
                            );

                            fs::write(&output_path, image_data)?;
                            println!("Saved image to {}", output_path);
                            image_count += 1;
                        }
                    }
                }
            },
            Err(e) => {
                eprintln!("Error processing slide: {:?}", e);
            }
        }
    }

    println!("Extracted {} images successfully!", image_count);

    Ok(())
}