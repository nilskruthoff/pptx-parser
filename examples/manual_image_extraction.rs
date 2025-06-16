//! Basic usage example for the pptx-to-md crate
//!
//! This example demonstrates how to open a PPTX file and convert all slides to Markdown.
//!
//! Run with: cargo run --example manual_image_extraction <path/to/your/presentation.pptx>

use pptx_to_md::{PptxContainer, Result, ParserConfig, ImageHandlingMode};
use std::{env, fs};
use std::fs::File;
use std::io::Write;
use std::path::Path;
use base64::Engine;
use base64::engine::general_purpose;

fn main() -> Result<()> {
    // Get the PPTX file path from command line arguments
    let args: Vec<String> = env::args().collect();
    let pptx_path = if args.len() > 1 {
        &args[1]
    } else {
        eprintln!("Usage: cargo run --example manual_image_extraction <path/to/presentation.pptx>");
        return Ok(());
    };

    println!("Processing PPTX file: {}", pptx_path);

    // Use the config builder to build your config
    let config = ParserConfig::builder()
        .extract_images(true)
        .compress_images(true)
        .quality(75)
        .image_handling_mode(ImageHandlingMode::Manually)
        .build();

    // Open the PPTX file
    let mut container = PptxContainer::open(Path::new(pptx_path), config)?;

    // Parse all slides
    let slides = container.parse_all()?;

    println!("Found {} slides", slides.len());

    // create a new Markdown file
    let mut md_file = File::create("output.md")?;

    // Create output directory
    let output_dir = "extracted_images";
    fs::create_dir_all(output_dir)?;

    // Process slides one by one using the iterator
    let mut image_count = 1;

    // Convert each slide to Markdown and save
    for slide in slides {
        if let Some(md_content) = slide.convert_to_md() {
            writeln!(md_file, "{}", md_content).expect("Couldn't write to file");
        }
        
        // Manually load the base64 encoded image strings from the slide
        if let Some(images) = slide.load_images_manually() {
            for image in images {
                
                // Decode the base64 strings back to raw image data
                let image_data = general_purpose::STANDARD.decode(image.base64_content.clone()).unwrap();

                // Extract image extension if the image is not compressed, otherwise its always `.jpg`
                let ext = slide.config.compress_images
                    .then(|| "jpg".to_string())
                    .unwrap_or_else(|| slide.get_image_extension(&image.img_ref.target.clone()));

                // Construct a unique file name
                let file_name = format!("slide{}_image{}_{}", slide.slide_number, image_count, &image.img_ref.id);
                
                // Save the image
                let output_path = format!(
                    "{}/{}.{}",
                    output_dir,
                    &file_name,
                    ext
                );
                fs::write(&output_path, image_data)?;
                println!("Saved image to {}", output_path);

                // Write the image data into the Markdown file
                writeln!(md_file, "![{}](data:image/{};base64,{})", file_name, ext, image.base64_content).expect("Couldn't write to file");
                
                image_count += 1;
            }
        }
    }

    println!("All slides converted successfully!");

    Ok(())
}