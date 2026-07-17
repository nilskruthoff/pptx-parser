//! Save presentation images next to the generated Markdown instead of embedding them.
//!
//! Run with:
//! cargo run --example save_images <presentation.pptx|presentation.odp> [image-directory] [output.md]

use pptx_to_md::{ImageHandlingMode, ParserConfig, PresentationContainer, Result};
use std::env;
use std::fs;
use std::path::{Path, PathBuf};

fn main() -> Result<()> {
    let args: Vec<String> = env::args().collect();
    let Some(input_path) = args.get(1) else {
        eprintln!(
            "Usage: cargo run --example save_images <presentation.pptx|presentation.odp> [image-directory] [output.md]"
        );
        return Ok(());
    };
    let image_directory = args
        .get(2)
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("extracted_images"));
    let output_path = args.get(3).map(String::as_str).unwrap_or("output.md");

    let config = ParserConfig::builder()
        .image_handling_mode(ImageHandlingMode::Save)
        .image_output_path(image_directory)
        .build();
    let mut presentation = PresentationContainer::open(Path::new(input_path), config)?;

    fs::write(output_path, presentation.convert_to_md()?)?;
    println!("Saved Markdown to {output_path}");
    Ok(())
}
