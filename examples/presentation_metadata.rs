//! Read structured presentation metadata without converting slides to Markdown.
//!
//! Run with: cargo run --example presentation_metadata <presentation.pptx|presentation.odp>

use pptx_to_md::{ParserConfig, PresentationContainer, Result};
use std::env;
use std::path::Path;

fn main() -> Result<()> {
    let args: Vec<String> = env::args().collect();
    let Some(input_path) = args.get(1) else {
        eprintln!(
            "Usage: cargo run --example presentation_metadata <presentation.pptx|presentation.odp>"
        );
        return Ok(());
    };

    let container = PresentationContainer::open(Path::new(input_path), ParserConfig::default())?;
    let metadata = container.metadata();

    println!("Format: {:?}", container.format());
    println!("Title: {:?}", metadata.title);
    println!("Author: {:?}", metadata.author);
    println!("Last modified by: {:?}", metadata.last_modified_by);
    println!("Subject: {:?}", metadata.subject);
    println!("Description: {:?}", metadata.description);
    println!("Keywords: {:?}", metadata.keywords);
    println!("Created: {:?}", metadata.created_at);
    println!("Modified: {:?}", metadata.modified_at);
    Ok(())
}
