//! Legacy PPTX-only API retained for existing integrations.
//!
//! New code should use `PresentationContainer`; see `basic_usage.rs`.
//!
//! Run with:
//! cargo run --example legacy_pptx_api <presentation.pptx> [output.md]

use pptx_to_md::{ParserConfig, PptxContainer, Result};
use std::env;
use std::fs;
use std::path::Path;

fn main() -> Result<()> {
    let args: Vec<String> = env::args().collect();
    let Some(input_path) = args.get(1) else {
        eprintln!("Usage: cargo run --example legacy_pptx_api <presentation.pptx> [output.md]");
        return Ok(());
    };
    let output_path = args.get(2).map(String::as_str).unwrap_or("output.md");

    // This is the pre-PresentationContainer flow: open a PPTX-specific
    // container, parse every slide, and render each slide separately. It does
    // not add the presentation-level metadata header.
    let mut container = PptxContainer::open(Path::new(input_path), ParserConfig::default())?;
    let markdown = container
        .parse_all()?
        .into_iter()
        .map(|slide| slide.convert_to_md())
        .collect::<Result<Vec<_>>>()?
        .join("\n");
    fs::write(output_path, markdown)?;

    println!("Converted PPTX with the legacy entry point to {output_path}");
    Ok(())
}
