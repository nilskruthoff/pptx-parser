use super::*;

#[test]
fn default_config_uses_documented_values() {
    let config = ParserConfig::default();

    assert!(config.extract_images);
    assert!(config.compress_images);
    assert_eq!(config.quality, 80);
    assert_eq!(config.image_handling_mode, ImageHandlingMode::InMarkdown);
    assert_eq!(config.image_output_path, None);
    assert!(config.include_slide_number_as_comment);
    assert!(!config.include_speaker_notes);
    assert!(!config.include_comments);
    assert!(config.include_presentation_metadata);
}

#[test]
fn builder_applies_every_override() {
    let output_path = PathBuf::from("custom-images");
    let config = ParserConfig::builder()
        .extract_images(false)
        .compress_images(false)
        .quality(42)
        .image_handling_mode(ImageHandlingMode::Save)
        .image_output_path(output_path.clone())
        .include_slide_number_as_comment(false)
        .include_speaker_notes(true)
        .include_comments(true)
        .include_presentation_metadata(false)
        .build();

    assert!(!config.extract_images);
    assert!(!config.compress_images);
    assert_eq!(config.quality, 42);
    assert_eq!(config.image_handling_mode, ImageHandlingMode::Save);
    assert_eq!(config.image_output_path, Some(output_path));
    assert!(!config.include_slide_number_as_comment);
    assert!(config.include_speaker_notes);
    assert!(config.include_comments);
    assert!(!config.include_presentation_metadata);
}
