use super::*;

#[test]
fn run_extract_returns_the_unmodified_text() {
    let run = Run {
        text: "Text with <markup> & whitespace".to_string(),
        formatting: Formatting::default(),
        link_target: Some("https://example.com".to_string()),
    };

    assert_eq!(run.extract(), "Text with <markup> & whitespace");
}
