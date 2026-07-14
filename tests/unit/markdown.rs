use super::*;
use crate::Formatting;

fn run(text: &str) -> Run {
    Run {
        text: text.to_string(),
        formatting: Formatting::default(),
        link_target: None,
    }
}

#[test]
fn escapes_inline_markdown_characters_in_one_pass() {
    assert_eq!(
        render_runs(&[run(r"text \\ ` * _ [ ] ~ < &")], MarkdownContext::Flow),
        r"text \\\\ \` \* \_ \[ \] \~ \< \&"
    );
}

#[test]
fn escapes_block_markers_only_at_line_start() {
    let input = "# heading\n> quote\n- item\n12. item\n---\n===\nplain # hash";
    assert_eq!(
        render_runs(&[run(input)], MarkdownContext::Flow),
        "\\# heading\n\n\\> quote\n\n\\- item\n\n12\\. item\n\n\\---\n\n\\===\n\nplain # hash"
    );
}

#[test]
fn escapes_pipes_only_in_table_cells() {
    assert_eq!(
        render_runs(&[run("left | right")], MarkdownContext::TableCell),
        r"left \| right"
    );
    assert_eq!(
        render_runs(&[run("left | right")], MarkdownContext::Flow),
        "left | right"
    );
}

#[test]
fn normalizes_line_breaks_and_preserves_flow_structure() {
    assert_eq!(
        render_runs(&[run("first\r\nsecond\rthird\n")], MarkdownContext::Flow),
        "first\n\nsecond\n\nthird\n"
    );
    assert_eq!(
        render_runs(&[run("first\n")], MarkdownContext::TableCell),
        "first"
    );
}

#[test]
fn preserves_line_start_state_across_runs() {
    assert_eq!(
        render_runs(
            &[run("  "), run("# heading\n"), run("12. item")],
            MarkdownContext::Flow
        ),
        "  \\# heading\n\n12\\. item"
    );
}

#[test]
fn public_run_renderer_keeps_one_structural_trailing_break() {
    assert_eq!(run("first\nsecond\n").render_as_md(), "first\n\nsecond\n");
}

#[test]
fn escapes_text_before_adding_formatting_and_links() {
    let linked = Run {
        text: "*label*".to_string(),
        formatting: Formatting {
            bold: true,
            ..Formatting::default()
        },
        link_target: Some("https://example.com/a_(b)".to_string()),
    };
    assert_eq!(
        render_runs(&[linked], MarkdownContext::Flow),
        r"[**\*label\***](https://example.com/a_(b))"
    );
}

#[test]
fn merges_adjacent_runs_with_the_same_markdown_style() {
    let formatting = Formatting {
        bold: true,
        ..Formatting::default()
    };
    let runs = [
        Run {
            text: "Bold".to_string(),
            formatting: formatting.clone(),
            link_target: None,
        },
        Run {
            text: " ".to_string(),
            formatting: formatting.clone(),
            link_target: None,
        },
        Run {
            text: "text\n".to_string(),
            formatting,
            link_target: None,
        },
    ];
    assert_eq!(render_runs(&runs, MarkdownContext::Flow), "**Bold text**\n");
}

#[test]
fn does_not_merge_runs_with_different_links() {
    let linked = |target: &str, text: &str| Run {
        text: text.to_string(),
        formatting: Formatting::default(),
        link_target: Some(target.to_string()),
    };
    assert_eq!(
        render_runs(
            &[
                linked("https://example.com/one", "one"),
                linked("https://example.com/two", "two")
            ],
            MarkdownContext::Flow
        ),
        "[one](https://example.com/one)[two](https://example.com/two)"
    );
}

#[test]
fn keeps_boundary_whitespace_outside_formatting_markers() {
    let formatted = |text: &str| Run {
        text: text.to_string(),
        formatting: Formatting {
            italic: true,
            ..Formatting::default()
        },
        link_target: None,
    };
    assert_eq!(
        render_runs(
            &[formatted(" leading and trailing ")],
            MarkdownContext::Flow
        ),
        " _leading and trailing_ "
    );
    assert_eq!(render_runs(&[formatted(" ")], MarkdownContext::Flow), " ");
}
