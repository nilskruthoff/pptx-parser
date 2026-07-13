use crate::Run;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(crate) enum MarkdownContext {
    Flow,
    ListItem,
    TableCell,
    Quote,
}

pub(crate) struct MarkdownEscaper {
    context: MarkdownContext,
    at_line_start: bool,
    leading_spaces: usize,
}

impl MarkdownEscaper {
    const LINE_FEED: &'static str = "\n";
    const CARRIAGE_RETURN: char = '\r';
    const CR_LF: &'static str = "\r\n";
    const DOUBLE_BACKSLASH: char = '\\';
    const BACKTICK: char = '`';
    const ASTERISK: char = '*';
    const UNDERSCORE: char = '_';
    const HYPHEN_MINUS: char = '-';
    const HYPHEN_PLUS: char = '+';
    const LEFT_BRACKET: char = '[';
    const RIGHT_BRACKET: char = ']';
    const TILDE: char = '~';
    const LESS_THAN: char = '<';
    const GREATER_THAN: char = '>';
    const AMPERSAND: char = '&';
    const PIPE: char = '|';
    const WHITESPACE: char = ' ';
    const NEWLINE: char = '\n';

    pub(crate) fn new(context: MarkdownContext) -> Self {
        Self {
            context,
            at_line_start: true,
            leading_spaces: 0,
        }
    }

    pub(crate) fn escape(&mut self, input: &str) -> String {
        let normalized = input
            .replace(Self::CR_LF, Self::LINE_FEED)
            .replace(Self::CARRIAGE_RETURN, Self::LINE_FEED);
        let mut escaped = String::with_capacity(normalized.len());

        for line in normalized.split_inclusive(Self::NEWLINE) {
            let (content, has_newline) = line
                .strip_suffix(Self::NEWLINE)
                .map(|content| (content, true))
                .unwrap_or((line, false));
            self.escape_line_content(content, &mut escaped);

            if has_newline {
                escaped.push(Self::NEWLINE);
                self.at_line_start = true;
                self.leading_spaces = 0;
            }
        }

        escaped
    }

    fn escape_line_content(&mut self, content: &str, output: &mut String) {
        let block_escape_index = self.block_escape_index(content);

        for (index, character) in content.char_indices() {
            let should_escape = Some(index) == block_escape_index
                || matches!(
                    character,
                    Self::DOUBLE_BACKSLASH
                        | Self::BACKTICK
                        | Self::ASTERISK
                        | Self::UNDERSCORE
                        | Self::LEFT_BRACKET
                        | Self::RIGHT_BRACKET
                        | Self::TILDE
                        | Self::LESS_THAN
                        | Self::AMPERSAND
                )
                || (character == Self::PIPE && self.context == MarkdownContext::TableCell);

            if should_escape {
                output.push(Self::DOUBLE_BACKSLASH);
            }
            output.push(character);
        }

        if self.at_line_start {
            let spaces = content
                .chars()
                .take_while(|character| *character == Self::WHITESPACE)
                .count();
            if spaces == content.chars().count() {
                self.leading_spaces += spaces;
                if self.leading_spaces > 3 {
                    self.at_line_start = false;
                }
            } else {
                self.at_line_start = false;
            }
        }
    }

    fn block_escape_index(&self, content: &str) -> Option<usize> {
        if !self.at_line_start || self.context == MarkdownContext::TableCell {
            return None;
        }

        let available_spaces = 3usize.saturating_sub(self.leading_spaces);
        let spaces = content
            .chars()
            .take_while(|character| *character == Self::WHITESPACE)
            .count();
        if spaces > available_spaces {
            return None;
        }

        let marker_start = spaces;
        let candidate = &content[marker_start..];
        if candidate.is_empty() {
            return None;
        }

        if starts_heading(candidate)
            || candidate.starts_with(Self::GREATER_THAN)
            || starts_unordered_list(candidate)
            || is_thematic_break(candidate)
            || is_setext_underline(candidate)
        {
            return Some(marker_start);
        }

        ordered_list_delimiter(candidate).map(|delimiter| marker_start + delimiter)
    }
}

pub(crate) fn render_runs(runs: &[Run], context: MarkdownContext) -> String {
    let mut escaper = MarkdownEscaper::new(context);
    let mut rendered = String::new();
    let mut group_run: Option<&Run> = None;
    let mut group_text = String::new();
    let preserve_trailing_line_break = context == MarkdownContext::Flow
        && runs.last().is_some_and(|run| {
            run.text.ends_with(MarkdownEscaper::NEWLINE)
                || run.text.ends_with(MarkdownEscaper::CARRIAGE_RETURN)
        });

    for run in runs {
        let escaped = escaper.escape(&run.text);

        if let Some(previous_run) = group_run {
            if has_same_markdown_style(previous_run, run) {
                group_text.push_str(&escaped);
                continue;
            }

            render_group(&mut rendered, previous_run, &group_text);
            group_text.clear();
        }

        group_run = Some(run);
        group_text.push_str(&escaped);
    }

    if let Some(previous_run) = group_run {
        render_group(&mut rendered, previous_run, &group_text);
    }

    while rendered.ends_with(MarkdownEscaper::NEWLINE) {
        rendered.pop();
    }

    let separator = match context {
        MarkdownContext::Flow => "\n\n",
        MarkdownContext::ListItem | MarkdownContext::TableCell => "<br>",
        MarkdownContext::Quote => MarkdownEscaper::LINE_FEED,
    };
    let mut rendered = rendered.replace(MarkdownEscaper::NEWLINE, separator);
    if preserve_trailing_line_break {
        rendered.push(MarkdownEscaper::NEWLINE);
    }
    rendered
}

pub(crate) fn render_run(run: &Run) -> String {
    render_runs(std::slice::from_ref(run), MarkdownContext::Flow)
}

fn render_fragment(run: &Run, escaped_text: &str) -> String {
    let content_start = escaped_text
        .find(|character: char| !character.is_whitespace())
        .unwrap_or(escaped_text.len());
    let content_end = escaped_text
        .rfind(|character: char| !character.is_whitespace())
        .map(|index| {
            index
                + escaped_text[index..]
                    .chars()
                    .next()
                    .expect("non-whitespace character")
                    .len_utf8()
        })
        .unwrap_or(content_start);

    let leading_whitespace = &escaped_text[..content_start];
    let trailing_whitespace = &escaped_text[content_end..];
    let mut result = escaped_text[content_start..content_end].to_string();

    if result.is_empty() {
        return escaped_text.to_string();
    }

    if run.formatting.bold && run.formatting.italic {
        result = format!("***{result}***");
    } else {
        if run.formatting.bold {
            result = format!("**{result}**");
        }
        if run.formatting.italic {
            result = format!("_{result}_");
        }
    }

    if run.formatting.underlined {
        result = format!("<u>{result}</u>");
    }

    if let Some(target) = &run.link_target {
        result = format!("[{result}]({target})");
    }

    format!("{leading_whitespace}{result}{trailing_whitespace}")
}

fn render_group(output: &mut String, run: &Run, escaped_text: &str) {
    let mut fragments = escaped_text.split(MarkdownEscaper::NEWLINE).peekable();

    while let Some(fragment) = fragments.next() {
        if !fragment.is_empty() {
            output.push_str(&render_fragment(run, fragment));
        }
        if fragments.peek().is_some() {
            output.push(MarkdownEscaper::NEWLINE);
        }
    }
}

fn has_same_markdown_style(left: &Run, right: &Run) -> bool {
    left.formatting.bold == right.formatting.bold
        && left.formatting.italic == right.formatting.italic
        && left.formatting.underlined == right.formatting.underlined
        && left.link_target == right.link_target
}

fn starts_heading(candidate: &str) -> bool {
    let hashes = candidate
        .chars()
        .take_while(|character| *character == '#')
        .count();
    (1..=6).contains(&hashes)
        && candidate[hashes..]
            .chars()
            .next()
            .is_none_or(char::is_whitespace)
}

fn starts_unordered_list(candidate: &str) -> bool {
    matches!(
        candidate.chars().next(),
        Some(MarkdownEscaper::HYPHEN_PLUS | MarkdownEscaper::HYPHEN_MINUS)
    ) && candidate.chars().nth(1).is_some_and(char::is_whitespace)
}

fn ordered_list_delimiter(candidate: &str) -> Option<usize> {
    let digit_count = candidate
        .bytes()
        .take_while(|byte| byte.is_ascii_digit())
        .count();
    if !(1..=9).contains(&digit_count) {
        return None;
    }

    let delimiter = candidate.as_bytes().get(digit_count)?;
    if !matches!(delimiter, b'.' | b')') {
        return None;
    }

    candidate
        .get(digit_count + 1..)
        .and_then(|rest| rest.chars().next())
        .filter(|character| character.is_whitespace())
        .map(|_| digit_count)
}

fn is_thematic_break(candidate: &str) -> bool {
    for marker in ['-', '_', '*'] {
        let marker_count = candidate
            .chars()
            .filter(|character| *character == marker)
            .count();
        if marker_count >= 3
            && candidate
                .chars()
                .all(|character| character == marker || character == ' ' || character == '\t')
        {
            return true;
        }
    }
    false
}

fn is_setext_underline(candidate: &str) -> bool {
    let trimmed = candidate.trim_end_matches([' ', '\t']);
    !trimmed.is_empty() && trimmed.chars().all(|character| character == '=')
}

#[cfg(test)]
mod tests {
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
}
