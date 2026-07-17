use crate::{Baseline, Run};

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

    if run.formatting.strikethrough {
        result = format!("~~{result}~~");
    }

    result = match run.formatting.baseline {
        Baseline::Normal => result,
        Baseline::Superscript => format!("<sup>{result}</sup>"),
        Baseline::Subscript => format!("<sub>{result}</sub>"),
    };

    if let Some(target) = &run.link_target {
        result = format!("[{result}]({})", markdown_link_destination(target));
    }

    format!("{leading_whitespace}{result}{trailing_whitespace}")
}

fn markdown_link_destination(target: &str) -> String {
    if target
        .chars()
        .any(|character| character.is_whitespace() || matches!(character, '(' | ')' | '<' | '>'))
    {
        format!(
            "<{}>",
            target
                .replace('<', "%3C")
                .replace('>', "%3E")
                .replace('\n', "%0A")
                .replace('\r', "%0D")
        )
    } else {
        target.to_string()
    }
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
        && left.formatting.strikethrough == right.formatting.strikethrough
        && left.formatting.baseline == right.formatting.baseline
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
#[path = "../tests/unit/markdown.rs"]
mod tests;
