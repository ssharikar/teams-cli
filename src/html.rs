//! Shared plain-text rendering of HTML message/email bodies, used by
//! `cli::mail`, `cli::calendar`, and `cli::message` (Outlook and Teams both
//! render plain-text-style content as HTML on the wire).

/// Tags whose open/close/self-close should become a line break rather than
/// being deleted outright. Without this, a message built from `<br/>`s or
/// `<p>`/`<div>` fields (the common shape of an automated notification
/// email: "Field: value<br/>Field: value<br/>...") collapses into one
/// unbroken line with no way to tell where one field ends and the next
/// begins.
const BLOCK_TAGS: &[&str] = &[
    "p",
    "div",
    "br",
    "hr",
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
    "li",
    "tr",
    "blockquote",
];

/// Strips HTML tags for terminal display, preserving line structure for
/// block-level tags (see `BLOCK_TAGS`) instead of silently deleting it.
pub fn strip_html(html: &str) -> String {
    let mut result = String::new();
    let mut in_tag = false;
    let mut tag_buf = String::new();

    for ch in html.chars() {
        match ch {
            '<' => {
                in_tag = true;
                tag_buf.clear();
            }
            '>' => {
                in_tag = false;
                if is_block_tag(&tag_buf) {
                    result.push('\n');
                }
            }
            _ if in_tag => tag_buf.push(ch),
            _ => result.push(ch),
        }
    }

    collapse_blank_lines(&result).trim().to_string()
}

fn is_block_tag(raw_tag: &str) -> bool {
    let name = raw_tag
        .trim_start_matches('/')
        .trim_start()
        .split(|c: char| c.is_whitespace() || c == '/')
        .next()
        .unwrap_or("")
        .to_ascii_lowercase();
    BLOCK_TAGS.contains(&name.as_str())
}

/// Caps runs of consecutive newlines at 2 (a single blank-line separator),
/// so e.g. adjacent `</p><p>` tags don't pile up extra blank lines.
fn collapse_blank_lines(s: &str) -> String {
    let mut out = String::new();
    let mut run = 0;
    for ch in s.chars() {
        if ch == '\n' {
            run += 1;
            if run <= 2 {
                out.push(ch);
            }
        } else {
            run = 0;
            out.push(ch);
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn strip_html_basic() {
        assert_eq!(strip_html("<p>hello</p>"), "hello");
    }

    #[test]
    fn strip_html_empty() {
        assert_eq!(strip_html(""), "");
    }

    #[test]
    fn strip_html_no_tags() {
        assert_eq!(strip_html("no tags"), "no tags");
    }

    #[test]
    fn strip_html_bold_and_italic() {
        assert_eq!(
            strip_html("<b>bold</b> and <i>italic</i>"),
            "bold and italic"
        );
    }

    #[test]
    fn strip_html_nested_tags() {
        assert_eq!(
            strip_html("<div>nested<span>tags</span></div>"),
            "nestedtags"
        );
    }

    /// A standalone break with no surrounding content trims away entirely —
    /// expected, since there's nothing for the line break to separate.
    #[test]
    fn strip_html_self_closing_tag_alone() {
        assert_eq!(strip_html("<br/>"), "");
    }

    /// The actual bug this fixes: line breaks between real content must
    /// survive stripping, not collapse into one run-on line.
    #[test]
    fn strip_html_self_closing_tag_separates_content() {
        assert_eq!(strip_html("a<br/>b<hr/>c"), "a\nb\nc");
    }

    #[test]
    fn strip_html_angle_brackets_in_text() {
        // "a < b > c" - the '<' starts a "tag", ' b ' is treated as tag content,
        // '>' ends the "tag", then ' c' is text. So we lose ' b '.
        // The result after trim is "a  c" (with the space before '<' and after '>').
        let result = strip_html("a < b > c");
        assert_eq!(result, "a  c");
    }

    #[test]
    fn strip_html_unclosed_tag() {
        // "unclosed <div tag" - '<' starts a tag, everything after is inside the tag
        // because there's no closing '>'. So only "unclosed " is kept, trimmed to "unclosed".
        let result = strip_html("unclosed <div tag");
        assert_eq!(result, "unclosed");
    }

    #[test]
    fn strip_html_complex_html() {
        let html = "<html><body><h1>Title</h1><p>Some <b>bold</b> text</p></body></html>";
        assert_eq!(strip_html(html), "Title\n\nSome bold text");
    }

    #[test]
    fn strip_html_with_attributes() {
        assert_eq!(
            strip_html(r#"<a href="https://example.com">link</a>"#),
            "link"
        );
    }

    #[test]
    fn strip_html_whitespace_only_content() {
        // After stripping tags, only whitespace remains -> trimmed to empty
        assert_eq!(strip_html("<p>  </p>"), "");
    }

    #[test]
    fn strip_html_preserves_inner_whitespace() {
        assert_eq!(strip_html("<p>hello   world</p>"), "hello   world");
    }

    /// Repro shape from a real automated-notification email: a run of
    /// "Field: value" lines joined with <br/>, which previously collapsed
    /// into one unreadable, unparseable line.
    #[test]
    fn strip_html_multiline_field_list() {
        let html = "UserID: AQ6O<br/>UserName OLUSOLA AWODERU<br/>Organization: SKY SOLUTIONS LLC";
        assert_eq!(
            strip_html(html),
            "UserID: AQ6O\nUserName OLUSOLA AWODERU\nOrganization: SKY SOLUTIONS LLC"
        );
    }

    #[test]
    fn strip_html_paragraphs_separated_by_blank_line() {
        assert_eq!(strip_html("<p>line1</p><p>line2</p>"), "line1\n\nline2");
    }
}
