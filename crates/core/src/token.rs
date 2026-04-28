//! Tokenization. Word-level splitter that keeps whitespace and punctuation
//! as separate tokens so diff output can be reassembled faithfully.

/// Split text into word/whitespace/punct tokens. Unicode-aware.
pub fn tokenize_words(text: &str) -> Vec<&str> {
    let mut out = Vec::new();
    let mut start = 0;
    let mut prev: Option<CharClass> = None;

    for (i, c) in text.char_indices() {
        let class = classify(c);
        if let Some(p) = prev {
            if p != class || class == CharClass::Punct {
                out.push(&text[start..i]);
                start = i;
            }
        }
        prev = Some(class);
    }
    if start < text.len() {
        out.push(&text[start..]);
    }
    out
}

#[derive(Copy, Clone, PartialEq, Eq)]
enum CharClass {
    Word,
    Space,
    Punct,
}

fn classify(c: char) -> CharClass {
    if c.is_whitespace() {
        CharClass::Space
    } else if c.is_alphanumeric() || c == '_' {
        CharClass::Word
    } else {
        CharClass::Punct
    }
}
