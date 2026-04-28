//! Two-level diff: paragraph alignment + word-level within matched pairs.
//!
//! This version is tuned for legal-document comparison:
//! - Move detection is disabled to avoid noisy false "moved" paragraphs.
//! - Paragraph replacement pairing prefers legal clause/list markers.
//! - Word-level diff still preserves whitespace/punctuation tokens.

use serde::Serialize;
use similar::{ChangeTag, DiffOp, TextDiff};

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
#[serde(rename_all = "lowercase")]
pub enum ChangeKind {
    Equal,
    Insert,
    Delete,
}

#[derive(Debug, Clone, Serialize)]
pub struct Change {
    pub kind: ChangeKind,
    pub text: String,
}

/// Paragraph-level operation.
#[derive(Debug, Clone, Serialize)]
#[serde(tag = "kind", rename_all = "lowercase")]
pub enum ParaOp {
    /// Paragraph identical in old and new.
    Equal { text: String },
    /// Paragraph present only in new.
    Insert { text: String },
    /// Paragraph present only in old.
    Delete { text: String },
    /// Paragraph present in both but content changed — word-level diff inside.
    Modified { changes: Vec<Change> },
    /// Disabled in current matching strategy, but retained for API compatibility.
    MovedFrom { text: String, move_id: u32 },
    /// Disabled in current matching strategy, but retained for API compatibility.
    MovedTo { text: String, move_id: u32 },
}

/// Word-level diff of two strings.
pub fn diff_words(old: &str, new: &str) -> Vec<Change> {
    let old_tokens = crate::token::tokenize_words(old);
    let new_tokens = crate::token::tokenize_words(new);

    let diff = TextDiff::from_slices(&old_tokens, &new_tokens);

    let mut out = Vec::new();
    for op in diff.ops() {
        for change in diff.iter_changes(op) {
            let kind = match change.tag() {
                ChangeTag::Equal => ChangeKind::Equal,
                ChangeTag::Insert => ChangeKind::Insert,
                ChangeTag::Delete => ChangeKind::Delete,
            };
            push_or_extend(&mut out, kind, change.value());
        }
    }
    out
}

/// Two-level diff: align paragraphs first, then diff words inside matched pairs.
pub fn diff_paragraphs(old: &[String], new: &[String]) -> Vec<ParaOp> {
    let old_refs: Vec<&str> = old.iter().map(String::as_str).collect();
    let new_refs: Vec<&str> = new.iter().map(String::as_str).collect();
    let diff = TextDiff::from_slices(&old_refs, &new_refs);

    let mut out = Vec::new();
    for op in diff.ops() {
        match *op {
            DiffOp::Equal { old_index, len, .. } => {
                for i in 0..len {
                    out.push(ParaOp::Equal {
                        text: old[old_index + i].clone(),
                    });
                }
            }
            DiffOp::Insert { new_index, new_len, .. } => {
                for i in 0..new_len {
                    out.push(ParaOp::Insert {
                        text: new[new_index + i].clone(),
                    });
                }
            }
            DiffOp::Delete { old_index, old_len, .. } => {
                for i in 0..old_len {
                    out.push(ParaOp::Delete {
                        text: old[old_index + i].clone(),
                    });
                }
            }
            DiffOp::Replace {
                old_index,
                old_len,
                new_index,
                new_len,
            } => {
                emit_clause_aware_replace(
                    &old[old_index..old_index + old_len],
                    &new[new_index..new_index + new_len],
                    &mut out,
                );
            }
        }
    }

    // Intentionally do NOT run move detection.
    // Legal redlines become much noisier when repeated boilerplate paragraphs
    // are incorrectly classified as moves.
    out
}

/// Minimum similarity for two paragraphs to be treated as a modified pair.
const SIMILARITY_THRESHOLD: f32 = 0.42;

/// Stronger threshold when no legal clause/list key matches.
const WEAK_PAIR_THRESHOLD: f32 = 0.58;

fn emit_clause_aware_replace(old_slice: &[String], new_slice: &[String], out: &mut Vec<ParaOp>) {
    if old_slice.is_empty() {
        for n in new_slice {
            out.push(ParaOp::Insert { text: n.clone() });
        }
        return;
    }
    if new_slice.is_empty() {
        for o in old_slice {
            out.push(ParaOp::Delete { text: o.clone() });
        }
        return;
    }

    let mut used_new = vec![false; new_slice.len()];

    for old_text in old_slice {
        let mut best: Option<(usize, f32)> = None;
        let old_key = legal_anchor_key(old_text);

        for (j, new_text) in new_slice.iter().enumerate() {
            if used_new[j] {
                continue;
            }

            let new_key = legal_anchor_key(new_text);
            let sim = similarity(old_text, new_text);

            let score = if !old_key.is_empty() && old_key == new_key {
                // Same clause/list marker: strongly prefer pairing these.
                sim + 0.35
            } else if !old_key.is_empty() && !new_key.is_empty() && old_key != new_key {
                // Different legal markers should usually not be paired.
                sim - 0.25
            } else {
                sim
            };

            if best.map_or(true, |(_, best_score)| score > best_score) {
                best = Some((j, score));
            }
        }

        if let Some((j, score)) = best {
            let base_sim = similarity(old_text, &new_slice[j]);
            let old_key = legal_anchor_key(old_text);
            let new_key = legal_anchor_key(&new_slice[j]);
            let key_match = !old_key.is_empty() && old_key == new_key;

            if key_match || score >= WEAK_PAIR_THRESHOLD || base_sim >= SIMILARITY_THRESHOLD {
                used_new[j] = true;
                if old_text == &new_slice[j] {
                    out.push(ParaOp::Equal { text: old_text.clone() });
                } else {
                    out.push(ParaOp::Modified {
                        changes: diff_words(old_text, &new_slice[j]),
                    });
                }
            } else {
                out.push(ParaOp::Delete {
                    text: old_text.clone(),
                });
            }
        } else {
            out.push(ParaOp::Delete {
                text: old_text.clone(),
            });
        }
    }

    for (j, new_text) in new_slice.iter().enumerate() {
        if !used_new[j] {
            out.push(ParaOp::Insert {
                text: new_text.clone(),
            });
        }
    }
}

fn similarity(a: &str, b: &str) -> f32 {
    if a.is_empty() && b.is_empty() {
        return 1.0;
    }
    let at = crate::token::tokenize_words(a);
    let bt = crate::token::tokenize_words(b);
    TextDiff::from_slices(&at, &bt).ratio()
}

/// Extracts a conservative legal-document anchor from the beginning of a paragraph.
/// Examples:
/// - "제 1 조 (용어의 정의)" -> "article:1"
/// - "(1) ..." -> "paren:1"
/// - "1. ..." -> "num:1"
/// - "가. ..." -> "kor:가"
/// - "① ..." -> "circle:①"
fn legal_anchor_key(text: &str) -> String {
    let t = text.trim_start();
    if t.is_empty() {
        return String::new();
    }

    if let Some(k) = article_key(t) {
        return k;
    }
    if let Some(k) = parenthesized_number_key(t) {
        return k;
    }
    if let Some(k) = dotted_number_key(t) {
        return k;
    }
    if let Some(k) = korean_list_key(t) {
        return k;
    }
    if let Some(k) = circled_number_key(t) {
        return k;
    }

    String::new()
}

fn article_key(t: &str) -> Option<String> {
    let chars: Vec<char> = t.chars().collect();
    if chars.first().copied()? != '제' {
        return None;
    }

    let mut i = 1usize;
    while i < chars.len() && chars[i].is_whitespace() {
        i += 1;
    }

    let start = i;
    while i < chars.len() && (chars[i].is_ascii_digit() || is_korean_number_char(chars[i])) {
        i += 1;
    }

    if i == start {
        return None;
    }

    while i < chars.len() && chars[i].is_whitespace() {
        i += 1;
    }

    if i < chars.len() && chars[i] == '조' {
        let n: String = chars[start..i].iter().collect();
        return Some(format!("article:{n}"));
    }

    None
}

fn parenthesized_number_key(t: &str) -> Option<String> {
    let chars: Vec<char> = t.chars().collect();
    if chars.first().copied()? != '(' {
        return None;
    }

    let mut i = 1usize;
    let start = i;
    while i < chars.len() && chars[i].is_ascii_digit() {
        i += 1;
    }

    if i == start {
        return None;
    }

    if i < chars.len() && chars[i] == ')' {
        let n: String = chars[start..i].iter().collect();
        return Some(format!("paren:{n}"));
    }

    None
}

fn dotted_number_key(t: &str) -> Option<String> {
    let chars: Vec<char> = t.chars().collect();
    let mut i = 0usize;

    while i < chars.len() && chars[i].is_ascii_digit() {
        i += 1;
    }

    if i == 0 {
        return None;
    }

    if i < chars.len() && (chars[i] == '.' || chars[i] == '．') {
        let n: String = chars[..i].iter().collect();
        return Some(format!("num:{n}"));
    }

    None
}

fn korean_list_key(t: &str) -> Option<String> {
    let mut chars = t.chars();
    let first = chars.next()?;
    let second = chars.next();

    if matches!(first, '가' | '나' | '다' | '라' | '마' | '바' | '사' | '아' | '자' | '차' | '카' | '타' | '파' | '하')
        && matches!(second, Some('.') | Some('．') | Some(')'))
    {
        return Some(format!("kor:{first}"));
    }

    None
}

fn circled_number_key(t: &str) -> Option<String> {
    let first = t.chars().next()?;
    if ('①'..='⑳').contains(&first) {
        return Some(format!("circle:{first}"));
    }
    None
}

fn is_korean_number_char(c: char) -> bool {
    matches!(
        c,
        '영' | '공' | '일' | '이' | '삼' | '사' | '오' | '육' | '륙' | '칠' | '팔' | '구'
            | '십' | '백' | '천' | '만'
    )
}

fn push_or_extend(out: &mut Vec<Change>, kind: ChangeKind, token: &str) {
    if let Some(last) = out.last_mut() {
        if last.kind == kind {
            last.text.push_str(token);
            return;
        }
    }
    out.push(Change {
        kind,
        text: token.to_string(),
    });
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn identical() {
        let changes = diff_words("hello world", "hello world");
        assert_eq!(changes.len(), 1);
        assert_eq!(changes[0].kind, ChangeKind::Equal);
    }

    #[test]
    fn simple_replace() {
        let changes = diff_words("hello world", "hello rust");
        let kinds: Vec<_> = changes.iter().map(|c| c.kind.clone()).collect();
        assert!(kinds.contains(&ChangeKind::Delete));
        assert!(kinds.contains(&ChangeKind::Insert));
    }

    #[test]
    fn para_insert() {
        let old = vec!["a".to_string(), "b".to_string()];
        let new = vec!["a".to_string(), "new".to_string(), "b".to_string()];
        let ops = diff_paragraphs(&old, &new);
        assert_eq!(ops.len(), 3);
        assert!(matches!(ops[1], ParaOp::Insert { .. }));
    }

    #[test]
    fn para_delete() {
        let old = vec!["a".to_string(), "gone".to_string(), "b".to_string()];
        let new = vec!["a".to_string(), "b".to_string()];
        let ops = diff_paragraphs(&old, &new);
        assert!(ops.iter().any(|o| matches!(o, ParaOp::Delete { .. })));
    }

    #[test]
    fn para_modified() {
        let old = vec!["제 1 조 (정의) hello world".to_string()];
        let new = vec!["제 1 조 (정의) hello rust".to_string()];
        let ops = diff_paragraphs(&old, &new);
        assert!(matches!(ops[0], ParaOp::Modified { .. }));
    }

    #[test]
    fn anchor_keys() {
        assert_eq!(legal_anchor_key("제 1 조 (정의)"), "article:1");
        assert_eq!(legal_anchor_key("(2) 내용"), "paren:2");
        assert_eq!(legal_anchor_key("14. 내용"), "num:14");
        assert_eq!(legal_anchor_key("가. 내용"), "kor:가");
        assert_eq!(legal_anchor_key("① 내용"), "circle:①");
    }
}
