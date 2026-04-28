//! Two-level diff: paragraph alignment + word-level within matched pairs.

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
    /// Original location of a paragraph that was moved to a new position.
    /// `move_id` pairs it with the matching MovedTo.
    MovedFrom { text: String, move_id: u32 },
    /// New location of a paragraph that was moved from elsewhere.
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

/// Minimum similarity ratio for a Delete+Insert pair to be reclassified
/// as a Move. Tuned conservatively — only very close matches are treated
/// as moves, so we never fabricate a move where two paragraphs happen to
/// share common words.
const MOVE_THRESHOLD: f32 = 0.85;

/// Scan the op list for Delete/Insert pairs whose content is ~identical and
/// convert them into MovedFrom/MovedTo pairs sharing a move_id.
fn detect_moves(ops: Vec<ParaOp>) -> Vec<ParaOp> {
    // Collect indices of Deletes and Inserts.
    let mut del_idx: Vec<usize> = Vec::new();
    let mut ins_idx: Vec<usize> = Vec::new();
    for (i, op) in ops.iter().enumerate() {
        match op {
            ParaOp::Delete { .. } => del_idx.push(i),
            ParaOp::Insert { .. } => ins_idx.push(i),
            _ => {}
        }
    }
    if del_idx.is_empty() || ins_idx.is_empty() {
        return ops;
    }

    // Greedy matching: for each Delete, find the best unmatched Insert.
    // O(D*I) but D and I are small (paragraphs added/removed, not entire doc).
    let mut out = ops;
    let mut used_ins = vec![false; ins_idx.len()];
    let mut next_move_id: u32 = 1;

    for &di in &del_idx {
        let del_text = match &out[di] {
            ParaOp::Delete { text } => text.clone(),
            _ => continue,
        };
        let mut best: Option<(usize, f32)> = None;
        for (k, &ii) in ins_idx.iter().enumerate() {
            if used_ins[k] {
                continue;
            }
            let ins_text = match &out[ii] {
                ParaOp::Insert { text } => text,
                _ => continue,
            };
            let s = similarity(&del_text, ins_text);
            if s >= MOVE_THRESHOLD && best.map_or(true, |(_, bs)| s > bs) {
                best = Some((k, s));
            }
        }
        if let Some((k, _)) = best {
            used_ins[k] = true;
            let ii = ins_idx[k];
            let move_id = next_move_id;
            next_move_id += 1;
            // Preserve original text from each side.
            if let (ParaOp::Delete { text: dt }, ParaOp::Insert { text: it }) =
                (out[di].clone(), out[ii].clone())
            {
                out[di] = ParaOp::MovedFrom { text: dt, move_id };
                out[ii] = ParaOp::MovedTo { text: it, move_id };
            }
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
                    out.push(ParaOp::Equal { text: old[old_index + i].clone() });
                }
            }
            DiffOp::Insert { new_index, new_len, .. } => {
                for i in 0..new_len {
                    out.push(ParaOp::Insert { text: new[new_index + i].clone() });
                }
            }
            DiffOp::Delete { old_index, old_len, .. } => {
                for i in 0..old_len {
                    out.push(ParaOp::Delete { text: old[old_index + i].clone() });
                }
            }
            DiffOp::Replace {
                old_index,
                old_len,
                new_index,
                new_len,
            } => {
                // Pair up overlapping positions. For each pair, check similarity:
                // if too low, prefer Delete+Insert over forcing a Modified pairing.
                let paired = old_len.min(new_len);
                for i in 0..paired {
                    let o = &old[old_index + i];
                    let n = &new[new_index + i];
                    if should_pair_as_modified(o, n) {
                        out.push(ParaOp::Modified { changes: diff_words(o, n) });
                    } else {
                        out.push(ParaOp::Delete { text: o.clone() });
                        out.push(ParaOp::Insert { text: n.clone() });
                    }
                }
                for i in paired..old_len {
                    out.push(ParaOp::Delete { text: old[old_index + i].clone() });
                }
                for i in paired..new_len {
                    out.push(ParaOp::Insert { text: new[new_index + i].clone() });
                }
            }
        }
    }
    // Move detection is intentionally disabled for legal-document comparison.
    out
}

/// Minimum word-level similarity for two paragraphs to be considered a "Modified pair"
/// rather than a clean Delete+Insert. Tuned empirically: above ~0.35 the word-level
/// redline stays readable; below that the cost of showing inline diff exceeds the
/// benefit of side-by-side matching.
const SIMILARITY_THRESHOLD: f32 = 0.25;


fn should_pair_as_modified(a: &str, b: &str) -> bool {
    if same_legal_anchor(a, b) {
        return true;
    }
    similarity(a, b) >= SIMILARITY_THRESHOLD
}

fn same_legal_anchor(a: &str, b: &str) -> bool {
    match (legal_anchor(a), legal_anchor(b)) {
        (Some(x), Some(y)) => x == y,
        _ => false,
    }
}

fn legal_anchor(s: &str) -> Option<String> {
    let t = s.trim_start();
    if t.is_empty() { return None; }
    if let Some(rest) = t.strip_prefix('제') {
        let mut digits = String::new();
        let mut saw_jo = false;
        for ch in rest.chars() {
            if ch.is_ascii_digit() { digits.push(ch); }
            else if ch == '조' { saw_jo = true; break; }
            else if ch.is_whitespace() { continue; }
            else { break; }
        }
        if saw_jo && !digits.is_empty() { return Some(format!("article:{}", digits)); }
    }
    let mut chars = t.chars();
    let first = chars.next()?;
    if first == '(' {
        let mut inner = String::new();
        for ch in chars {
            if ch == ')' { break; }
            if ch.is_ascii_digit() { inner.push(ch); } else { return None; }
        }
        if !inner.is_empty() { return Some(format!("paren:{}", inner)); }
    }
    if first.is_ascii_digit() {
        let mut num = first.to_string();
        for ch in chars {
            if ch.is_ascii_digit() { num.push(ch); }
            else if ch == '.' || ch == ')' || ch.is_whitespace() { return Some(format!("num:{}", num)); }
            else { break; }
        }
    }
    if ('①'..='⑳').contains(&first) { return Some(format!("circled:{}", first)); }
    if ('가'..='하').contains(&first) {
        if let Some(ch) = chars.next() {
            if ch == '.' || ch == ')' || ch.is_whitespace() || ch == '\t' { return Some(format!("korean:{}", first)); }
        }
    }
    None
}

fn similarity(a: &str, b: &str) -> f32 {
    if a.is_empty() && b.is_empty() {
        return 1.0;
    }
    let at = crate::token::tokenize_words(a);
    let bt = crate::token::tokenize_words(b);
    let d = TextDiff::from_slices(&at, &bt);
    d.ratio()
}

fn push_or_extend(out: &mut Vec<Change>, kind: ChangeKind, token: &str) {
    if let Some(last) = out.last_mut() {
        if last.kind == kind {
            last.text.push_str(token);
            return;
        }
    }
    out.push(Change { kind, text: token.to_string() });
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
        let old = vec!["hello world".to_string()];
        let new = vec!["hello rust".to_string()];
        let ops = diff_paragraphs(&old, &new);
        assert!(matches!(ops[0], ParaOp::Modified { .. }));
    }
}
