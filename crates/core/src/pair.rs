//! Filename-based pair detection.
//!
//! Takes a set of file paths and groups them into (old, new) pairs by
//! recognizing common version markers in filenames:
//!   - State tags: old/new, before/after, 원본/수정, 올드/뉴, 초안/최종, 구/신, …
//!   - Versions: v1, v2, 버전1, (v1.0), …
//!   - Dates: 20260101, 2026-01-01, …
//!   - Trailing numbers: _1, _2, -1, -2
//!
//! Unpaired files are returned separately so the UI can show them for review.

use once_cell::sync::Lazy;
use regex::Regex;
use serde::Serialize;
use std::collections::HashMap;
use std::path::PathBuf;

#[derive(Debug, Clone, Serialize, PartialEq, Eq)]
pub enum StateTag {
    Old,
    New,
}

#[derive(Debug, Clone, Serialize)]
#[serde(tag = "kind", rename_all = "lowercase")]
pub enum Marker {
    State { tag: StateTag, token: String },
    Version { value: String, numeric: u64 },
    Date { value: String },
    Number { value: u64 },
}

#[derive(Debug, Clone, Serialize)]
pub struct PairMatch {
    pub old: PathBuf,
    pub new: PathBuf,
    pub base: String,
    pub reason: String,
}

#[derive(Debug, Default, Serialize)]
pub struct PairDetection {
    pub pairs: Vec<PairMatch>,
    pub unpaired: Vec<PathBuf>,
}

// Tokens that mean "old / original / draft / before".
// NOTE on ordering: regex alternation is left-first, so multi-syllable tokens
// (e.g. "수정전") must appear BEFORE their shorter prefixes ("수정") so the
// pattern matches the longer one. Plain "수정" is reserved for NEW_TOKENS.
const OLD_TOKENS: &[&str] = &[
    // English
    "old", "older", "oldest", "original", "orig", "source", "src",
    "before", "pre", "prev", "previous", "prior",
    "draft", "base", "baseline", "initial",
    "v0", "rev0", "r0",
    // Korean — longer tokens first
    "수정전", "수정_전", "변경전", "변경_전", "개정전", "개정_전",
    "편집전", "편집_전", "작성전", "작성_전", "검토전", "검토_전",
    "이전본", "이전판", "이전버전", "구버전", "구본",
    "원본", "원문", "원안", "원고",
    "초안", "초본", "초고", "초판", "초기본", "초기",
    "작성본", "기존", "기존본",
    "올드", "구", "이전",
];
// Tokens that mean "new / revised / final / after".
const NEW_TOKENS: &[&str] = &[
    // English
    "new", "newer", "newest", "revised", "revise", "revision", "rev",
    "after", "post", "next",
    "final", "fin", "latest", "current", "updated", "update", "edited", "edit",
    // Korean — longer tokens first
    "수정후", "수정_후", "변경후", "변경_후", "개정후", "개정_후",
    "편집후", "편집_후", "작성후", "작성_후", "검토후", "검토_후",
    "이후본", "이후판", "이후버전", "신버전", "신본",
    "수정본", "변경본", "개정본", "편집본", "검토본", "확정본", "최종본", "완성본", "최신본",
    "수정", "변경", "개정", "편집", "검토", "확정", "최종", "완성", "최신",
    "뉴", "신", "이후",
];

static STATE_SUFFIX_RE: Lazy<Regex> = Lazy::new(|| {
    let alt = |xs: &[&str]| xs.join("|");
    // Suffix form: `<base><sep><token>` — or just `<token>` with empty base.
    let pattern = format!(
        r"(?ix) ^ (.*?) [\s_\-().\[\]]* ( {old} | {new} ) [\s_\-().\[\]]* $",
        old = alt(OLD_TOKENS),
        new = alt(NEW_TOKENS),
    );
    Regex::new(&pattern).unwrap()
});

static STATE_PREFIX_RE: Lazy<Regex> = Lazy::new(|| {
    let alt = |xs: &[&str]| xs.join("|");
    // Prefix form: `<token><sep><base>`. Separator required so we don't
    // mis-read e.g. "original" as "orig" + "inal".
    let pattern = format!(
        r"(?ix) ^ [\s_\-().\[\]]* ( {old} | {new} ) [\s_\-().\[\]]+ (.+?) [\s_\-().\[\]]* $",
        old = alt(OLD_TOKENS),
        new = alt(NEW_TOKENS),
    );
    Regex::new(&pattern).unwrap()
});

static VERSION_RE: Lazy<Regex> = Lazy::new(|| {
    Regex::new(
        r"(?ix) ^ (.*?) [\s_\-().\[\]]* (?: v | ver | version | 버전 | 판 ) [\s_\-.]* (\d+(?:\.\d+)*) [\s_\-().\[\]]* $"
    ).unwrap()
});

static DATE_RE: Lazy<Regex> = Lazy::new(|| {
    Regex::new(
        r"(?x) ^ (.*?) [\s_\-().\[\]]+ (\d{4}[-_]?\d{2}[-_]?\d{2}) [\s_\-().\[\]]* $"
    ).unwrap()
});

static TRAILING_NUM_RE: Lazy<Regex> = Lazy::new(|| {
    Regex::new(r"(?x) ^ (.*?) [\s_\-().\[\]]+ (\d{1,4}) [\s_\-().\[\]]* $").unwrap()
});

fn classify_state(token: &str) -> StateTag {
    if OLD_TOKENS.iter().any(|t| t.eq_ignore_ascii_case(token)) {
        StateTag::Old
    } else {
        StateTag::New
    }
}

fn strip_marker(stem: &str) -> Option<(String, Marker)> {
    // Order matters: State > Version > Date > Number.
    // Try suffix form first — it's the most common and least ambiguous.
    if let Some(c) = STATE_SUFFIX_RE.captures(stem) {
        let base = c.get(1)?.as_str().trim().to_string();
        let token = c.get(2)?.as_str().to_string();
        let tag = classify_state(&token);
        // Allow empty base — standalone "old.docx"/"new.docx" still pair.
        return Some((base, Marker::State { tag, token }));
    }
    // Then prefix form: "수정전_문서.docx", "original_contract.docx".
    if let Some(c) = STATE_PREFIX_RE.captures(stem) {
        let token = c.get(1)?.as_str().to_string();
        let base = c.get(2)?.as_str().trim().to_string();
        let tag = classify_state(&token);
        return Some((base, Marker::State { tag, token }));
    }
    if let Some(c) = VERSION_RE.captures(stem) {
        let base = c.get(1)?.as_str().trim().to_string();
        let v = c.get(2)?.as_str().to_string();
        let numeric = v.split('.').next().and_then(|s| s.parse().ok()).unwrap_or(0);
        if !base.is_empty() {
            return Some((base, Marker::Version { value: v, numeric }));
        }
    }
    if let Some(c) = DATE_RE.captures(stem) {
        let base = c.get(1)?.as_str().trim().to_string();
        let v = c.get(2)?.as_str().to_string();
        if !base.is_empty() {
            return Some((base, Marker::Date { value: v }));
        }
    }
    if let Some(c) = TRAILING_NUM_RE.captures(stem) {
        let base = c.get(1)?.as_str().trim().to_string();
        let v: u64 = c.get(2)?.as_str().parse().ok()?;
        if !base.is_empty() {
            return Some((base, Marker::Number { value: v }));
        }
    }
    None
}

pub fn detect_pairs(files: &[PathBuf]) -> PairDetection {
    let mut groups: HashMap<String, Vec<(PathBuf, Marker)>> = HashMap::new();
    let mut unpaired = Vec::new();

    for p in files {
        let stem = match p.file_stem().and_then(|s| s.to_str()) {
            Some(s) => s,
            None => {
                unpaired.push(p.clone());
                continue;
            }
        };
        match strip_marker(stem) {
            Some((base, marker)) => {
                let key = normalize_base(&base);
                groups.entry(key).or_default().push((p.clone(), marker));
            }
            None => unpaired.push(p.clone()),
        }
    }

    let mut pairs = Vec::new();
    for (base, items) in groups {
        make_pairs(&base, items, &mut pairs, &mut unpaired);
    }

    pairs.sort_by(|a, b| a.base.cmp(&b.base));
    PairDetection { pairs, unpaired }
}

fn normalize_base(s: &str) -> String {
    s.trim_matches(|c: char| c.is_whitespace() || "-_().[]".contains(c))
        .to_lowercase()
}

fn make_pairs(
    base: &str,
    mut items: Vec<(PathBuf, Marker)>,
    pairs: &mut Vec<PairMatch>,
    unpaired: &mut Vec<PathBuf>,
) {
    if items.len() < 2 {
        unpaired.extend(items.into_iter().map(|(p, _)| p));
        return;
    }

    // State-tag group: take one Old and one New.
    if items.iter().all(|(_, m)| matches!(m, Marker::State { .. })) {
        let (olds, news): (Vec<_>, Vec<_>) = items
            .into_iter()
            .partition(|(_, m)| matches!(m, Marker::State { tag: StateTag::Old, .. }));
        let mut olds = olds;
        let mut news = news;
        while let (Some(o), Some(n)) = (olds.pop(), news.pop()) {
            pairs.push(PairMatch {
                old: o.0,
                new: n.0,
                base: base.to_string(),
                reason: "state suffix (old/new)".into(),
            });
        }
        unpaired.extend(olds.into_iter().map(|(p, _)| p));
        unpaired.extend(news.into_iter().map(|(p, _)| p));
        return;
    }

    // Sortable markers (version / date / number): sort ascending, pair consecutive.
    items.sort_by(|a, b| marker_sort_key(&a.1).cmp(&marker_sort_key(&b.1)));
    let reason = match items.first().map(|x| &x.1) {
        Some(Marker::Version { .. }) => "version suffix",
        Some(Marker::Date { .. }) => "date suffix",
        Some(Marker::Number { .. }) => "numeric suffix",
        _ => "suffix",
    };
    // Pair consecutive: (0,1), (2,3), ...  Leftover goes unpaired.
    let mut iter = items.into_iter();
    loop {
        match (iter.next(), iter.next()) {
            (Some(a), Some(b)) => pairs.push(PairMatch {
                old: a.0,
                new: b.0,
                base: base.to_string(),
                reason: reason.into(),
            }),
            (Some(a), None) => unpaired.push(a.0),
            _ => break,
        }
    }
}

fn marker_sort_key(m: &Marker) -> String {
    match m {
        Marker::Version { numeric, value } => format!("v{:010}-{}", numeric, value),
        Marker::Date { value } => format!("d{}", value.replace(['-', '_'], "")),
        Marker::Number { value } => format!("n{:010}", value),
        Marker::State { tag, .. } => match tag {
            StateTag::Old => "s0".into(),
            StateTag::New => "s1".into(),
        },
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn p(s: &str) -> PathBuf { PathBuf::from(s) }

    #[test]
    fn state_english() {
        let got = detect_pairs(&[p("contract_old.docx"), p("contract_new.docx")]);
        assert_eq!(got.pairs.len(), 1);
        assert_eq!(got.pairs[0].old.file_name().unwrap(), "contract_old.docx");
        assert_eq!(got.pairs[0].new.file_name().unwrap(), "contract_new.docx");
    }

    #[test]
    fn state_korean() {
        let got = detect_pairs(&[p("계약서_원본.docx"), p("계약서_수정.docx")]);
        assert_eq!(got.pairs.len(), 1);
    }

    #[test]
    fn state_olympic_style() {
        let got = detect_pairs(&[p("계약서(올드).docx"), p("계약서(뉴).docx")]);
        assert_eq!(got.pairs.len(), 1);
    }

    #[test]
    fn version_pair() {
        let got = detect_pairs(&[p("report_v3.docx"), p("report_v1.docx"), p("report_v2.docx")]);
        // three v's -> first two consecutive paired, one leftover
        assert_eq!(got.pairs.len(), 1);
        assert_eq!(got.unpaired.len(), 1);
    }

    #[test]
    fn date_pair() {
        let got = detect_pairs(&[p("memo_20260101.docx"), p("memo_20260315.docx")]);
        assert_eq!(got.pairs.len(), 1);
        assert!(got.pairs[0].old.file_name().unwrap().to_str().unwrap().contains("20260101"));
    }

    #[test]
    fn unpaired_orphan() {
        let got = detect_pairs(&[p("lonely.docx")]);
        assert_eq!(got.pairs.len(), 0);
        assert_eq!(got.unpaired.len(), 1);
    }

    #[test]
    fn bare_old_new() {
        let got = detect_pairs(&[p("old.docx"), p("new.docx")]);
        assert_eq!(got.pairs.len(), 1);
        assert_eq!(got.unpaired.len(), 0);
    }

    #[test]
    fn prefix_english() {
        let got = detect_pairs(&[p("old_contract.docx"), p("new_contract.docx")]);
        assert_eq!(got.pairs.len(), 1);
    }

    #[test]
    fn korean_before_after() {
        let got = detect_pairs(&[p("수정전_보고서.docx"), p("수정후_보고서.docx")]);
        assert_eq!(got.pairs.len(), 1);
    }

    #[test]
    fn korean_revised_final() {
        let got = detect_pairs(&[p("보고서_초안.docx"), p("보고서_최종.docx")]);
        assert_eq!(got.pairs.len(), 1);
    }

    #[test]
    fn english_draft_final() {
        let got = detect_pairs(&[p("spec_draft.docx"), p("spec_final.docx")]);
        assert_eq!(got.pairs.len(), 1);
    }

    #[test]
    fn mixed_bases() {
        let got = detect_pairs(&[
            p("a_old.docx"), p("a_new.docx"),
            p("b_old.docx"), p("b_new.docx"),
        ]);
        assert_eq!(got.pairs.len(), 2);
    }
}
