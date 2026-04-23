//! Block model: a document is a sequence of paragraphs and tables.
//! Tables preserve row/cell structure so the diff engine can match rows
//! and cells properly when old/new shapes differ slightly.

use serde::Serialize;

#[derive(Debug, Clone, Serialize)]
#[serde(tag = "kind", rename_all = "lowercase")]
pub enum Block {
    Paragraph(RichParagraph),
    Table { rows: Vec<TableRow> },
}

/// Paragraph with format information preserved from the source DOCX so the
/// writer can emit unchanged paragraphs byte-for-byte (keeping fonts, style
/// references, list bullets, etc.).
#[derive(Debug, Clone, Serialize, Default)]
pub struct RichParagraph {
    /// Concatenated text used by the diff engine.
    pub text: String,
    /// `<w:pPr>...</w:pPr>` raw fragment, or empty if none.
    pub ppr_xml: String,
    /// Original runs with their `<w:rPr>`, in document order.
    pub runs: Vec<RichRun>,
}

#[derive(Debug, Clone, Serialize, Default)]
pub struct RichRun {
    pub rpr_xml: String,
    pub text: String,
}

#[derive(Debug, Clone, Serialize, Default)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
}

#[derive(Debug, Clone, Serialize, Default)]
pub struct TableCell {
    /// Paragraphs inside the cell, each with its pPr and runs preserved so
    /// unchanged cells can be emitted verbatim.
    pub paragraphs: Vec<RichParagraph>,
}

impl TableCell {
    /// Concatenated plain text, used by the diff engine for comparison.
    pub fn text(&self) -> String {
        self.paragraphs
            .iter()
            .map(|p| p.text.as_str())
            .collect::<Vec<_>>()
            .join("\n")
    }
    pub fn from_text(text: impl Into<String>) -> Self {
        Self {
            paragraphs: vec![RichParagraph {
                text: text.into(),
                ppr_xml: String::new(),
                runs: Vec::new(),
            }],
        }
    }
}

impl From<String> for TableCell {
    fn from(s: String) -> Self { TableCell::from_text(s) }
}
impl From<&str> for TableCell {
    fn from(s: &str) -> Self { TableCell::from_text(s) }
}

impl Block {
    /// Flatten to a canonical text signature used for block-level matching.
    /// Tables are represented as a TSV-like string so LCS can pair similar tables.
    pub fn signature(&self) -> String {
        match self {
            Block::Paragraph(p) => p.text.clone(),
            Block::Table { rows } => {
                let mut s = String::from("\x02TABLE\x02");
                for r in rows {
                    for (i, c) in r.cells.iter().enumerate() {
                        if i > 0 { s.push('\t'); }
                        s.push_str(&c.text());
                    }
                    s.push('\n');
                }
                s
            }
        }
    }
}
