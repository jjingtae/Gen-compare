//! Diff statistics for JSON reporting.

use crate::block_diff::{BlockOp, CellOp, RowOp};
use crate::diff::{ChangeKind, ParaOp};
use serde::Serialize;

#[derive(Debug, Default, Clone, Serialize)]
pub struct Stats {
    pub paragraphs_equal: usize,
    pub paragraphs_inserted: usize,
    pub paragraphs_deleted: usize,
    pub paragraphs_modified: usize,
    pub paragraphs_moved: usize,
    pub words_inserted: usize,
    pub words_deleted: usize,
    pub words_moved: usize,
    pub tables_inserted: usize,
    pub tables_deleted: usize,
    pub tables_modified: usize,
    pub rows_inserted: usize,
    pub rows_deleted: usize,
    pub rows_modified: usize,
}

pub fn stats_of(ops: &[ParaOp]) -> Stats {
    let mut s = Stats::default();
    for op in ops {
        match op {
            ParaOp::Equal { .. } => s.paragraphs_equal += 1,
            ParaOp::Insert { text } => {
                s.paragraphs_inserted += 1;
                s.words_inserted += count_words(text);
            }
            ParaOp::Delete { text } => {
                s.paragraphs_deleted += 1;
                s.words_deleted += count_words(text);
            }
            ParaOp::Modified { changes } => {
                s.paragraphs_modified += 1;
                for c in changes {
                    match c.kind {
                        ChangeKind::Insert => s.words_inserted += count_words(&c.text),
                        ChangeKind::Delete => s.words_deleted += count_words(&c.text),
                        ChangeKind::Equal => {}
                    }
                }
            }
            // Count the moved paragraph once (on the MovedTo side) so it
            // doesn't appear twice in the stats.
            ParaOp::MovedTo { text, .. } => {
                s.paragraphs_moved += 1;
                s.words_moved += count_words(text);
            }
            ParaOp::MovedFrom { .. } => {}
        }
    }
    s
}

fn count_words(text: &str) -> usize {
    text.split_whitespace().count()
}

/// Stats computed over a block stream. Calls stats_of internally for ParaOps
/// and adds table / row counts.
pub fn stats_of_blocks(ops: &[BlockOp]) -> Stats {
    let mut s = Stats::default();
    for b in ops {
        match b {
            BlockOp::Para { op, .. } => merge_para_op(&mut s, op),
            BlockOp::TableInsert { rows } => {
                s.tables_inserted += 1;
                s.rows_inserted += rows.len();
                for r in rows {
                    for cell in &r.cells {
                        s.words_inserted += count_words(&cell.text());
                    }
                }
            }
            BlockOp::TableDelete { rows } => {
                s.tables_deleted += 1;
                s.rows_deleted += rows.len();
                for r in rows {
                    for cell in &r.cells {
                        s.words_deleted += count_words(&cell.text());
                    }
                }
            }
            BlockOp::TableDiff { diff } => {
                let mut changed = false;
                for row in &diff.rows {
                    match row {
                        RowOp::Equal { .. } => {}
                        RowOp::Insert { row } => {
                            changed = true;
                            s.rows_inserted += 1;
                            for c in &row.cells { s.words_inserted += count_words(&c.text()); }
                        }
                        RowOp::Delete { row } => {
                            changed = true;
                            s.rows_deleted += 1;
                            for c in &row.cells { s.words_deleted += count_words(&c.text()); }
                        }
                        RowOp::Modified { cells } => {
                            changed = true;
                            s.rows_modified += 1;
                            for c in cells {
                                if let CellOp::Changed { changes } = c {
                                    for ch in changes {
                                        match ch.kind {
                                            ChangeKind::Insert => s.words_inserted += count_words(&ch.text),
                                            ChangeKind::Delete => s.words_deleted += count_words(&ch.text),
                                            ChangeKind::Equal => {}
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if changed {
                    s.tables_modified += 1;
                }
            }
        }
    }
    s
}

fn merge_para_op(s: &mut Stats, op: &ParaOp) {
    match op {
        ParaOp::Equal { .. } => s.paragraphs_equal += 1,
        ParaOp::Insert { text } => {
            s.paragraphs_inserted += 1;
            s.words_inserted += count_words(text);
        }
        ParaOp::Delete { text } => {
            s.paragraphs_deleted += 1;
            s.words_deleted += count_words(text);
        }
        ParaOp::Modified { changes } => {
            s.paragraphs_modified += 1;
            for c in changes {
                match c.kind {
                    ChangeKind::Insert => s.words_inserted += count_words(&c.text),
                    ChangeKind::Delete => s.words_deleted += count_words(&c.text),
                    ChangeKind::Equal => {}
                }
            }
        }
        ParaOp::MovedTo { text, .. } => {
            s.paragraphs_moved += 1;
            s.words_moved += count_words(text);
        }
        ParaOp::MovedFrom { .. } => {}
    }
}
