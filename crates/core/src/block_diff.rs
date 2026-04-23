//! Block-level diff: matches paragraphs and tables, then within each matched
//! table matches rows and cells. Rows that have no counterpart are marked
//! Insert/Delete; matched rows get cell-level word diffs.

use serde::Serialize;
use similar::{DiffOp, TextDiff};

use crate::block::{Block, RichParagraph, TableRow};
use crate::diff::{diff_paragraphs, diff_words, Change, ParaOp};

#[derive(Debug, Clone, Serialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum BlockOp {
    /// Paragraph block with either equal/insert/delete/modified/moved semantics.
    /// `source` carries the original pPr + runs so the writer can emit
    /// unchanged paragraphs with their formatting intact.
    Para {
        op: ParaOp,
        #[serde(skip)]
        source: RichParagraph,
    },
    /// Entire table inserted (present only in new).
    TableInsert { rows: Vec<TableRow> },
    /// Entire table deleted (present only in old).
    TableDelete { rows: Vec<TableRow> },
    /// Matched tables whose content may differ — row/cell-level diff inside.
    TableDiff { diff: TableDiff },
}

#[derive(Debug, Clone, Serialize)]
pub struct TableDiff {
    pub rows: Vec<RowOp>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(tag = "kind", rename_all = "lowercase")]
pub enum RowOp {
    Equal { row: TableRow },
    Insert { row: TableRow },
    Delete { row: TableRow },
    /// Same-position row with at least one cell changed. Each cell is either
    /// unchanged or a word-level change list.
    Modified { cells: Vec<CellOp> },
}

#[derive(Debug, Clone, Serialize)]
#[serde(tag = "kind", rename_all = "lowercase")]
pub enum CellOp {
    Equal { text: String },
    Changed { changes: Vec<Change> },
}

impl CellOp {
    pub fn text_equal(&self) -> Option<&str> {
        match self { CellOp::Equal { text } => Some(text), _ => None }
    }
}

const BLOCK_TABLE_SIMILARITY: f32 = 0.4;
const ROW_MATCH_THRESHOLD: f32 = 0.5;

/// Diff two block streams. Paragraphs use the existing paragraph-level engine.
/// Tables are matched by content signature; matched tables get structural diff.
pub fn diff_blocks(old: &[Block], new: &[Block]) -> Vec<BlockOp> {
    let old_sigs: Vec<String> = old.iter().map(Block::signature).collect();
    let new_sigs: Vec<String> = new.iter().map(Block::signature).collect();
    let old_refs: Vec<&str> = old_sigs.iter().map(String::as_str).collect();
    let new_refs: Vec<&str> = new_sigs.iter().map(String::as_str).collect();
    let diff = TextDiff::from_slices(&old_refs, &new_refs);

    let mut out = Vec::new();
    for op in diff.ops() {
        match *op {
            DiffOp::Equal { old_index, len, .. } => {
                for i in 0..len {
                    out.push(emit_equal(&old[old_index + i]));
                }
            }
            DiffOp::Insert { new_index, new_len, .. } => {
                for i in 0..new_len {
                    out.push(emit_insert(&new[new_index + i]));
                }
            }
            DiffOp::Delete { old_index, old_len, .. } => {
                for i in 0..old_len {
                    out.push(emit_delete(&old[old_index + i]));
                }
            }
            DiffOp::Replace { old_index, old_len, new_index, new_len } => {
                emit_replace(
                    &old[old_index..old_index + old_len],
                    &new[new_index..new_index + new_len],
                    &mut out,
                );
            }
        }
    }
    // Post-pass: reclassify high-similarity Para{Insert}+Para{Delete} pairs
    // that came from separate top-level LCS ops (not a Replace block) as
    // MovedFrom/MovedTo. Without this, a paragraph that moved across matched
    // boundaries would show as a delete-and-insert instead of a green move.
    detect_cross_block_moves(out)
}

const CROSS_MOVE_THRESHOLD: f32 = 0.85;

fn detect_cross_block_moves(mut ops: Vec<BlockOp>) -> Vec<BlockOp> {
    // Collect indices of Para Inserts and Deletes.
    let mut ins_idx = Vec::new();
    let mut del_idx = Vec::new();
    for (i, b) in ops.iter().enumerate() {
        if let BlockOp::Para { op, .. } = b {
            match op {
                ParaOp::Insert { .. } => ins_idx.push(i),
                ParaOp::Delete { .. } => del_idx.push(i),
                _ => {}
            }
        }
    }
    if ins_idx.is_empty() || del_idx.is_empty() {
        return ops;
    }

    let mut used_ins = vec![false; ins_idx.len()];
    let mut next_move_id: u32 = 1;
    for &di in &del_idx {
        let del_text = match &ops[di] {
            BlockOp::Para { op: ParaOp::Delete { text }, .. } => text.clone(),
            _ => continue,
        };
        let mut best: Option<(usize, f32)> = None;
        for (k, &ii) in ins_idx.iter().enumerate() {
            if used_ins[k] { continue; }
            let ins_text = match &ops[ii] {
                BlockOp::Para { op: ParaOp::Insert { text }, .. } => text,
                _ => continue,
            };
            let sim = block_similarity(&del_text, ins_text);
            if sim >= CROSS_MOVE_THRESHOLD && best.map_or(true, |(_, bs)| sim > bs) {
                best = Some((k, sim));
            }
        }
        if let Some((k, _)) = best {
            used_ins[k] = true;
            let ii = ins_idx[k];
            let move_id = next_move_id;
            next_move_id += 1;
            if let BlockOp::Para { op: ParaOp::Delete { text: dt }, source: ds } = ops[di].clone() {
                ops[di] = BlockOp::Para {
                    op: ParaOp::MovedFrom { text: dt, move_id },
                    source: ds,
                };
            }
            if let BlockOp::Para { op: ParaOp::Insert { text: it }, source: is } = ops[ii].clone() {
                ops[ii] = BlockOp::Para {
                    op: ParaOp::MovedTo { text: it, move_id },
                    source: is,
                };
            }
        }
    }
    ops
}

fn block_similarity(a: &str, b: &str) -> f32 {
    if a.is_empty() && b.is_empty() { return 1.0; }
    let at = crate::token::tokenize_words(a);
    let bt = crate::token::tokenize_words(b);
    TextDiff::from_slices(&at, &bt).ratio()
}

fn emit_equal(b: &Block) -> BlockOp {
    match b {
        Block::Paragraph(p) => BlockOp::Para {
            op: ParaOp::Equal { text: p.text.clone() },
            source: p.clone(),
        },
        Block::Table { rows } => BlockOp::TableDiff {
            diff: TableDiff {
                rows: rows.iter().map(|r| RowOp::Equal { row: r.clone() }).collect(),
            },
        },
    }
}

fn emit_insert(b: &Block) -> BlockOp {
    match b {
        Block::Paragraph(p) => BlockOp::Para {
            op: ParaOp::Insert { text: p.text.clone() },
            source: p.clone(),
        },
        Block::Table { rows } => BlockOp::TableInsert { rows: rows.clone() },
    }
}

fn emit_delete(b: &Block) -> BlockOp {
    match b {
        Block::Paragraph(p) => BlockOp::Para {
            op: ParaOp::Delete { text: p.text.clone() },
            source: p.clone(),
        },
        Block::Table { rows } => BlockOp::TableDelete { rows: rows.clone() },
    }
}

fn emit_replace(olds: &[Block], news: &[Block], out: &mut Vec<BlockOp>) {
    // Split by kind first — paragraphs pair with paragraphs, tables with tables.
    let (old_rich, old_tables) = partition_by_kind(olds);
    let (new_rich, new_tables) = partition_by_kind(news);
    let old_paras: Vec<String> = old_rich.iter().map(|p| p.text.clone()).collect();
    let new_paras: Vec<String> = new_rich.iter().map(|p| p.text.clone()).collect();

    // Paragraphs: reuse the paragraph-level engine which already does similarity
    // threshold pairing and move detection. We pair the resulting ParaOps back
    // to their source RichParagraph by matching text — imperfect when the same
    // text appears twice, good enough in practice for legal docs.
    if !old_paras.is_empty() || !new_paras.is_empty() {
        let ops = diff_paragraphs(&old_paras, &new_paras);
        for op in ops {
            let source = match &op {
                ParaOp::Equal { text } | ParaOp::Delete { text } | ParaOp::MovedFrom { text, .. } => {
                    find_rich(&old_rich, text)
                }
                ParaOp::Insert { text } | ParaOp::MovedTo { text, .. } => {
                    find_rich(&new_rich, text)
                }
                ParaOp::Modified { .. } => RichParagraph::default(),
            };
            out.push(BlockOp::Para { op, source });
        }
    }

    // Tables: pair by signature similarity; matched pairs get row/cell diff,
    // unmatched become full Insert/Delete.
    let mut used_new = vec![false; new_tables.len()];
    for (oi, (old_rows, old_sig)) in old_tables.iter().enumerate() {
        let mut best: Option<(usize, f32)> = None;
        for (nj, (_, new_sig)) in new_tables.iter().enumerate() {
            if used_new[nj] { continue; }
            let s = string_similarity(old_sig, new_sig);
            if s >= BLOCK_TABLE_SIMILARITY && best.map_or(true, |(_, bs)| s > bs) {
                best = Some((nj, s));
            }
        }
        match best {
            Some((nj, _)) => {
                used_new[nj] = true;
                let diff = diff_table_rows(old_rows, &new_tables[nj].0);
                out.push(BlockOp::TableDiff { diff });
            }
            None => {
                out.push(BlockOp::TableDelete { rows: old_rows.clone() });
            }
        }
        let _ = oi;
    }
    for (nj, (new_rows, _)) in new_tables.iter().enumerate() {
        if !used_new[nj] {
            out.push(BlockOp::TableInsert { rows: new_rows.clone() });
        }
    }
}

fn partition_by_kind(blocks: &[Block]) -> (Vec<RichParagraph>, Vec<(Vec<TableRow>, String)>) {
    let mut paras = Vec::new();
    let mut tables = Vec::new();
    for b in blocks {
        match b {
            Block::Paragraph(p) => paras.push(p.clone()),
            Block::Table { rows } => {
                let sig = b.signature();
                tables.push((rows.clone(), sig));
            }
        }
    }
    (paras, tables)
}

fn find_rich(list: &[RichParagraph], text: &str) -> RichParagraph {
    list.iter().find(|p| p.text == text).cloned().unwrap_or_default()
}

fn string_similarity(a: &str, b: &str) -> f32 {
    if a.is_empty() && b.is_empty() { return 1.0; }
    let at = crate::token::tokenize_words(a);
    let bt = crate::token::tokenize_words(b);
    TextDiff::from_slices(&at, &bt).ratio()
}

/// Structural diff of two table row vectors.
fn diff_table_rows(old: &[TableRow], new: &[TableRow]) -> TableDiff {
    let old_sigs: Vec<String> = old.iter().map(row_signature).collect();
    let new_sigs: Vec<String> = new.iter().map(row_signature).collect();
    let old_refs: Vec<&str> = old_sigs.iter().map(String::as_str).collect();
    let new_refs: Vec<&str> = new_sigs.iter().map(String::as_str).collect();
    let diff = TextDiff::from_slices(&old_refs, &new_refs);

    let mut rows = Vec::new();
    for op in diff.ops() {
        match *op {
            DiffOp::Equal { old_index, len, .. } => {
                for i in 0..len {
                    rows.push(RowOp::Equal { row: old[old_index + i].clone() });
                }
            }
            DiffOp::Insert { new_index, new_len, .. } => {
                for i in 0..new_len {
                    rows.push(RowOp::Insert { row: new[new_index + i].clone() });
                }
            }
            DiffOp::Delete { old_index, old_len, .. } => {
                for i in 0..old_len {
                    rows.push(RowOp::Delete { row: old[old_index + i].clone() });
                }
            }
            DiffOp::Replace { old_index, old_len, new_index, new_len } => {
                let paired = old_len.min(new_len);
                for i in 0..paired {
                    let o = &old[old_index + i];
                    let n = &new[new_index + i];
                    let sim = string_similarity(&row_signature(o), &row_signature(n));
                    if sim >= ROW_MATCH_THRESHOLD {
                        rows.push(RowOp::Modified { cells: diff_cells(o, n) });
                    } else {
                        rows.push(RowOp::Delete { row: o.clone() });
                        rows.push(RowOp::Insert { row: n.clone() });
                    }
                }
                for i in paired..old_len {
                    rows.push(RowOp::Delete { row: old[old_index + i].clone() });
                }
                for i in paired..new_len {
                    rows.push(RowOp::Insert { row: new[new_index + i].clone() });
                }
            }
        }
    }
    TableDiff { rows }
}

fn row_signature(r: &TableRow) -> String {
    r.cells.iter().map(|c| c.text()).collect::<Vec<_>>().join("\t")
}

fn diff_cells(old: &TableRow, new: &TableRow) -> Vec<CellOp> {
    let n = old.cells.len().max(new.cells.len());
    let mut out = Vec::with_capacity(n);
    for i in 0..n {
        let empty = String::new();
        let o_text: String = old.cells.get(i).map(|c| c.text()).unwrap_or_else(|| empty.clone());
        let m_text: String = new.cells.get(i).map(|c| c.text()).unwrap_or_else(|| empty.clone());
        if o_text == m_text {
            out.push(CellOp::Equal { text: o_text });
        } else {
            out.push(CellOp::Changed { changes: diff_words(&o_text, &m_text) });
        }
    }
    out
}
