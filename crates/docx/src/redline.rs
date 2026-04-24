//! Redline DOCX writer. Emits two styles of output:
//!
//! * **Color**: direct run formatting (color + underline/strike). Always
//!   renders the intended red/blue/green because Word/LibreOffice don't
//!   override regular run properties.
//! * **TrackChange**: real `<w:ins>/<w:del>/<w:moveFrom>/<w:moveTo>`
//!   revision marks. Accept/reject works, but per-author display color can
//!   override our baked colors.
//!
//! Tables are emitted as proper `<w:tbl>` structures with row-level and
//! cell-level change marking so structural edits (row insert/delete) are
//! visible to the reader.

use anyhow::Result;
use compare_core::{
    stats_of_blocks, BlockOp, CellOp, ChangeKind, ParaOp, RowOp, Stats, TableRow,
};
use std::fs::File;
use std::io::Write;
use std::path::Path;
use time::format_description::well_known::Rfc3339;
use time::OffsetDateTime;
use zip::write::SimpleFileOptions;

use crate::read::RawParts;

const COLOR_INSERT: &str = "0000FF"; // blue
const COLOR_DELETE: &str = "FF0000"; // red
const COLOR_MOVE: &str = "008000";   // green

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RedlineStyle {
    Color,
    TrackChange,
}

#[derive(Debug, Clone)]
pub struct RedlineOptions {
    pub style: RedlineStyle,
    pub author: String,
    pub date: String,
    /// Paragraph-level header diff (appended in summary).
    pub header_changes: Option<Vec<ParaOp>>,
    /// Paragraph-level footer diff (appended in summary).
    pub footer_changes: Option<Vec<ParaOp>>,
    /// Paragraph-level comments diff (appended in summary).
    pub comment_changes: Option<Vec<ParaOp>>,
    /// Paragraph-level footnotes diff.
    pub footnote_changes: Option<Vec<ParaOp>>,
    /// Paragraph-level endnotes diff.
    pub endnote_changes: Option<Vec<ParaOp>>,
    /// Original document filename (shown in Litera-style summary).
    pub original_name: Option<String>,
    /// Modified document filename (shown in Litera-style summary).
    pub modified_name: Option<String>,
    /// Raw XML parts from the source document. When present, fonts/styles/
    /// theme colors are preserved in the output.
    pub source_parts: Option<RawParts>,
}

impl Default for RedlineOptions {
    fn default() -> Self {
        Self {
            style: RedlineStyle::Color,
            author: default_author(),
            date: default_date(),
            header_changes: None,
            footer_changes: None,
            comment_changes: None,
            footnote_changes: None,
            endnote_changes: None,
            original_name: None,
            modified_name: None,
            source_parts: None,
        }
    }
}

fn default_author() -> String {
    std::env::var("USERNAME")
        .or_else(|_| std::env::var("USER"))
        .unwrap_or_else(|_| "compare".to_string())
}

fn default_date() -> String {
    OffsetDateTime::now_utc()
        .format(&Rfc3339)
        .unwrap_or_else(|_| "2026-01-01T00:00:00Z".into())
}

pub fn write_redline<P: AsRef<Path>>(
    path: P,
    ops: &[BlockOp],
    opts: &RedlineOptions,
) -> Result<()> {
    // Atomic write: build the zip in a temp file next to the target, then
    // rename into place. If anything fails mid-write we won't leave a
    // half-written .docx that Word would report as "corrupted".
    let final_path = path.as_ref();
    let tmp_path = tmp_sibling(final_path);
    // Remove any leftover temp from a previous failed run.
    let _ = std::fs::remove_file(&tmp_path);

    // Wrap the actual writing so we can clean up the temp file on error.
    let result: Result<()> = (|| {
        let file = File::create(&tmp_path)?;
        let mut zip = zip::ZipWriter::new(file);
        let zopts = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);

        let parts = opts.source_parts.clone().unwrap_or_default();

        zip.start_file("[Content_Types].xml", zopts)?;
        zip.write_all(content_types_for(&parts).as_bytes())?;

        zip.start_file("_rels/.rels", zopts)?;
        zip.write_all(ROOT_RELS.as_bytes())?;

        zip.start_file("word/_rels/document.xml.rels", zopts)?;
        zip.write_all(doc_rels_for(&parts).as_bytes())?;

        zip.start_file("word/document.xml", zopts)?;
        zip.write_all(build_document(ops, opts).as_bytes())?;

        // Copy preserved parts verbatim. Missing parts are harmless — style/font
        // references simply fall back to reader defaults.
        copy_part(&mut zip, zopts, "word/styles.xml", &parts.styles_xml)?;
        copy_part(&mut zip, zopts, "word/numbering.xml", &parts.numbering_xml)?;
        copy_part(&mut zip, zopts, "word/theme/theme1.xml", &parts.theme1_xml)?;
        copy_part(&mut zip, zopts, "word/fontTable.xml", &parts.font_table_xml)?;
        copy_part(&mut zip, zopts, "word/settings.xml", &parts.settings_xml)?;
        copy_part(&mut zip, zopts, "word/webSettings.xml", &parts.web_settings_xml)?;
        copy_part(&mut zip, zopts, "word/stylesWithEffects.xml", &parts.style_with_effects_xml)?;
        copy_part(&mut zip, zopts, "word/comments.xml", &parts.comments_xml)?;
        copy_part(&mut zip, zopts, "word/footnotes.xml", &parts.footnotes_xml)?;
        copy_part(&mut zip, zopts, "word/endnotes.xml", &parts.endnotes_xml)?;
        for (name, bytes) in &parts.headers {
            zip.start_file(name, zopts)?;
            zip.write_all(bytes)?;
        }
        for (name, bytes) in &parts.footers {
            zip.start_file(name, zopts)?;
            zip.write_all(bytes)?;
        }
        for (name, bytes) in &parts.header_footer_rels {
            zip.start_file(name, zopts)?;
            zip.write_all(bytes)?;
        }
        // Copy binary resources (images, embedded objects, embedded fonts)
        // referenced by header/footer rels. Without these, Word considers
        // the rels "dangling" and refuses to open the document as corrupted.
        // Images/fonts must use Stored (no deflate) since they're already
        // compressed or binary — deflating them can cause reader quirks.
        let store_opts = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
        for (name, bytes) in &parts.binary_resources {
            zip.start_file(name, store_opts)?;
            zip.write_all(bytes)?;
        }

        let finished = zip.finish()?;
        // Force data to disk before rename.
        finished.sync_all()?;
        Ok(())
    })();

    if let Err(e) = result {
        let _ = std::fs::remove_file(&tmp_path);
        return Err(e);
    }

    // Atomic replace. On Windows, rename fails if target exists, so remove first.
    let _ = std::fs::remove_file(final_path);
    std::fs::rename(&tmp_path, final_path).map_err(|e| {
        let _ = std::fs::remove_file(&tmp_path);
        e
    })?;
    Ok(())
}

fn tmp_sibling(final_path: &Path) -> std::path::PathBuf {
    let parent = final_path.parent().unwrap_or(Path::new("."));
    let nonce = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_nanos())
        .unwrap_or(0);
    let name = final_path
        .file_name()
        .and_then(|s| s.to_str())
        .unwrap_or("out.docx");
    parent.join(format!(".{}.{}.tmp", name, nonce))
}

fn mime_for_ext(ext: &str) -> &'static str {
    match ext {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "tif" | "tiff" => "image/tiff",
        "svg" => "image/svg+xml",
        "wmf" => "image/x-wmf",
        "emf" => "image/x-emf",
        "ico" => "image/x-icon",
        "ttf" => "application/x-font-ttf",
        "otf" => "application/x-font-otf",
        "bin" => "application/vnd.openxmlformats-officedocument.oleObject",
        _ => "application/octet-stream",
    }
}

fn copy_part(
    zip: &mut zip::ZipWriter<File>,
    zopts: SimpleFileOptions,
    name: &str,
    bytes: &Option<Vec<u8>>,
) -> Result<()> {
    if let Some(b) = bytes {
        zip.start_file(name, zopts)?;
        zip.write_all(b)?;
    }
    Ok(())
}

fn content_types_for(parts: &RawParts) -> String {
    let mut s = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>"#,
    );
    // Declare content types for any binary-resource extensions we copy.
    // Word is picky — if a file's extension isn't declared here, it flags
    // the package as corrupted.
    let mut seen_ext: std::collections::HashSet<String> = std::collections::HashSet::new();
    for (name, _) in &parts.binary_resources {
        if let Some(ext) = std::path::Path::new(name).extension().and_then(|e| e.to_str()) {
            let ext_lower = ext.to_ascii_lowercase();
            if seen_ext.insert(ext_lower.clone()) {
                let mime = mime_for_ext(&ext_lower);
                s.push_str(&format!(
                    r#"<Default Extension="{}" ContentType="{}"/>"#,
                    ext_lower, mime
                ));
            }
        }
    }
    s.push_str(r#"<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>"#);
    if parts.styles_xml.is_some() {
        s.push_str(r#"<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>"#);
    }
    if parts.numbering_xml.is_some() {
        s.push_str(r#"<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>"#);
    }
    if parts.theme1_xml.is_some() {
        s.push_str(r#"<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>"#);
    }
    if parts.font_table_xml.is_some() {
        s.push_str(r#"<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>"#);
    }
    if parts.settings_xml.is_some() {
        s.push_str(r#"<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>"#);
    }
    if parts.web_settings_xml.is_some() {
        s.push_str(r#"<Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>"#);
    }
    if parts.style_with_effects_xml.is_some() {
        s.push_str(r#"<Override PartName="/word/stylesWithEffects.xml" ContentType="application/vnd.ms-word.stylesWithEffects+xml"/>"#);
    }
    s.push_str("</Types>");
    s
}

fn doc_rels_for(parts: &RawParts) -> String {
    let mut s = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    let mut next_id: u32 = 1;
    let mut add = |s: &mut String, target: &str, typ: &str| {
        s.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="{}" Target="{}"/>"#,
            next_id, typ, target
        ));
        next_id += 1;
    };
    if parts.styles_xml.is_some() {
        add(&mut s, "styles.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
    }
    if parts.numbering_xml.is_some() {
        add(&mut s, "numbering.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering");
    }
    if parts.theme1_xml.is_some() {
        add(&mut s, "theme/theme1.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme");
    }
    if parts.font_table_xml.is_some() {
        add(&mut s, "fontTable.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable");
    }
    if parts.settings_xml.is_some() {
        add(&mut s, "settings.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings");
    }
    if parts.web_settings_xml.is_some() {
        add(&mut s, "webSettings.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings");
    }
    s.push_str("</Relationships>");
    s
}

struct RevId(u32);
impl RevId {
    fn next(&mut self) -> u32 {
        let v = self.0;
        self.0 += 1;
        v
    }
}

fn build_document(ops: &[BlockOp], opts: &RedlineOptions) -> String {
    let mut body = String::new();
    let mut rev = RevId(1);

    for op in ops {
        emit_block_op(&mut body, op, &mut rev, opts);
    }

    body.push_str(&summary_page(&stats_of_blocks(ops), opts, &mut rev));

    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>{}</w:body>
</w:document>"#,
        body
    )
}

fn emit_block_op(body: &mut String, op: &BlockOp, rev: &mut RevId, opts: &RedlineOptions) {
    match op {
        BlockOp::Para { op, source } => {
            emit_rich_paragraph(body, op, source, rev, opts);
        }
        BlockOp::TableInsert { rows } => emit_table(body, rows, BlockChange::Insert, rev, opts),
        BlockOp::TableDelete { rows } => emit_table(body, rows, BlockChange::Delete, rev, opts),
        BlockOp::TableDiff { diff } => emit_table_diff(body, &diff.rows, rev, opts),
    }
}

/// Emit a paragraph preserving pPr and — for unchanged paragraphs — the
/// original runs with their rPr intact. Changed paragraphs fall back to
/// flat diff runs with coloring.
fn emit_rich_paragraph(
    body: &mut String,
    op: &ParaOp,
    source: &compare_core::RichParagraph,
    rev: &mut RevId,
    opts: &RedlineOptions,
) {
    let ppr = &source.ppr_xml;
    body.push_str("<w:p>");
    if !ppr.is_empty() {
        body.push_str(ppr);
    }
    match op {
        ParaOp::Equal { .. } if !source.runs.is_empty() => {
            // Emit original runs verbatim so fonts, bold, italic, color, etc.
            // are perfectly preserved on unchanged paragraphs.
            for r in &source.runs {
                body.push_str(r#"<w:r>"#);
                if !r.rpr_xml.is_empty() {
                    body.push_str(&r.rpr_xml);
                }
                body.push_str(&format!(
                    r#"<w:t xml:space="preserve">{}</w:t></w:r>"#,
                    xml_escape(&r.text)
                ));
            }
        }
        _ => emit_para_op(body, op, rev, opts),
    }
    body.push_str("</w:p>");
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum BlockChange {
    Insert,
    Delete,
}

fn emit_table(body: &mut String, rows: &[TableRow], change: BlockChange, rev: &mut RevId, opts: &RedlineOptions) {
    let col_count = rows.iter().map(|r| r.cells.len()).max().unwrap_or(0);
    body.push_str(&table_open(col_count));
    for row in rows {
        let row_change = match change {
            BlockChange::Insert => RowOp::Insert { row: row.clone() },
            BlockChange::Delete => RowOp::Delete { row: row.clone() },
        };
        emit_row(body, &row_change, rev, opts);
    }
    body.push_str("</w:tbl>");
    body.push_str("<w:p/>"); // spacer paragraph after table (OOXML requires a p after tbl)
}

fn emit_table_diff(body: &mut String, rows: &[RowOp], rev: &mut RevId, opts: &RedlineOptions) {
    let col_count = rows.iter().map(|r| row_cells(r).len()).max().unwrap_or(0);
    body.push_str(&table_open(col_count));
    for row in rows {
        emit_row(body, row, rev, opts);
    }
    body.push_str("</w:tbl>");
    body.push_str("<w:p/>");
}

fn row_cells(r: &RowOp) -> Vec<String> {
    match r {
        RowOp::Equal { row } | RowOp::Insert { row } | RowOp::Delete { row } => {
            row.cells.iter().map(|c| c.text()).collect()
        }
        RowOp::Modified { cells } => cells
            .iter()
            .map(|c| match c {
                CellOp::Equal { text } => text.clone(),
                CellOp::Changed { changes } => changes.iter().map(|ch| ch.text.as_str()).collect(),
            })
            .collect(),
    }
}

fn emit_row(body: &mut String, row: &RowOp, rev: &mut RevId, opts: &RedlineOptions) {
    match row {
        RowOp::Equal { row } => {
            body.push_str("<w:tr>");
            for cell in &row.cells {
                body.push_str(&cell_equal(cell));
            }
            body.push_str("</w:tr>");
        }
        RowOp::Insert { row } => {
            body.push_str(&tr_open_with_marker(BlockChange::Insert, rev, opts));
            for cell in &row.cells {
                body.push_str(&cell_whole(cell, BlockChange::Insert, rev, opts));
            }
            body.push_str("</w:tr>");
        }
        RowOp::Delete { row } => {
            body.push_str(&tr_open_with_marker(BlockChange::Delete, rev, opts));
            for cell in &row.cells {
                body.push_str(&cell_whole(cell, BlockChange::Delete, rev, opts));
            }
            body.push_str("</w:tr>");
        }
        RowOp::Modified { cells } => {
            body.push_str("<w:tr>");
            for cop in cells {
                body.push_str(&cell_modified(cop, rev, opts));
            }
            body.push_str("</w:tr>");
        }
    }
}

fn tr_open_with_marker(change: BlockChange, rev: &mut RevId, opts: &RedlineOptions) -> String {
    // Row-level track change marker. For color style we still use the marker so
    // row count changes are unambiguous; cell content styling makes the visual.
    if opts.style == RedlineStyle::TrackChange {
        let id = rev.next();
        let tag = match change {
            BlockChange::Insert => "ins",
            BlockChange::Delete => "del",
        };
        format!(
            r#"<w:tr><w:trPr><w:{tag} w:id="{}" w:author="{}" w:date="{}"/></w:trPr>"#,
            id,
            xml_escape(&opts.author),
            xml_escape(&opts.date),
        )
    } else {
        "<w:tr>".to_string()
    }
}

/// Emit an unchanged cell with its original paragraphs verbatim (pPr + runs),
/// preserving every font/style detail from the source DOCX.
fn cell_equal(cell: &compare_core::TableCell) -> String {
    let mut out = String::from("<w:tc>");
    if cell.paragraphs.is_empty() {
        out.push_str("<w:p/>");
    } else {
        for para in &cell.paragraphs {
            out.push_str("<w:p>");
            if !para.ppr_xml.is_empty() {
                out.push_str(&para.ppr_xml);
            }
            if para.runs.is_empty() {
                out.push_str(&run(&para.text, None));
            } else {
                for r in &para.runs {
                    out.push_str("<w:r>");
                    if !r.rpr_xml.is_empty() {
                        out.push_str(&r.rpr_xml);
                    }
                    out.push_str(&format!(
                        r#"<w:t xml:space="preserve">{}</w:t></w:r>"#,
                        xml_escape(&r.text)
                    ));
                }
            }
            out.push_str("</w:p>");
        }
    }
    out.push_str("</w:tc>");
    out
}

/// Emit a whole-cell insert or delete (every paragraph/run in the cell gets
/// the diff styling). Paragraph structure is preserved; only the visual is
/// changed to indicate the change type.
fn cell_whole(
    cell: &compare_core::TableCell,
    change: BlockChange,
    rev: &mut RevId,
    opts: &RedlineOptions,
) -> String {
    let mut out = String::from("<w:tc>");
    if cell.paragraphs.is_empty() {
        out.push_str("<w:p/>");
    } else {
        for para in &cell.paragraphs {
            out.push_str("<w:p>");
            if !para.ppr_xml.is_empty() {
                out.push_str(&para.ppr_xml);
            }
            out.push_str(&cell_whole_runs(&para.text, change, rev, opts));
            out.push_str("</w:p>");
        }
    }
    out.push_str("</w:tc>");
    out
}

fn cell_whole_runs(text: &str, change: BlockChange, rev: &mut RevId, opts: &RedlineOptions) -> String {
    match (opts.style, change) {
        (RedlineStyle::Color, BlockChange::Insert) => styled_run(text, COLOR_INSERT, RunStyle::Underline),
        (RedlineStyle::Color, BlockChange::Delete) => styled_run(text, COLOR_DELETE, RunStyle::Strike),
        (RedlineStyle::TrackChange, BlockChange::Insert) => {
            ins_wrap(run(text, Some(COLOR_INSERT)), rev.next(), opts)
        }
        (RedlineStyle::TrackChange, BlockChange::Delete) => {
            del_wrap(del_run(text, Some(COLOR_DELETE)), rev.next(), opts)
        }
    }
}

fn cell_modified(cell: &CellOp, rev: &mut RevId, opts: &RedlineOptions) -> String {
    let body = match cell {
        CellOp::Equal { text } => run(text, None),
        CellOp::Changed { changes } => {
            let mut inner = String::new();
            for ch in changes {
                match opts.style {
                    RedlineStyle::Color => match ch.kind {
                        ChangeKind::Equal => inner.push_str(&run(&ch.text, None)),
                        ChangeKind::Insert => inner.push_str(&styled_run(&ch.text, COLOR_INSERT, RunStyle::Underline)),
                        ChangeKind::Delete => inner.push_str(&styled_run(&ch.text, COLOR_DELETE, RunStyle::Strike)),
                    },
                    RedlineStyle::TrackChange => match ch.kind {
                        ChangeKind::Equal => inner.push_str(&run(&ch.text, None)),
                        ChangeKind::Insert => {
                            inner.push_str(&ins_wrap(run(&ch.text, Some(COLOR_INSERT)), rev.next(), opts));
                        }
                        ChangeKind::Delete => {
                            inner.push_str(&del_wrap(del_run(&ch.text, Some(COLOR_DELETE)), rev.next(), opts));
                        }
                    },
                }
            }
            inner
        }
    };
    format!(r#"<w:tc><w:p>{}</w:p></w:tc>"#, body)
}

fn table_open(col_count: usize) -> String {
    let mut grid = String::new();
    for _ in 0..col_count {
        grid.push_str("<w:gridCol/>");
    }
    format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders><w:top w:val="single" w:sz="4" w:color="auto"/><w:left w:val="single" w:sz="4" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:color="auto"/><w:right w:val="single" w:sz="4" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:color="auto"/></w:tblBorders></w:tblPr><w:tblGrid>{}</w:tblGrid>"#,
        grid
    )
}

// ---------- paragraph op emission ----------

fn emit_para_op(body: &mut String, op: &ParaOp, rev: &mut RevId, opts: &RedlineOptions) {
    match opts.style {
        RedlineStyle::Color => emit_para_op_color(body, op),
        RedlineStyle::TrackChange => emit_para_op_track(body, op, rev, opts),
    }
}

fn emit_para_op_color(body: &mut String, op: &ParaOp) {
    match op {
        ParaOp::Equal { text } => body.push_str(&run(text, None)),
        ParaOp::Insert { text } => body.push_str(&styled_run(text, COLOR_INSERT, RunStyle::Underline)),
        ParaOp::Delete { text } => body.push_str(&styled_run(text, COLOR_DELETE, RunStyle::Strike)),
        ParaOp::Modified { changes } => {
            for ch in changes {
                match ch.kind {
                    ChangeKind::Equal => body.push_str(&run(&ch.text, None)),
                    ChangeKind::Insert => body.push_str(&styled_run(&ch.text, COLOR_INSERT, RunStyle::Underline)),
                    ChangeKind::Delete => body.push_str(&styled_run(&ch.text, COLOR_DELETE, RunStyle::Strike)),
                }
            }
        }
        ParaOp::MovedFrom { text, .. } => {
            body.push_str(&styled_run(text, COLOR_MOVE, RunStyle::DoubleStrike));
        }
        ParaOp::MovedTo { text, .. } => {
            body.push_str(&styled_run(text, COLOR_MOVE, RunStyle::DoubleUnderline));
        }
    }
}

fn emit_para_op_track(body: &mut String, op: &ParaOp, rev: &mut RevId, opts: &RedlineOptions) {
    // Multi-author trick: Word/LibreOffice color track-change marks per
    // author. By issuing each change type under a distinct author name, the
    // reader's palette assigns a distinct color to adds vs deletes vs moves —
    // so the three kinds are always visually distinguishable even when our
    // baked <w:color> is overridden by the track-change display.
    let author_ins = change_author(&opts.author, "추가");
    let author_del = change_author(&opts.author, "삭제");
    let author_mov = change_author(&opts.author, "이동");

    match op {
        ParaOp::Equal { text } => body.push_str(&run(text, None)),
        ParaOp::Insert { text } => {
            body.push_str(&ins_wrap_as(run(text, Some(COLOR_INSERT)), rev.next(), &author_ins, &opts.date));
        }
        ParaOp::Delete { text } => {
            body.push_str(&del_wrap_as(del_run(text, Some(COLOR_DELETE)), rev.next(), &author_del, &opts.date));
        }
        ParaOp::Modified { changes } => {
            for ch in changes {
                match ch.kind {
                    ChangeKind::Equal => body.push_str(&run(&ch.text, None)),
                    ChangeKind::Insert => {
                        body.push_str(&ins_wrap_as(run(&ch.text, Some(COLOR_INSERT)), rev.next(), &author_ins, &opts.date));
                    }
                    ChangeKind::Delete => {
                        body.push_str(&del_wrap_as(del_run(&ch.text, Some(COLOR_DELETE)), rev.next(), &author_del, &opts.date));
                    }
                }
            }
        }
        ParaOp::MovedFrom { text, move_id } => {
            body.push_str(&move_from_wrap_as(run(text, Some(COLOR_MOVE)), rev.next(), *move_id, &author_mov, &opts.date));
        }
        ParaOp::MovedTo { text, move_id } => {
            body.push_str(&move_to_wrap_as(run(text, Some(COLOR_MOVE)), rev.next(), *move_id, &author_mov, &opts.date));
        }
    }
}

fn change_author(base: &str, label: &str) -> String {
    format!("{base} ({label})")
}

fn ins_wrap_as(inner: String, id: u32, author: &str, date: &str) -> String {
    format!(
        r#"<w:ins w:id="{}" w:author="{}" w:date="{}">{}</w:ins>"#,
        id, xml_escape(author), xml_escape(date), inner
    )
}

fn del_wrap_as(inner: String, id: u32, author: &str, date: &str) -> String {
    format!(
        r#"<w:del w:id="{}" w:author="{}" w:date="{}">{}</w:del>"#,
        id, xml_escape(author), xml_escape(date), inner
    )
}

fn move_from_wrap_as(inner: String, id: u32, move_id: u32, author: &str, date: &str) -> String {
    format!(
        r#"<w:moveFrom w:id="{}" w:name="move{}" w:author="{}" w:date="{}">{}</w:moveFrom>"#,
        id, move_id, xml_escape(author), xml_escape(date), inner
    )
}

fn move_to_wrap_as(inner: String, id: u32, move_id: u32, author: &str, date: &str) -> String {
    format!(
        r#"<w:moveTo w:id="{}" w:name="move{}" w:author="{}" w:date="{}">{}</w:moveTo>"#,
        id, move_id, xml_escape(author), xml_escape(date), inner
    )
}

// ---------- primitives ----------

#[derive(Clone, Copy)]
enum RunStyle {
    Underline,
    DoubleUnderline,
    Strike,
    DoubleStrike,
}

fn styled_run(text: &str, color: &str, style: RunStyle) -> String {
    let tag = match style {
        RunStyle::Underline => r#"<w:u w:val="single"/>"#,
        RunStyle::DoubleUnderline => r#"<w:u w:val="double"/>"#,
        RunStyle::Strike => r#"<w:strike/>"#,
        RunStyle::DoubleStrike => r#"<w:dstrike/>"#,
    };
    format!(
        r#"<w:r><w:rPr><w:color w:val="{}"/>{}</w:rPr><w:t xml:space="preserve">{}</w:t></w:r>"#,
        color, tag, xml_escape(text)
    )
}

// Legacy single-author wraps are kept only as thin shims over the per-change
// author path below, used by the color-mode table cell renderer that still
// passes `opts` through.
fn ins_wrap(inner: String, id: u32, opts: &RedlineOptions) -> String {
    ins_wrap_as(inner, id, &opts.author, &opts.date)
}
fn del_wrap(inner: String, id: u32, opts: &RedlineOptions) -> String {
    del_wrap_as(inner, id, &opts.author, &opts.date)
}

fn run(text: &str, color: Option<&str>) -> String {
    format!(
        r#"<w:r>{}<w:t xml:space="preserve">{}</w:t></w:r>"#,
        rpr(color),
        xml_escape(text)
    )
}

fn del_run(text: &str, color: Option<&str>) -> String {
    format!(
        r#"<w:r>{}<w:delText xml:space="preserve">{}</w:delText></w:r>"#,
        rpr(color),
        xml_escape(text)
    )
}

fn rpr(color: Option<&str>) -> String {
    match color {
        Some(hex) => format!(r#"<w:rPr><w:color w:val="{}"/></w:rPr>"#, hex),
        None => String::new(),
    }
}

fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

// ---------- summary page ----------

fn summary_page(body_stats: &Stats, opts: &RedlineOptions, _rev: &mut RevId) -> String {
    let mut out = String::new();
    out.push_str(r#"<w:p><w:r><w:br w:type="page"/></w:r></w:p>"#);
    out.push_str(
        r#"<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="36"/></w:rPr><w:t xml:space="preserve">변경 사항 요약</w:t></w:r></w:p>"#,
    );
    out.push_str("<w:p/>");

    // Metadata table.
    let original_name = opts.original_name.as_deref().unwrap_or("(unknown)");
    let modified_name = opts.modified_name.as_deref().unwrap_or("(unknown)");
    out.push_str(TABLE_SUMMARY_OPEN);
    out.push_str(&sum_row("원본 파일 (Original)", original_name, None, false));
    out.push_str(&sum_row("수정본 파일 (Modified)", modified_name, None, false));
    out.push_str(&sum_row("검토자 (Reviewer)", &opts.author, None, false));
    out.push_str("</w:tbl>");
    out.push_str("<w:p/>");

    // Changes table — body-only counts, no duplicates, no strikethrough.
    let total_changes = body_stats.words_inserted
        + body_stats.words_deleted
        + body_stats.words_moved
        + body_stats.tables_inserted
        + body_stats.tables_deleted
        + body_stats.rows_inserted
        + body_stats.rows_deleted;

    out.push_str(TABLE_SUMMARY_OPEN);
    out.push_str(&sum_row("변경 항목 (Changes)", "건수 (Count)", None, true));

    let rows: [(&str, String, Option<&'static str>); 8] = [
        ("추가 (Add)",                body_stats.words_inserted.to_string(),  Some(COLOR_INSERT)),
        ("삭제 (Delete)",             body_stats.words_deleted.to_string(),   Some(COLOR_DELETE)),
        ("표 삽입 (Table Insert)",    body_stats.tables_inserted.to_string(), Some(COLOR_INSERT)),
        ("표 삭제 (Table Delete)",    body_stats.tables_deleted.to_string(),  Some(COLOR_DELETE)),
        ("이동 (Move)",               body_stats.words_moved.to_string(),     Some(COLOR_MOVE)),
        ("표 이동 (Table Move)",      "0".to_string(),                        Some(COLOR_MOVE)),
        ("서식 변경 (Format)",        "0".to_string(),                        None),
        ("총 변경 (Total)",           total_changes.to_string(),              None),
    ];
    let last_idx = rows.len() - 1;
    for (i, (label, value, color)) in rows.iter().enumerate() {
        let bold = i == last_idx;
        out.push_str(&sum_row(label, value, *color, bold));
    }
    out.push_str("</w:tbl>");
    out.push_str("<w:p/>");

    out.push_str(&legend_paragraph());

    out
}

const TABLE_SUMMARY_OPEN: &str = r#"<w:tbl>
<w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>
<w:top w:val="single" w:sz="4" w:color="auto"/>
<w:left w:val="single" w:sz="4" w:color="auto"/>
<w:bottom w:val="single" w:sz="4" w:color="auto"/>
<w:right w:val="single" w:sz="4" w:color="auto"/>
<w:insideH w:val="single" w:sz="4" w:color="auto"/>
<w:insideV w:val="single" w:sz="4" w:color="auto"/>
</w:tblBorders></w:tblPr>
<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="2000"/></w:tblGrid>"#;

fn sum_row(a: &str, b: &str, color: Option<&str>, bold: bool) -> String {
    format!(
        "<w:tr>{}{}</w:tr>",
        sum_cell(a, color, bold),
        sum_cell(b, color, bold),
    )
}

fn sum_cell(text: &str, color: Option<&str>, bold: bool) -> String {
    let rpr = match (color, bold) {
        (Some(c), true) => format!(r#"<w:rPr><w:b/><w:color w:val="{}"/></w:rPr>"#, c),
        (Some(c), false) => format!(r#"<w:rPr><w:color w:val="{}"/></w:rPr>"#, c),
        (None, true) => r#"<w:rPr><w:b/></w:rPr>"#.to_string(),
        (None, false) => String::new(),
    };
    format!(
        r#"<w:tc><w:p><w:r>{}<w:t xml:space="preserve">{}</w:t></w:r></w:p></w:tc>"#,
        rpr,
        xml_escape(text)
    )
}

fn legend_paragraph() -> String {
    format!(
        "<w:p>{}{}{}</w:p>",
        format!(r#"<w:r><w:rPr><w:color w:val="{}"/></w:rPr><w:t xml:space="preserve">■ 추가  </w:t></w:r>"#, COLOR_INSERT),
        format!(r#"<w:r><w:rPr><w:color w:val="{}"/></w:rPr><w:t xml:space="preserve">■ 삭제  </w:t></w:r>"#, COLOR_DELETE),
        format!(r#"<w:r><w:rPr><w:color w:val="{}"/></w:rPr><w:t xml:space="preserve">■ 이동</w:t></w:r>"#, COLOR_MOVE),
    )
}

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;
