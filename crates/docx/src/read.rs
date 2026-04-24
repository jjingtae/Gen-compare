//! DOCX reader. Captures the full block structure plus per-paragraph pPr and
//! per-run rPr so the writer can preserve formatting on unchanged paragraphs.

use anyhow::{Context, Result};
use compare_core::{Block, RichParagraph, RichRun, TableCell, TableRow};
use quick_xml::events::Event;
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;
use std::fs::File;
use std::io::{Cursor, Read};
use std::path::Path;

#[derive(Debug, Clone)]
pub struct Paragraph {
    pub text: String,
}

#[derive(Debug, Clone, Default)]
pub struct DocxContent {
    pub body: Vec<Block>,
    pub headers: Vec<String>,
    pub footers: Vec<String>,
    /// Comment contents (one string per comment paragraph).
    pub comments: Vec<String>,
    /// Footnote contents. Separators (system notes) are excluded.
    pub footnotes: Vec<String>,
    /// Endnote contents. Separators (system notes) are excluded.
    pub endnotes: Vec<String>,
    /// dc:creator from docProps/core.xml (original author).
    pub creator: Option<String>,
    /// cp:lastModifiedBy from docProps/core.xml — the user we attribute
    /// redline authorship to by default.
    pub last_modified_by: Option<String>,
    pub raw_parts: RawParts,
}

#[derive(Debug, Clone, Default)]
pub struct RawParts {
    pub styles_xml: Option<Vec<u8>>,
    pub numbering_xml: Option<Vec<u8>>,
    pub theme1_xml: Option<Vec<u8>>,
    pub font_table_xml: Option<Vec<u8>>,
    pub settings_xml: Option<Vec<u8>>,
    pub web_settings_xml: Option<Vec<u8>>,
    pub style_with_effects_xml: Option<Vec<u8>>,
    pub comments_xml: Option<Vec<u8>>,
    pub footnotes_xml: Option<Vec<u8>>,
    pub endnotes_xml: Option<Vec<u8>>,
    /// header1.xml, header2.xml, header3.xml etc. as (name, bytes)
    /// Binary resources referenced by header/footer rels (images in
    /// `word/media/*`, embedded objects in `word/embeddings/*`, embedded
    /// fonts in `word/fonts/*`). Copied verbatim into the output so that
    /// rels don't become dangling → Word would otherwise flag the output
    /// as "corrupted".
    pub binary_resources: Vec<(String, Vec<u8>)>,
    pub headers: Vec<(String, Vec<u8>)>,
    /// footer1.xml, footer2.xml, footer3.xml etc.
    pub footers: Vec<(String, Vec<u8>)>,
    /// header/footer rels (e.g. header1.xml.rels) needed for hyperlinks/images inside.
    pub header_footer_rels: Vec<(String, Vec<u8>)>,
}

pub fn read_paragraphs<P: AsRef<Path>>(path: P) -> Result<Vec<Paragraph>> {
    let c = read_document(path)?;
    let mut out = Vec::new();
    for b in c.body {
        match b {
            Block::Paragraph(p) => out.push(Paragraph { text: p.text }),
            Block::Table { rows } => {
                for r in rows {
                    for cell in r.cells {
                        out.push(Paragraph { text: cell.text() });
                    }
                }
            }
        }
    }
    Ok(out)
}

pub fn read_document<P: AsRef<Path>>(path: P) -> Result<DocxContent> {
    let file = File::open(path.as_ref())
        .with_context(|| format!("open docx {}", path.as_ref().display()))?;
    let mut zip = zip::ZipArchive::new(file).context("parse docx as zip")?;

    let body = {
        let mut xml = String::new();
        zip.by_name("word/document.xml")
            .context("docx missing word/document.xml")?
            .read_to_string(&mut xml)
            .context("read document.xml")?;
        parse_blocks(&xml)?
    };

    let mut headers = Vec::new();
    let mut footers = Vec::new();
    let mut raw_headers = Vec::new();
    let mut raw_footers = Vec::new();
    let mut raw_hf_rels = Vec::new();
    let mut binary_resources: Vec<(String, Vec<u8>)> = Vec::new();
    let names: Vec<String> = zip.file_names().map(|s| s.to_string()).collect();
    for name in &names {
        // Collect binary resources (images, embedded objects, fonts) — these
        // are referenced by header/footer rels and must exist in the output
        // or Word refuses to open the file.
        if name.starts_with("word/media/")
            || name.starts_with("word/embeddings/")
            || name.starts_with("word/fonts/")
        {
            if let Some(bytes) = read_bytes(&mut zip, name) {
                binary_resources.push((name.clone(), bytes));
            }
            continue;
        }
        if name.starts_with("word/header") && name.ends_with(".xml") {
            if let Some(bytes) = read_bytes(&mut zip, name) {
                if let Ok(ps) = parse_flat_paragraphs(&bytes) {
                    headers.extend(ps);
                }
                raw_headers.push((name.clone(), bytes));
            }
        } else if name.starts_with("word/footer") && name.ends_with(".xml") {
            if let Some(bytes) = read_bytes(&mut zip, name) {
                if let Ok(ps) = parse_flat_paragraphs(&bytes) {
                    footers.extend(ps);
                }
                raw_footers.push((name.clone(), bytes));
            }
        } else if name.starts_with("word/_rels/header")
            || name.starts_with("word/_rels/footer")
        {
            if let Some(bytes) = read_bytes(&mut zip, name) {
                raw_hf_rels.push((name.clone(), bytes));
            }
        }
    }

    let core_xml = read_bytes(&mut zip, "docProps/core.xml");
    let (creator, last_modified_by) = core_xml
        .as_deref()
        .map(parse_core_props)
        .unwrap_or((None, None));

    let comments_xml = read_bytes(&mut zip, "word/comments.xml");
    let footnotes_xml = read_bytes(&mut zip, "word/footnotes.xml");
    let endnotes_xml = read_bytes(&mut zip, "word/endnotes.xml");

    let comments = comments_xml
        .as_deref()
        .and_then(|b| parse_noted_paragraphs(b, b"comment").ok())
        .unwrap_or_default();
    let footnotes = footnotes_xml
        .as_deref()
        .and_then(|b| parse_noted_paragraphs(b, b"footnote").ok())
        .unwrap_or_default();
    let endnotes = endnotes_xml
        .as_deref()
        .and_then(|b| parse_noted_paragraphs(b, b"endnote").ok())
        .unwrap_or_default();

    let raw_parts = RawParts {
        styles_xml: read_bytes(&mut zip, "word/styles.xml"),
        numbering_xml: read_bytes(&mut zip, "word/numbering.xml"),
        theme1_xml: read_bytes(&mut zip, "word/theme/theme1.xml"),
        font_table_xml: read_bytes(&mut zip, "word/fontTable.xml"),
        settings_xml: read_bytes(&mut zip, "word/settings.xml"),
        web_settings_xml: read_bytes(&mut zip, "word/webSettings.xml"),
        style_with_effects_xml: read_bytes(&mut zip, "word/stylesWithEffects.xml"),
        comments_xml,
        footnotes_xml,
        endnotes_xml,
        headers: raw_headers,
        footers: raw_footers,
        header_footer_rels: raw_hf_rels,
        binary_resources,
    };

    Ok(DocxContent { body, headers, footers, comments, footnotes, endnotes, creator, last_modified_by, raw_parts })
}

/// Parse docProps/core.xml for dc:creator and cp:lastModifiedBy.
fn parse_core_props(xml_bytes: &[u8]) -> (Option<String>, Option<String>) {
    let Ok(xml) = std::str::from_utf8(xml_bytes) else { return (None, None); };
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut creator = None;
    let mut last_mod = None;
    let mut current: Option<&'static str> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let local = e.local_name().as_ref().to_owned();
                current = match local.as_slice() {
                    b"creator" => Some("creator"),
                    b"lastModifiedBy" => Some("lastModifiedBy"),
                    _ => None,
                };
            }
            Ok(Event::End(_)) => current = None,
            Ok(Event::Text(t)) => {
                if let Some(which) = current {
                    if let Ok(s) = t.unescape() {
                        let v = s.trim().to_string();
                        if !v.is_empty() {
                            match which {
                                "creator" => creator = Some(v),
                                "lastModifiedBy" => last_mod = Some(v),
                                _ => {}
                            }
                        }
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    (creator, last_mod)
}

/// Extract paragraph text from a comments/footnotes/endnotes XML part.
/// Wrapper element is `w:comment`, `w:footnote`, or `w:endnote`. Footnote/
/// endnote parts also contain system "separator" items which are excluded.
fn parse_noted_paragraphs(xml_bytes: &[u8], wrapper: &[u8]) -> Result<Vec<String>> {
    let xml = std::str::from_utf8(xml_bytes)?;
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    let mut out = Vec::new();
    let mut in_wrapper = false;
    let mut skip_wrapper = false;
    let mut in_p = false;
    let mut in_t = false;
    let mut current = String::new();

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let name = e.local_name().as_ref().to_owned();
                if name == wrapper {
                    in_wrapper = true;
                    // Exclude system notes (w:type="separator" etc.)
                    skip_wrapper = e
                        .attributes()
                        .flatten()
                        .any(|a| a.key.local_name().as_ref() == b"type");
                } else if in_wrapper && !skip_wrapper {
                    match name.as_slice() {
                        b"p" => {
                            in_p = true;
                            current.clear();
                        }
                        b"t" => in_t = true,
                        b"tab" => current.push('\t'),
                        b"br" => current.push('\n'),
                        _ => {}
                    }
                }
            }
            Event::End(e) => {
                let name = e.local_name().as_ref().to_owned();
                if name == wrapper {
                    in_wrapper = false;
                    skip_wrapper = false;
                } else if in_wrapper && !skip_wrapper {
                    match name.as_slice() {
                        b"p" => {
                            if in_p {
                                out.push(std::mem::take(&mut current));
                                in_p = false;
                            }
                        }
                        b"t" => in_t = false,
                        _ => {}
                    }
                }
            }
            Event::Text(t) if in_t => {
                current.push_str(&t.unescape()?);
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(out)
}

fn read_bytes<R: Read + std::io::Seek>(
    zip: &mut zip::ZipArchive<R>,
    name: &str,
) -> Option<Vec<u8>> {
    let mut f = zip.by_name(name).ok()?;
    let mut buf = Vec::new();
    std::io::Read::read_to_end(&mut f, &mut buf).ok()?;
    Some(buf)
}

fn parse_flat_paragraphs(xml_bytes: &[u8]) -> Result<Vec<String>> {
    let xml = std::str::from_utf8(xml_bytes)?;
    let blocks = parse_blocks(xml)?;
    let mut out = Vec::new();
    for b in blocks {
        match b {
            Block::Paragraph(p) => out.push(p.text),
            Block::Table { rows } => {
                for r in rows {
                    for c in r.cells {
                        out.push(c.text());
                    }
                }
            }
        }
    }
    Ok(out)
}

/// Parse the main XML body into blocks with pPr + rPr captured verbatim.
fn parse_blocks(xml: &str) -> Result<Vec<Block>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    let mut blocks: Vec<Block> = Vec::new();
    let mut table_stack: Vec<Vec<TableRow>> = Vec::new();
    let mut row_stack: Vec<Vec<TableCell>> = Vec::new();
    // Rich paragraphs collected for the currently-open <w:tc>.
    let mut cell_paras: Vec<RichParagraph> = Vec::new();
    let mut tc_depth: u32 = 0;

    // Current paragraph being built.
    let mut in_p = false;
    let mut cur_ppr: String = String::new();
    let mut cur_runs: Vec<RichRun> = Vec::new();

    // Current run being built.
    let mut in_r = false;
    let mut cur_rpr: String = String::new();
    let mut cur_run_text = String::new();

    // Text accumulator (inside <w:t>).
    let mut in_t = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                match e.local_name().as_ref() {
                    b"tbl" => table_stack.push(Vec::new()),
                    b"tr" => row_stack.push(Vec::new()),
                    b"tc" => {
                        tc_depth += 1;
                        cell_paras.clear();
                    }
                    b"p" => {
                        in_p = true;
                        cur_ppr.clear();
                        cur_runs.clear();
                    }
                    b"pPr" => {
                        cur_ppr = capture_element_full(&mut reader, "pPr")?;
                    }
                    b"r" => {
                        in_r = true;
                        cur_rpr.clear();
                        cur_run_text.clear();
                    }
                    b"rPr" if in_r => {
                        cur_rpr = capture_element_full(&mut reader, "rPr")?;
                    }
                    b"t" => in_t = true,
                    b"tab" => {
                        if in_r {
                            cur_run_text.push('\t');
                        }
                    }
                    b"br" => {
                        if in_r {
                            cur_run_text.push('\n');
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => match e.local_name().as_ref() {
                b"p" => {
                    if in_p {
                        let text: String = cur_runs.iter().map(|r| r.text.as_str()).collect();
                        let ppr = std::mem::take(&mut cur_ppr);
                        let runs = std::mem::take(&mut cur_runs);
                        let rp = RichParagraph { text, ppr_xml: ppr, runs };
                        if tc_depth > 0 {
                            cell_paras.push(rp);
                        } else {
                            blocks.push(Block::Paragraph(rp));
                        }
                        in_p = false;
                    }
                }
                b"r" => {
                    if in_r {
                        let rpr = std::mem::take(&mut cur_rpr);
                        let text = std::mem::take(&mut cur_run_text);
                        if !text.is_empty() || !rpr.is_empty() {
                            cur_runs.push(RichRun { rpr_xml: rpr, text });
                        }
                        in_r = false;
                    }
                }
                b"t" => in_t = false,
                b"tc" => {
                    tc_depth -= 1;
                    let paragraphs = std::mem::take(&mut cell_paras);
                    if let Some(cells) = row_stack.last_mut() {
                        cells.push(TableCell { paragraphs });
                    }
                }
                b"tr" => {
                    if let Some(cells) = row_stack.pop() {
                        if let Some(rows) = table_stack.last_mut() {
                            rows.push(TableRow { cells });
                        }
                    }
                }
                b"tbl" => {
                    if let Some(rows) = table_stack.pop() {
                        blocks.push(Block::Table { rows });
                    }
                }
                _ => {}
            },
            Event::Text(t) if in_t && in_r => {
                cur_run_text.push_str(&t.unescape()?);
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(blocks)
}

/// Capture the full `<w:tag>...</w:tag>` (including opening and closing tags)
/// from the stream, assuming the Start event has already been consumed.
/// Used to preserve pPr/rPr fragments verbatim for later re-emission.
fn capture_element_full(
    reader: &mut Reader<&[u8]>,
    local: &str,
) -> Result<String> {
    let full = format!("w:{}", local);
    let local_bytes = local.as_bytes();
    let mut buf = Vec::new();
    {
        let mut w = Writer::new(Cursor::new(&mut buf));
        w.write_event(Event::Start(quick_xml::events::BytesStart::new(
            full.as_str(),
        )))?;
        let mut depth: u32 = 1;
        loop {
            let ev = reader.read_event()?;
            match &ev {
                Event::Start(e) if e.local_name().as_ref() == local_bytes => depth += 1,
                Event::End(e) if e.local_name().as_ref() == local_bytes => {
                    depth -= 1;
                    if depth == 0 {
                        w.write_event(Event::End(quick_xml::events::BytesEnd::new(
                            full.as_str(),
                        )))?;
                        break;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            if depth > 0 {
                w.write_event(ev)?;
            }
        }
    }
    Ok(String::from_utf8(buf)?)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_paragraph_with_ppr_and_rpr() {
        let xml = r#"<w:document xmlns:w="x"><w:body>
            <w:p>
              <w:pPr><w:jc w:val="center"/></w:pPr>
              <w:r><w:rPr><w:b/></w:rPr><w:t>Hello</w:t></w:r>
              <w:r><w:t> world</w:t></w:r>
            </w:p>
        </w:body></w:document>"#;
        let blocks = parse_blocks(xml).unwrap();
        assert_eq!(blocks.len(), 1);
        if let Block::Paragraph(p) = &blocks[0] {
            assert_eq!(p.text, "Hello world");
            assert!(p.ppr_xml.contains("jc"));
            assert_eq!(p.runs.len(), 2);
            assert!(p.runs[0].rpr_xml.contains("<w:b"));
        } else {
            panic!("expected paragraph");
        }
    }
}
