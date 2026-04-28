#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

//! Compare GUI — thin Tauri shell over the core engine.
//!
//! The frontend (dist/index.html) handles drag-and-drop, pair display, and
//! save-format button state. It calls into these Rust commands which do the
//! actual work using compare-core + compare-docx (same engine as the CLI).

use anyhow::Context;
use compare_core::{detect_pairs as core_detect_pairs, diff_blocks, stats_of_blocks, BlockOp, Stats};
use compare_docx::{read_document, write_redline, RedlineOptions, RedlineStyle};
use rayon::prelude::*;
use serde::{Deserialize, Serialize};
use std::path::{Path, PathBuf};
use std::sync::Mutex;

// Word COM and LibreOffice headless both behave poorly when called
// concurrently from the same user session (Word.Application is effectively a
// per-user singleton; soffice shares its user profile). We run diff/DOCX work
// in parallel but serialize the PDF step behind this mutex.
static PDF_LOCK: Mutex<()> = Mutex::new(());

#[derive(Serialize)]
struct PairSuggestion {
    old: String,
    new: String,
    base: String,
    reason: String,
}

#[derive(Serialize)]
struct DetectionResult {
    pairs: Vec<PairSuggestion>,
    unpaired: Vec<String>,
}

/// Given a set of file paths dropped by the user, suggest (old, new) pairs.
#[tauri::command]
fn detect_pairs(paths: Vec<String>) -> DetectionResult {
    let bufs: Vec<PathBuf> = paths.into_iter().map(PathBuf::from).collect();
    let det = core_detect_pairs(&bufs);
    DetectionResult {
        pairs: det
            .pairs
            .into_iter()
            .map(|m| PairSuggestion {
                old: m.old.to_string_lossy().into_owned(),
                new: m.new.to_string_lossy().into_owned(),
                base: m.base,
                reason: m.reason,
            })
            .collect(),
        unpaired: det
            .unpaired
            .into_iter()
            .map(|p| p.to_string_lossy().into_owned())
            .collect(),
    }
}

#[derive(Deserialize, Clone)]
struct RunPair {
    old: String,
    new: String,
    /// Output base name (no extension). Extensions are added per selected format.
    out_base: String,
    /// Output directory.
    out_dir: String,
}

#[derive(Deserialize, Clone, Copy)]
struct Outputs {
    /// Color DOCX (visible red/blue/green).
    word: bool,
    /// Track change DOCX (Word-style revisions).
    track_change: bool,
    /// PDF (rendered from color DOCX).
    pdf: bool,
    /// CPO — PDF containing only blocks that have changes.
    cpo: bool,
}

#[derive(Serialize)]
struct PairResult {
    pair_index: usize,
    old: String,
    new: String,
    outputs: Vec<String>,
    stats: Option<Stats>,
    error: Option<String>,
    elapsed_ms: u128,
}

#[derive(Serialize)]
struct RunReport {
    total_elapsed_ms: u128,
    workers: usize,
    results: Vec<PairResult>,
}

/// Run N pairs in parallel with the selected output formats.
#[tauri::command]
fn run_batch(pairs: Vec<RunPair>, outputs: Outputs, author: Option<String>) -> RunReport {
    let t0 = std::time::Instant::now();
    let results: Vec<PairResult> = pairs
        .par_iter()
        .enumerate()
        .map(|(i, p)| run_one_ui(i, p, outputs, author.as_deref()))
        .collect();
    RunReport {
        total_elapsed_ms: t0.elapsed().as_millis(),
        workers: rayon::current_num_threads(),
        results,
    }
}

fn run_one_ui(i: usize, p: &RunPair, outputs: Outputs, author: Option<&str>) -> PairResult {
    let t0 = std::time::Instant::now();
    match try_run_one(p, outputs, author) {
        Ok((outs, stats)) => PairResult {
            pair_index: i,
            old: p.old.clone(),
            new: p.new.clone(),
            outputs: outs,
            stats: Some(stats),
            error: None,
            elapsed_ms: t0.elapsed().as_millis(),
        },
        Err(e) => PairResult {
            pair_index: i,
            old: p.old.clone(),
            new: p.new.clone(),
            outputs: vec![],
            stats: None,
            error: Some(format!("{:#}", e)),
            elapsed_ms: t0.elapsed().as_millis(),
        },
    }
}

fn try_run_one(p: &RunPair, outputs: Outputs, author: Option<&str>) -> anyhow::Result<(Vec<String>, Stats)> {
    let old = Path::new(&p.old);
    let new = Path::new(&p.new);
    let out_dir = Path::new(&p.out_dir);
    std::fs::create_dir_all(out_dir).context("create output directory")?;

    let (old_docx, _old_tmp) = ensure_docx(old)?;
    let (new_docx, _new_tmp) = ensure_docx(new)?;

    let od = read_document(&old_docx)?;
    let nd = read_document(&new_docx)?;
    let body_ops: Vec<BlockOp> = diff_blocks(&od.body, &nd.body);

    let resolved_author: Option<String> = author
        .filter(|s| !s.trim().is_empty())
        .map(|s| s.to_string())
        .or_else(|| nd.last_modified_by.clone())
        .or_else(|| nd.creator.clone());

    let original_name = old.file_name().map(|s| s.to_string_lossy().into_owned());
    let modified_name = new.file_name().map(|s| s.to_string_lossy().into_owned());

    let base_opts = |style: RedlineStyle| {
        let mut o = RedlineOptions::default();
        o.style = style;
        if let Some(a) = &resolved_author { o.author = a.clone(); }
        o.original_name = original_name.clone();
        o.modified_name = modified_name.clone();
        o.source_parts = Some(nd.raw_parts.clone());
        o
    };

    let out_base = output_base_for_pair(old, new);
    let mut produced = Vec::new();

    if outputs.word || outputs.pdf {
        let color_path = out_dir.join(format!("Redline_{}.docx", out_base));
        write_redline(&color_path, &body_ops, &base_opts(RedlineStyle::Color))?;
        if outputs.word {
            produced.push(color_path.to_string_lossy().into_owned());
        }
        if outputs.pdf {
            let pdf_path = out_dir.join(format!("Redline_{}.pdf", out_base));
            convert_to_pdf(&color_path, &pdf_path)?;
            produced.push(pdf_path.to_string_lossy().into_owned());
            if !outputs.word {
                let _ = std::fs::remove_file(&color_path);
            }
        }
    }

    if outputs.track_change {
        let tc_path = out_dir.join(format!("TrackChange_{}.docx", out_base));
        write_redline(&tc_path, &body_ops, &base_opts(RedlineStyle::TrackChange))?;
        produced.push(tc_path.to_string_lossy().into_owned());
    }

    if outputs.cpo {
        let reuse_color = out_dir.join(format!("Redline_{}.docx", out_base));
        let (color_src, cleanup_color) = if reuse_color.exists() {
            (reuse_color.clone(), false)
        } else {
            let tmp = out_dir.join(format!("_cpo_src_{}.docx", out_base));
            write_redline(&tmp, &body_ops, &base_opts(RedlineStyle::Color))?;
            (tmp, true)
        };

        let full_pdf = out_dir.join(format!("_cpo_full_{}.pdf", out_base));
        convert_to_pdf(&color_src, &full_pdf)?;

        let cpo_path = out_dir.join(format!("CPO_{}.pdf", out_base));
        match detect_change_pages(&color_src) {
            Ok(pages) if !pages.is_empty() => {
                extract_pdf_pages(&full_pdf, &cpo_path, &pages)?;
                produced.push(cpo_path.to_string_lossy().into_owned());
            }
            Ok(_) => { /* no changes detected — skip CPO */ }
            Err(e) => {
                let _ = std::fs::remove_file(&full_pdf);
                if cleanup_color { let _ = std::fs::remove_file(&color_src); }
                return Err(e);
            }
        }

        let _ = std::fs::remove_file(&full_pdf);
        if cleanup_color { let _ = std::fs::remove_file(&color_src); }
    }

    Ok((produced, stats_of_blocks(&body_ops)))
}

/// Ask Word COM to enumerate the page numbers of every word in the color DOCX
/// that has a non-default font color — i.e. insertions/deletions/moves. Returns
/// sorted unique 1-based page numbers.
fn detect_change_pages(docx: &Path) -> anyhow::Result<Vec<u32>> {
    // Serialize: Word COM is a per-session singleton.
    let _guard = PDF_LOCK.lock().unwrap_or_else(|e| e.into_inner());

    let abs = std::fs::canonicalize(docx).unwrap_or_else(|_| docx.to_path_buf());
    // Information(3) = wdActiveEndPageNumber.
    // wdColorAutomatic = -16777216, pure black = 0 — any other Font.Color is
    // one of our change markers (ins/del/move).
    //
    // We write page numbers to a temp file (one per line) instead of StdOut
    // because cscript.exe in //B (batch) mode doesn't always flush StdOut
    // reliably, and WScript.StdOut only works under cscript (not wscript).
    const VBS: &str = r#"
On Error Resume Next
Set args = WScript.Arguments
inputPath = args(0)
outputPath = args(1)
Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.CreateTextFile(outputPath, True)
Set w = CreateObject("Word.Application")
If Err.Number <> 0 Then WScript.Quit 2
w.Visible = False
w.DisplayAlerts = 0
Set doc = w.Documents.Open(inputPath, False, True)
If Err.Number <> 0 Then w.Quit : WScript.Quit 3
doc.Repaginate
Set seen = CreateObject("Scripting.Dictionary")
For Each word In doc.Words
  c = word.Font.Color
  If c <> -16777216 And c <> 0 Then
    pg = word.Information(3)
    If Not seen.Exists(pg) Then
      seen.Add pg, True
      outFile.WriteLine pg
    End If
  End If
Next
outFile.Close
doc.Close False
w.Quit
WScript.Quit 0
"#;
    // Output to a temp .txt file.
    let nonce = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_nanos())
        .unwrap_or(0);
    let out_txt = std::env::temp_dir().join(format!("compare_pages_{}.txt", nonce));
    let out = run_vbs(VBS, &[&abs.to_string_lossy(), &out_txt.to_string_lossy()])?;
    if !out.status.success() {
        let _ = std::fs::remove_file(&out_txt);
        anyhow::bail!("Word COM via cscript failed while detecting changed pages");
    }
    let contents = std::fs::read_to_string(&out_txt).unwrap_or_default();
    let _ = std::fs::remove_file(&out_txt);
    let mut pages: Vec<u32> = contents
        .lines()
        .filter_map(|l| l.trim().parse::<u32>().ok())
        .collect();
    pages.sort_unstable();
    pages.dedup();
    Ok(pages)
}

/// Copy `src` to `dst`, keeping only the given 1-based page numbers.
/// The PDF's original layout (page size, margins, headers/footers, images)
/// is preserved — we just drop the pages we don't want.
fn extract_pdf_pages(src: &Path, dst: &Path, keep_pages: &[u32]) -> anyhow::Result<()> {
    let mut doc = lopdf::Document::load(src).context("load full PDF for CPO extraction")?;
    let total = doc.get_pages().len() as u32;
    let keep: std::collections::HashSet<u32> = keep_pages.iter().copied().collect();
    let to_delete: Vec<u32> = (1..=total).filter(|n| !keep.contains(n)).collect();
    if !to_delete.is_empty() {
        doc.delete_pages(&to_delete);
    }
    doc.save(dst).context("write CPO PDF")?;
    Ok(())
}

/// DOCX → PDF conversion. Tries in order:
///   1. MS Word COM via PowerShell (best fidelity, usually installed)
///   2. LibreOffice headless (fallback, works when Word is absent)
fn convert_to_pdf(src: &Path, dst: &Path) -> anyhow::Result<()> {
    // Serialize PDF generation across parallel workers — Word COM and the
    // default soffice profile both break under concurrent use. Keep the lock
    // held for the whole conversion so nothing else touches Word/soffice
    // until we're done.
    let _guard = PDF_LOCK.lock().unwrap_or_else(|e| e.into_inner());

    // Absolute paths for COM.
    let abs_src = std::fs::canonicalize(src).unwrap_or_else(|_| src.to_path_buf());

    // 1) Word COM
    if try_word_pdf(&abs_src, dst).is_ok() && dst.exists() {
        return Ok(());
    }

    // 2) LibreOffice
    if let Some(soffice) = locate_soffice() {
        let out_dir = dst.parent().unwrap_or_else(|| Path::new("."));
        let status = std::process::Command::new(&soffice)
            .args([
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                &out_dir.to_string_lossy(),
                &abs_src.to_string_lossy(),
            ])
            .stderr(std::process::Stdio::null())
            .stdout(std::process::Stdio::null())
            .status()?;
        if status.success() {
            let auto = out_dir.join(format!(
                "{}.pdf",
                src.file_stem().and_then(|s| s.to_str()).unwrap_or("out")
            ));
            if auto != *dst && auto.exists() {
                std::fs::rename(&auto, dst)?;
            }
            if dst.exists() {
                return Ok(());
            }
        }
    }

    anyhow::bail!(
        "PDF 변환 실패. Microsoft Word 또는 LibreOffice가 설치되어 있어야 합니다. \
         (Foxit PDF Reader만으로는 DOCX→PDF 변환이 불가능합니다.)"
    )
}

/// Create a temporary .vbs file, run it with cscript.exe, return stdout.
/// We use VBScript instead of PowerShell because many corporate environments
/// block PowerShell execution by GPO but leave cscript (Windows Script Host)
/// available. cscript has been a Windows built-in since NT4.
fn run_vbs(vbs_source: &str, args: &[&str]) -> anyhow::Result<std::process::Output> {
    let tmp_dir = std::env::temp_dir();
    // Unique filename so parallel runs don't clobber each other even though
    // PDF_LOCK already serializes Word COM access.
    let nonce = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_nanos())
        .unwrap_or(0);
    let vbs_path = tmp_dir.join(format!("compare_{}.vbs", nonce));
    std::fs::write(&vbs_path, vbs_source).context("write vbs script")?;

    let mut cmd = std::process::Command::new("cscript.exe");
    cmd.arg("//Nologo").arg("//B").arg(&vbs_path);
    for a in args { cmd.arg(a); }
    let out = cmd.output();
    let _ = std::fs::remove_file(&vbs_path);
    Ok(out?)
}

fn try_word_pdf(src: &Path, dst: &Path) -> anyhow::Result<()> {
    // VBScript uses `""` for an embedded quote inside a string. Paths are
    // passed as process arguments (WScript.Arguments) to avoid any escaping
    // in the script body itself.
    const VBS: &str = r#"
On Error Resume Next
Set args = WScript.Arguments
Set w = CreateObject("Word.Application")
If Err.Number <> 0 Then WScript.Quit 2
w.Visible = False
w.DisplayAlerts = 0
Set doc = w.Documents.Open(args(0), False, True)
If Err.Number <> 0 Then w.Quit : WScript.Quit 3
doc.SaveAs2 args(1), 17
If Err.Number <> 0 Then doc.Close False : w.Quit : WScript.Quit 4
doc.Close False
w.Quit
WScript.Quit 0
"#;
    let out = run_vbs(VBS, &[&src.to_string_lossy(), &dst.to_string_lossy()])?;
    if out.status.success() { Ok(()) } else { anyhow::bail!("Word COM via cscript failed") }
}

fn locate_soffice() -> Option<PathBuf> {
    for p in [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ] {
        let pb = PathBuf::from(p);
        if pb.exists() {
            return Some(pb);
        }
    }
    None
}

#[tauri::command]
fn open_path(path: String) -> Result<(), String> {
    std::process::Command::new("cmd")
        .args(["/C", "start", "", &path])
        .spawn()
        .map_err(|e| e.to_string())?;
    Ok(())
}

#[tauri::command]
fn reveal_in_folder(path: String) -> Result<(), String> {
    // If `path` is a file, open its parent folder with the file selected;
    // if it's a directory, open the directory.
    let p = Path::new(&path);
    if p.is_file() {
        std::process::Command::new("explorer")
            .args(["/select,", &path])
            .spawn()
            .map_err(|e| e.to_string())?;
    } else {
        std::process::Command::new("explorer")
            .arg(&path)
            .spawn()
            .map_err(|e| e.to_string())?;
    }
    Ok(())
}

#[tauri::command]
fn list_docx_in_dir(dir: String) -> Result<Vec<String>, String> {
    let d = PathBuf::from(&dir);
    let mut out = Vec::new();
    for entry in std::fs::read_dir(&d).map_err(|e| e.to_string())? {
        let e = entry.map_err(|e| e.to_string())?;
        let p = e.path();
        if let Some(ext) = p.extension().and_then(|s| s.to_str()) {
            let low = ext.to_ascii_lowercase();
            if low == "docx" || low == "doc" {
                out.push(p.to_string_lossy().into_owned());
            }
        }
    }
    out.sort();
    Ok(out)
}

fn main() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .invoke_handler(tauri::generate_handler![
            detect_pairs,
            run_batch,
            open_path,
            reveal_in_folder,
            list_docx_in_dir,
        ])
        .run(tauri::generate_context!())
        .expect("run tauri app");
}

struct TempDocx(PathBuf);

impl Drop for TempDocx {
    fn drop(&mut self) {
        let _ = std::fs::remove_file(&self.0);
    }
}

fn ensure_docx(path: &Path) -> anyhow::Result<(PathBuf, Option<TempDocx>)> {
    let ext = path
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase());

    match ext.as_deref() {
        Some("docx") => Ok((path.to_path_buf(), None)),
        Some("doc") => {
            let nonce = std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map(|d| d.as_nanos())
                .unwrap_or(0);
            let tmp_docx = std::env::temp_dir().join(format!("gencompare_doc_{}.docx", nonce));

            convert_doc_to_docx_word_com(path, &tmp_docx)
                .with_context(|| format!("DOC → DOCX 변환 실패: {}", path.display()))?;

            Ok((tmp_docx.clone(), Some(TempDocx(tmp_docx))))
        }
        Some(e) => anyhow::bail!(
            "'{}': 지원하지 않는 확장자 '.{}'. DOCX 또는 DOC 파일을 입력하세요.",
            path.display(),
            e
        ),
        None => anyhow::bail!(
            "'{}': 파일 확장자가 없습니다. DOCX 또는 DOC 파일을 입력하세요.",
            path.display()
        ),
    }
}

fn convert_doc_to_docx_word_com(src: &Path, dst: &Path) -> anyhow::Result<()> {
    let _guard = PDF_LOCK.lock().unwrap_or_else(|e| e.into_inner());

    let abs_src = std::fs::canonicalize(src).unwrap_or_else(|_| src.to_path_buf());
    if let Some(parent) = dst.parent() {
        std::fs::create_dir_all(parent)?;
    }

    const VBS: &str = r#"
On Error Resume Next
Set args = WScript.Arguments
Set w = CreateObject("Word.Application")
If Err.Number <> 0 Then WScript.Quit 2
w.Visible = False
w.DisplayAlerts = 0
Set doc = w.Documents.Open(args(0), False, True)
If Err.Number <> 0 Then w.Quit : WScript.Quit 3
doc.SaveAs2 args(1), 16
If Err.Number <> 0 Then doc.Close False : w.Quit : WScript.Quit 4
doc.Close False
w.Quit
WScript.Quit 0
"#;

    let out = run_vbs(VBS, &[&abs_src.to_string_lossy(), &dst.to_string_lossy()])?;
    if out.status.success() && dst.exists() {
        Ok(())
    } else {
        anyhow::bail!("DOC → DOCX 변환 실패. Microsoft Word가 설치되어 있고 자동화 실행이 허용되어야 합니다.")
    }
}

fn output_base_for_pair(old: &Path, new: &Path) -> String {
    format!(
        "{}_{}",
        sanitize_filename(&file_stem_or_name(old)),
        sanitize_filename(&file_stem_or_name(new))
    )
}

fn file_stem_or_name(path: &Path) -> String {
    path.file_stem()
        .or_else(|| path.file_name())
        .map(|s| s.to_string_lossy().into_owned())
        .unwrap_or_else(|| "document".to_string())
}

fn sanitize_filename(s: &str) -> String {
    let cleaned: String = s
        .chars()
        .map(|c| {
            if c.is_alphanumeric() || c == '_' || c == '-' {
                c
            } else {
                '_'
            }
        })
        .collect();

    let trimmed = cleaned.trim_matches('_');
    if trimmed.is_empty() {
        "document".to_string()
    } else {
        trimmed.to_string()
    }
}
