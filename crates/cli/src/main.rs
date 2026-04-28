mod convert;

use anyhow::{Context, Result};
use clap::{Parser, Subcommand};
use compare_core::{detect_pairs, diff_blocks, diff_paragraphs, stats_of_blocks, BlockOp, Stats};
use compare_docx::{read_document, write_redline, RedlineOptions, RedlineStyle};
use rayon::prelude::*;
use serde::{Deserialize, Serialize};
use std::path::{Path, PathBuf};
use std::time::Instant;

use convert::{convert_doc_to_docx, TempFile};

/// DOCX document comparator. Supports single diff and parallel batch mode.
#[derive(Parser)]
#[command(version, about)]
struct Cli {
    #[command(subcommand)]
    cmd: Cmd,
}

#[derive(Parser, Debug, Clone)]
struct AuthorOpts {
    /// Author name written to track-change marks (default: $USERNAME)
    #[arg(long, global = false)]
    author: Option<String>,
}

#[derive(clap::ValueEnum, Clone, Copy, Debug)]
enum StyleArg {
    /// Visible colors (red/blue/green) via direct formatting — always displays correctly
    Color,
    /// Real track-change marks (accept/reject in Word) — color may be overridden by reader
    Track,
}

impl From<StyleArg> for RedlineStyle {
    fn from(v: StyleArg) -> Self {
        match v {
            StyleArg::Color => RedlineStyle::Color,
            StyleArg::Track => RedlineStyle::TrackChange,
        }
    }
}

#[derive(Subcommand)]
enum Cmd {
    /// Compare one pair of DOCX files.
    Diff {
        /// Original (old) DOCX
        old: PathBuf,
        /// Modified (new) DOCX
        new: PathBuf,
        /// Output redline DOCX
        #[arg(short, long, default_value = "redline.docx")]
        out: PathBuf,
        /// Output style: `color` (visible red/blue/green) or `track` (Word track changes)
        #[arg(long, value_enum, default_value_t = StyleArg::Color)]
        style: StyleArg,
        /// Author name for track-change marks (default: $USERNAME)
        #[arg(long)]
        author: Option<String>,
        /// Emit machine-readable JSON summary to stdout
        #[arg(long)]
        json: bool,
    },
    /// Compare many pairs in parallel from a JSON manifest.
    /// Manifest format: [{"old": "...", "new": "...", "out": "..."}, ...]
    Batch {
        /// JSON manifest listing pairs
        manifest: PathBuf,
        /// Output style
        #[arg(long, value_enum, default_value_t = StyleArg::Color)]
        style: StyleArg,
        /// Author name for track-change marks (default: $USERNAME)
        #[arg(long)]
        author: Option<String>,
        /// Emit machine-readable JSON summary to stdout
        #[arg(long)]
        json: bool,
    },
    /// Auto-detect pairs in a directory and compare in parallel.
    /// Recognizes old/new, 원본/수정, v1/v2, dates, numeric suffixes, etc.
    AutoBatch {
        /// Directory containing .docx files
        dir: PathBuf,
        /// Directory to write redline outputs (defaults to <dir>/redlines)
        #[arg(long)]
        out_dir: Option<PathBuf>,
        /// Only show detected pairs, do not run
        #[arg(long)]
        dry_run: bool,
        /// Output style
        #[arg(long, value_enum, default_value_t = StyleArg::Color)]
        style: StyleArg,
        /// Author name for track-change marks (default: $USERNAME)
        #[arg(long)]
        author: Option<String>,
        /// Emit machine-readable JSON summary to stdout
        #[arg(long)]
        json: bool,
    },
}

#[derive(Deserialize, Serialize, Clone)]
struct Pair {
    old: PathBuf,
    new: PathBuf,
    out: PathBuf,
}

#[derive(Serialize)]
struct DiffReport {
    old: PathBuf,
    new: PathBuf,
    out: PathBuf,
    stats: Stats,
    header_stats: Stats,
    footer_stats: Stats,
    elapsed_ms: u128,
}

#[derive(Serialize)]
struct BatchReport {
    pairs: usize,
    total_elapsed_ms: u128,
    workers: usize,
    results: Vec<PairResult>,
}

#[derive(Serialize)]
struct PairResult {
    index: usize,
    old: PathBuf,
    new: PathBuf,
    out: PathBuf,
    #[serde(skip_serializing_if = "Option::is_none")]
    stats: Option<Stats>,
    #[serde(skip_serializing_if = "Option::is_none")]
    error: Option<String>,
    elapsed_ms: u128,
}

fn main() {
    let cli = Cli::parse();
    let json_mode = matches!(&cli.cmd,
        Cmd::Diff { json, .. } | Cmd::Batch { json, .. } | Cmd::AutoBatch { json, .. } if *json
    );

    let result = match cli.cmd {
        Cmd::Diff { old, new, out, style, author, json } => {
            cmd_diff(&old, &new, &out, style.into(), author, json)
        }
        Cmd::Batch { manifest, style, author, json } => {
            cmd_batch(&manifest, style.into(), author, json)
        }
        Cmd::AutoBatch { dir, out_dir, dry_run, style, author, json } => {
            cmd_auto_batch(&dir, out_dir.as_deref(), dry_run, style.into(), author, json)
        }
    };

    if let Err(e) = result {
        if json_mode {
            // Machine-readable error for callers like Genie.
            let err = serde_json::json!({
                "error": true,
                "message": format!("{e}"),
                "detail": format!("{e:#}"),
            });
            println!("{}", serde_json::to_string_pretty(&err).unwrap_or_else(|_| "{}".into()));
        } else {
            eprintln!("error: {e:#}");
        }
        std::process::exit(1);
    }
}

fn make_opts(style: RedlineStyle, author: Option<String>) -> RedlineOptions {
    let mut opts = RedlineOptions::default();
    opts.style = style;
    if let Some(a) = author {
        opts.author = a;
    }
    opts
}

fn none_if_empty<T: Clone>(v: &[T]) -> Option<Vec<T>> {
    if v.is_empty() { None } else { Some(v.to_vec()) }
}

struct OneResult {
    body_stats: Stats,
    header_stats: Stats,
    footer_stats: Stats,
    elapsed_ms: u128,
}

fn run_one(
    old: &Path,
    new: &Path,
    out: &Path,
    style: RedlineStyle,
    author: Option<&str>,
) -> Result<OneResult> {
    let t0 = Instant::now();
    let (old_docx, _old_temp) = ensure_docx(old)?;
    let (new_docx, _new_temp) = ensure_docx(new)?;
    let od = read_document(&old_docx)?;
    let nd = read_document(&new_docx)?;

    let body_ops: Vec<BlockOp> = diff_blocks(&od.body, &nd.body);
    let diff_if_changed = |o: &[String], n: &[String]| -> Vec<_> {
        if o != n { diff_paragraphs(o, n) } else { Vec::new() }
    };
    let header_ops   = diff_if_changed(&od.headers,   &nd.headers);
    let footer_ops   = diff_if_changed(&od.footers,   &nd.footers);
    let comment_ops  = diff_if_changed(&od.comments,  &nd.comments);
    let footnote_ops = diff_if_changed(&od.footnotes, &nd.footnotes);
    let endnote_ops  = diff_if_changed(&od.endnotes,  &nd.endnotes);

    let mut opts = make_opts(style, author.map(String::from));
    opts.header_changes   = none_if_empty(&header_ops);
    opts.footer_changes   = none_if_empty(&footer_ops);
    opts.comment_changes  = none_if_empty(&comment_ops);
    opts.footnote_changes = none_if_empty(&footnote_ops);
    opts.endnote_changes  = none_if_empty(&endnote_ops);
    opts.original_name = old.file_name().and_then(|s| s.to_str()).map(String::from);
    opts.modified_name = new.file_name().and_then(|s| s.to_str()).map(String::from);
    // Use `new`'s style/theme/numbering/font definitions so the redline
    // inherits the authoritative formatting of the current version.
    opts.source_parts = Some(nd.raw_parts.clone());

    write_redline(out, &body_ops, &opts)?;

    Ok(OneResult {
        body_stats: stats_of_blocks(&body_ops),
        header_stats: compare_core::stats_of(&header_ops),
        footer_stats: compare_core::stats_of(&footer_ops),
        elapsed_ms: t0.elapsed().as_millis(),
    })
}

fn cmd_diff(
    old: &Path,
    new: &Path,
    out: &Path,
    style: RedlineStyle,
    author: Option<String>,
    json: bool,
) -> Result<()> {
    let r = run_one(old, new, out, style, author.as_deref())?;

    if json {
        let report = DiffReport {
            old: old.to_path_buf(),
            new: new.to_path_buf(),
            out: out.to_path_buf(),
            stats: r.body_stats,
            header_stats: r.header_stats,
            footer_stats: r.footer_stats,
            elapsed_ms: r.elapsed_ms,
        };
        println!("{}", serde_json::to_string_pretty(&report)?);
    } else {
        println!("redline written: {}", out.display());
        println!(
            "body: equal={} +{} -{} ~{}  words +{}/-{}",
            r.body_stats.paragraphs_equal,
            r.body_stats.paragraphs_inserted,
            r.body_stats.paragraphs_deleted,
            r.body_stats.paragraphs_modified,
            r.body_stats.words_inserted,
            r.body_stats.words_deleted,
        );
        println!(
            "header words: +{}/-{}    footer words: +{}/-{}",
            r.header_stats.words_inserted,
            r.header_stats.words_deleted,
            r.footer_stats.words_inserted,
            r.footer_stats.words_deleted,
        );
        println!("elapsed: {} ms", r.elapsed_ms);
    }
    Ok(())
}

fn cmd_batch(
    manifest: &Path,
    style: RedlineStyle,
    author: Option<String>,
    json: bool,
) -> Result<()> {
    let raw = std::fs::read_to_string(manifest)
        .with_context(|| format!("read manifest {}", manifest.display()))?;
    let pairs: Vec<Pair> = serde_json::from_str(&raw).context("parse manifest JSON")?;
    run_pairs_parallel(&pairs, style, author.as_deref(), json, 0)
}

fn cmd_auto_batch(
    dir: &Path,
    out_dir: Option<&Path>,
    dry_run: bool,
    style: RedlineStyle,
    author: Option<String>,
    json: bool,
) -> Result<()> {
    let files = collect_docx(dir)?;
    let detection = detect_pairs(&files);

    let out_root = out_dir.map(|p| p.to_path_buf()).unwrap_or_else(|| dir.join("redlines"));

let pairs: Vec<Pair> = detection.pairs.iter().map(|m| {
    let old_name = m.old
        .file_stem()
        .unwrap_or_default()
        .to_string_lossy();

    let new_name = m.new
        .file_stem()
        .unwrap_or_default()
        .to_string_lossy();

    let out = out_root.join(format!(
        "redline_{}_{}.docx",
        sanitize(&old_name),
        sanitize(&new_name)
    ));

    Pair {
        old: m.old.clone(),
        new: m.new.clone(),
        out,
    }
}).collect();

    if dry_run {
        if json {
            #[derive(Serialize)]
            struct DryOut<'a> { pairs: &'a [Pair], unpaired: &'a [PathBuf] }
            println!("{}", serde_json::to_string_pretty(&DryOut {
                pairs: &pairs, unpaired: &detection.unpaired
            })?);
        } else {
            println!("detected {} pair(s), {} unpaired:", pairs.len(), detection.unpaired.len());
            for (i, p) in pairs.iter().enumerate() {
                println!("  [{}] {} -> {}", i,
                    p.old.file_name().unwrap_or_default().to_string_lossy(),
                    p.new.file_name().unwrap_or_default().to_string_lossy());
            }
            for u in &detection.unpaired {
                println!("  (unpaired) {}", u.file_name().unwrap_or_default().to_string_lossy());
            }
        }
        return Ok(());
    }

    if !out_root.exists() {
        std::fs::create_dir_all(&out_root)?;
    }

    run_pairs_parallel(&pairs, style, author.as_deref(), json, detection.unpaired.len())
}

fn run_pairs_parallel(
    pairs: &[Pair],
    style: RedlineStyle,
    author: Option<&str>,
    json: bool,
    unpaired_count: usize,
) -> Result<()> {
    let t0 = Instant::now();

    let results: Vec<PairResult> = pairs
        .par_iter()
        .enumerate()
        .map(|(i, p)| {
            let t1 = Instant::now();
            match run_one(&p.old, &p.new, &p.out, style, author) {
                Ok(r) => PairResult {
                    index: i,
                    old: p.old.clone(),
                    new: p.new.clone(),
                    out: p.out.clone(),
                    stats: Some(r.body_stats),
                    error: None,
                    elapsed_ms: t1.elapsed().as_millis(),
                },
                Err(e) => PairResult {
                    index: i,
                    old: p.old.clone(),
                    new: p.new.clone(),
                    out: p.out.clone(),
                    stats: None,
                    error: Some(format!("{:#}", e)),
                    elapsed_ms: t1.elapsed().as_millis(),
                },
            }
        })
        .collect();

    let total_elapsed_ms = t0.elapsed().as_millis();
    let workers = rayon::current_num_threads();

    if json {
        let report = BatchReport { pairs: pairs.len(), total_elapsed_ms, workers, results };
        println!("{}", serde_json::to_string_pretty(&report)?);
    } else {
        let ok = results.iter().filter(|r| r.error.is_none()).count();
        let err = results.len() - ok;
        let tail = if unpaired_count > 0 { format!(", unpaired={}", unpaired_count) } else { String::new() };
        println!(
            "batch done: {} pairs in {} ms on {} workers  (ok={}, err={}{})",
            results.len(), total_elapsed_ms, workers, ok, err, tail
        );
        for r in &results {
            match (&r.stats, &r.error) {
                (Some(s), _) => println!(
                    "  [{}] {}  ->  Δp={}/+{}/-{}/~{}  Δw=+{}/-{}  ({} ms)",
                    r.index,
                    r.out.display(),
                    s.paragraphs_equal,
                    s.paragraphs_inserted,
                    s.paragraphs_deleted,
                    s.paragraphs_modified,
                    s.words_inserted,
                    s.words_deleted,
                    r.elapsed_ms,
                ),
                (_, Some(e)) => println!("  [{}] ERROR: {}", r.index, e),
                _ => {}
            }
        }
    }
    Ok(())
}

/// Normalize an input file to DOCX. Accepts .docx directly; auto-converts
/// .doc via Word COM. Returns the DOCX path plus an optional TempFile guard
/// that deletes the temporary conversion when dropped.
fn ensure_docx(path: &Path) -> Result<(PathBuf, Option<TempFile>)> {
    let ext = path
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase());
    match ext.as_deref() {
        Some("docx") => Ok((path.to_path_buf(), None)),
        Some("doc") => {
            let converted = convert_doc_to_docx(path)
                .with_context(|| format!("convert {}", path.display()))?;
            let guard = TempFile(converted.clone());
            Ok((converted, Some(guard)))
        }
        Some("hwp") | Some("hwpx") => anyhow::bail!(
            "'{}': 한글(HWP) 파일은 직접 지원하지 않습니다. \
             DOCX로 변환해서 사용해주세요.",
            path.display()
        ),
        Some("pdf") => anyhow::bail!(
            "'{}': PDF는 직접 지원하지 않습니다. \
             DOCX로 변환해서 사용해주세요.",
            path.display()
        ),
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

fn collect_docx(dir: &Path) -> Result<Vec<PathBuf>> {
    let mut out = Vec::new();
    for entry in std::fs::read_dir(dir)
        .with_context(|| format!("read dir {}", dir.display()))?
    {
        let e = entry?;
        let p = e.path();
        if p.extension().and_then(|s| s.to_str()).map_or(false, |s| s.eq_ignore_ascii_case("docx")) {
            out.push(p);
        }
    }
    out.sort();
    Ok(out)
}

fn sanitize(s: &str) -> String {
    s.chars()
        .map(|c| if c.is_alphanumeric() || c == '_' || c == '-' {
            c
        } else {
            '_'
        })
        .collect::<String>()
        .trim_matches('_')
        .to_string()
}
