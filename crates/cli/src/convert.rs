//! DOC → DOCX auto-conversion.
//!
//! Strategy: try Word COM first (fastest, highest fidelity), fall back to
//! LibreOffice headless if Word is not installed or failed. This way the
//! tool "just works" in the usual corporate environment (Word installed) and
//! also in any environment where LibreOffice is present.

use anyhow::{bail, Context, Result};
use std::path::{Path, PathBuf};
use std::process::Command;

pub fn convert_doc_to_docx(input: &Path) -> Result<PathBuf> {
    let abs_input = std::fs::canonicalize(input)
        .with_context(|| format!("resolve path {}", input.display()))?;
    let stem = abs_input
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("converted");
    let nonce = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_nanos())
        .unwrap_or(0);
    let out = std::env::temp_dir().join(format!("compare_{}_{}.docx", stem, nonce));

    // 1) Try Word COM via PowerShell
    match try_word(&abs_input, &out) {
        Ok(()) if out.exists() => return Ok(out),
        _ => {}
    }

    // 2) Fall back to LibreOffice headless
    match try_libreoffice(&abs_input, &out) {
        Ok(()) if out.exists() => return Ok(out),
        _ => {}
    }

    bail!(
        "'{}': DOC → DOCX 변환 실패. Microsoft Word 또는 LibreOffice 중 하나가 설치되어 있어야 합니다.",
        input.display()
    )
}

fn try_word(input: &Path, output: &Path) -> Result<()> {
    let script = format!(
        r#"
$ErrorActionPreference = 'Stop'
try {{
  $w = New-Object -ComObject Word.Application
  $w.Visible = $false
  $w.DisplayAlerts = 0
  $doc = $w.Documents.Open('{input}', $false, $true)
  $doc.SaveAs2('{output}', 16)
  $doc.Close($false)
  $w.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($w) | Out-Null
}} catch {{
  exit 1
}}
"#,
        input = ps_escape(&input.to_string_lossy()),
        output = ps_escape(&output.to_string_lossy()),
    );
    let status = Command::new("powershell")
        .args(["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", &script])
        .stderr(std::process::Stdio::null())
        .status()
        .context("spawn powershell")?;
    if status.success() { Ok(()) } else { bail!("Word COM path failed") }
}

fn try_libreoffice(input: &Path, output: &Path) -> Result<()> {
    let soffice = locate_soffice().context("LibreOffice not found")?;
    let out_dir = output.parent().unwrap_or_else(|| Path::new("."));
    let status = Command::new(&soffice)
        .args([
            "--headless",
            "--convert-to",
            "docx",
            "--outdir",
            &out_dir.to_string_lossy(),
            &input.to_string_lossy(),
        ])
        .stderr(std::process::Stdio::null())
        .stdout(std::process::Stdio::null())
        .status()
        .context("spawn soffice")?;
    if !status.success() { bail!("LibreOffice conversion failed"); }
    // LibreOffice writes <stem>.docx to out_dir. Rename to our expected path.
    let lo_out = out_dir.join(format!(
        "{}.docx",
        input.file_stem().and_then(|s| s.to_str()).unwrap_or("out")
    ));
    if lo_out == *output {
        return Ok(());
    }
    std::fs::rename(&lo_out, output).with_context(|| format!("rename {} -> {}", lo_out.display(), output.display()))?;
    Ok(())
}

fn locate_soffice() -> Option<PathBuf> {
    let candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ];
    for p in candidates {
        let path = PathBuf::from(p);
        if path.exists() { return Some(path); }
    }
    if Command::new("soffice").arg("--version").output().is_ok() {
        return Some(PathBuf::from("soffice"));
    }
    None
}

fn ps_escape(s: &str) -> String {
    s.replace('\'', "''")
}

pub struct TempFile(pub PathBuf);

impl Drop for TempFile {
    fn drop(&mut self) {
        let _ = std::fs::remove_file(&self.0);
    }
}
