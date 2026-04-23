// Compare GUI — all interactive behaviour.
//
// Global __TAURI__ is exposed via tauri.conf.json `withGlobalTauri: true`.
// We defensively check in case something is loaded out of order.

if (!window.__TAURI__) {
  document.body.innerHTML = '<div style="padding:40px;font-family:sans-serif;color:#c62828">' +
    '<h2>Tauri bridge가 로드되지 않았습니다.</h2>' +
    '<p>tauri.conf.json의 <code>withGlobalTauri: true</code> 설정 + 재빌드가 필요합니다.</p></div>';
  throw new Error('TAURI global missing');
}

const { invoke } = window.__TAURI__.core;
const { open } = window.__TAURI__.dialog;

// ---------- state ----------
let currentPairs = [];
let currentUnpaired = [];
let currentOutDir = null;
// Quick 2-zone input — holds the single pair the user is building manually.
// Filled by dropping one file (fills the next empty slot) or when both slots
// have a file — the user clicks "짝 추가" to commit it into currentPairs.
const quickFiles = { old: null, new: null };

// Persisted settings (localStorage).
const DEFAULTS = {
  author: '',
  outputs: { word: true, track: false, pdf: false, cpo: false },
};
const settings = loadSettings();

// ---------- helpers ----------
function $(id) { return document.getElementById(id); }
function show(id) { $(id).classList.remove('hidden'); }
function hide(id) { $(id).classList.add('hidden'); }

function basename(p) { return p.split(/[\\/]/).pop(); }
function dirname(p) {
  const idx = Math.max(p.lastIndexOf('/'), p.lastIndexOf('\\'));
  return idx >= 0 ? p.slice(0, idx) : '.';
}

function loadSettings() {
  try {
    const raw = localStorage.getItem('compare.settings');
    if (raw) return { ...DEFAULTS, ...JSON.parse(raw) };
  } catch {}
  return { ...DEFAULTS };
}
function saveSettings() {
  localStorage.setItem('compare.settings', JSON.stringify(settings));
}

function toast(message, kind = 'info') {
  const el = document.createElement('div');
  el.className = `toast ${kind}`;
  el.textContent = message;
  $('toasts').appendChild(el);
  setTimeout(() => {
    el.style.opacity = '0';
    el.style.transform = 'translateX(100%)';
    el.style.transition = 'all 200ms';
    setTimeout(() => el.remove(), 220);
  }, kind === 'error' ? 6000 : 3200);
}

function setProgress(title, detail, visible = true) {
  if (visible) {
    $('progress-title').textContent = title;
    $('progress-detail').textContent = detail || '';
    show('progress-overlay');
  } else {
    hide('progress-overlay');
  }
}

// ---------- file handling ----------
async function handleFiles(paths) {
  if (!paths || !paths.length) return;

  setProgress('파일 분석 중…', `${paths.length}개 파일`);
  await new Promise((r) => setTimeout(r, 30));
  try {
    const result = await invoke('detect_pairs', { paths });
    currentPairs = result.pairs;
    currentUnpaired = result.unpaired;
    currentOutDir = dirname(paths[0]) + '\\redlines';
    renderPairs();
    updateOutDirDisplay();
    show('pairs-section');
    if (currentPairs.length > 0) {
      show('options-section');
    }
    if (currentPairs.length === 0 && currentUnpaired.length > 0) {
      toast('감지된 짝이 없습니다. 파일 이름에 _old/_new 등의 마커를 추가하세요.', 'warning');
    }
  } catch (err) {
    toast('파일 분석 실패: ' + err, 'error');
  } finally {
    setProgress('', '', false);
  }
}

function renderPairs() {
  const tbody = document.querySelector('#pairs-table tbody');
  tbody.innerHTML = '';
  for (let i = 0; i < currentPairs.length; i++) {
    const p = currentPairs[i];
    const row = document.createElement('tr');
    row.dataset.idx = i;
    const oldCell = p.old
      ? `<td class="filename drop-cell" data-idx="${i}" data-side="old" draggable="true" title="${p.old}">${basename(p.old)}</td>`
      : `<td class="filename drop-cell empty-slot" data-idx="${i}" data-side="old">＋ 원본을 여기에 끌어놓기</td>`;
    const newCell = p.new
      ? `<td class="filename drop-cell" data-idx="${i}" data-side="new" draggable="true" title="${p.new}">${basename(p.new)}</td>`
      : `<td class="filename drop-cell empty-slot" data-idx="${i}" data-side="new">＋ 수정본을 여기에 끌어놓기</td>`;
    row.innerHTML = `
      ${oldCell}
      <td class="arrow-swap"><button class="swap-btn" data-idx="${i}" title="원본·수정본 교체">⇄</button></td>
      ${newCell}
      <td class="reason-col">${p.reason || ''}</td>
      <td><button class="remove-pair-btn" data-idx="${i}" title="이 짝 제외">✕</button></td>
    `;
    tbody.appendChild(row);
  }
  $('pairs-summary').textContent =
    `${currentPairs.length}쌍` +
    (currentUnpaired.length ? ` · ${currentUnpaired.length}개 미매칭` : '');

  const up = $('unpaired');
  const ul = up.querySelector('ul');
  ul.innerHTML = '';
  if (currentUnpaired.length) {
    up.classList.remove('hidden');
    for (const f of currentUnpaired) {
      const li = document.createElement('li');
      li.className = 'unpaired-item';
      li.textContent = '📄 ' + basename(f);
      li.title = f;
      li.draggable = true;
      li.dataset.path = f;
      li.addEventListener('dragstart', (e) => {
        e.dataTransfer.setData('application/compare-unpaired', f);
        e.dataTransfer.effectAllowed = 'copyMove';
        li.classList.add('dragging');
      });
      li.addEventListener('dragend', () => li.classList.remove('dragging'));
      ul.appendChild(li);
    }
  } else {
    up.classList.add('hidden');
  }

  // Remove-pair handlers — push non-null files back to the unpaired list.
  tbody.querySelectorAll('.remove-pair-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      const idx = parseInt(btn.dataset.idx, 10);
      const removed = currentPairs.splice(idx, 1)[0];
      for (const f of [removed.old, removed.new]) {
        if (f && !currentUnpaired.includes(f)) currentUnpaired.push(f);
      }
      renderPairs();
    });
  });

  // Swap button: flip old/new within the same pair.
  tbody.querySelectorAll('.swap-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      const idx = parseInt(btn.dataset.idx, 10);
      const p = currentPairs[idx];
      [p.old, p.new] = [p.new, p.old];
      renderPairs();
    });
  });

  // Cell-level drag & drop:
  //   - drag a filename cell onto another cell to move it there
  //   - drag an unpaired list item onto a cell to assign it
  //   - hold Ctrl while dragging to copy (source keeps its file)
  tbody.querySelectorAll('.drop-cell').forEach((cell) => {
    // Only filled cells are source-draggable (empty placeholders are not).
    if (!cell.classList.contains('empty-slot')) {
      cell.addEventListener('dragstart', (e) => {
        const idx = cell.dataset.idx;
        const side = cell.dataset.side;
        e.dataTransfer.setData('application/compare-cell', JSON.stringify({ idx, side }));
        e.dataTransfer.effectAllowed = 'copyMove';
        cell.classList.add('dragging');
      });
      cell.addEventListener('dragend', () => cell.classList.remove('dragging'));
    }
    cell.addEventListener('dragover', (e) => {
      const types = e.dataTransfer.types;
      if (!types.includes('application/compare-cell') &&
          !types.includes('application/compare-unpaired')) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = e.ctrlKey ? 'copy' : 'move';
      cell.classList.add('drop-target');
    });
    cell.addEventListener('dragleave', () => cell.classList.remove('drop-target'));
    cell.addEventListener('drop', (e) => {
      cell.classList.remove('drop-target');
      const dstIdx = parseInt(cell.dataset.idx, 10);
      const dstSide = cell.dataset.side;
      const isCopy = e.ctrlKey;

      // Case 1: dropping from unpaired list.
      const unpairedPath = e.dataTransfer.getData('application/compare-unpaired');
      if (unpairedPath) {
        e.preventDefault();
        currentPairs[dstIdx][dstSide] = unpairedPath;
        // Auto-derive base name from the old side if not yet set.
        if (!currentPairs[dstIdx].base) {
          const src = currentPairs[dstIdx].old || currentPairs[dstIdx].new;
          if (src) currentPairs[dstIdx].base = basename(src).replace(/\.[^.]+$/, '');
        }
        if (!isCopy) {
          // Move: remove from unpaired list.
          currentUnpaired = currentUnpaired.filter((f) => f !== unpairedPath);
        }
        renderPairs();
        if (currentPairs.some((p) => p.old && p.new)) show('options-section');
        return;
      }

      // Case 2: dropping from another cell.
      const raw = e.dataTransfer.getData('application/compare-cell');
      if (!raw) return;
      e.preventDefault();
      const { idx: srcIdxStr, side: srcSide } = JSON.parse(raw);
      const srcIdx = parseInt(srcIdxStr, 10);
      if (srcIdx === dstIdx && srcSide === dstSide) return;
      const srcFile = currentPairs[srcIdx][srcSide];
      const dstFile = currentPairs[dstIdx][dstSide];
      if (isCopy) {
        currentPairs[dstIdx][dstSide] = srcFile;
      } else {
        currentPairs[dstIdx][dstSide] = srcFile;
        currentPairs[srcIdx][srcSide] = dstFile;
      }
      renderPairs();
    });
  });
}

function updateOutDirDisplay() {
  $('out-dir-display').textContent = currentOutDir || '(자동)';
}

// ---------- quick 2-zone input ----------
function renderQuick() {
  for (const side of ['old', 'new']) {
    const zone = $(`quick-${side}`);
    const fileEl = zone.querySelector('.quick-file');
    const path = quickFiles[side];
    if (path) {
      fileEl.textContent = basename(path);
      fileEl.title = path;
      fileEl.classList.remove('muted');
      zone.classList.add('filled');
    } else {
      fileEl.textContent = side === 'old' ? '여기에 원본 끌어놓기' : '여기에 수정본 끌어놓기';
      fileEl.title = '';
      fileEl.classList.add('muted');
      zone.classList.remove('filled');
    }
  }
  $('quick-add-btn').disabled = !(quickFiles.old && quickFiles.new);
}

function fillNextQuickSlot(path) {
  if (!quickFiles.old) quickFiles.old = path;
  else if (!quickFiles.new) quickFiles.new = path;
  else quickFiles.new = path; // overwrite the last one
  renderQuick();
  // Auto-commit once both slots are filled — user shouldn't need to click
  // "짝 추가" for the common single-pair case. Options/Run appear immediately.
  if (quickFiles.old && quickFiles.new) {
    commitQuickPair();
  }
}

function commitQuickPair() {
  if (!quickFiles.old || !quickFiles.new) return;
  if (quickFiles.old === quickFiles.new) {
    toast('원본과 수정본은 서로 달라야 합니다.', 'warning');
    return;
  }
  const base = basename(quickFiles.old).replace(/\.[^.]+$/, '');
  currentPairs.push({ old: quickFiles.old, new: quickFiles.new, base, reason: '직접 지정' });
  if (!currentOutDir) currentOutDir = dirname(quickFiles.old) + '\\redlines';
  quickFiles.old = null;
  quickFiles.new = null;
  renderQuick();
  renderPairs();
  updateOutDirDisplay();
  show('pairs-section');
  show('options-section');
}

// Click a zone to browse for a file manually.
for (const side of ['old', 'new']) {
  $(`quick-${side}`).addEventListener('click', async () => {
    const picked = await open({
      multiple: false,
      filters: [{ name: 'Word 문서', extensions: ['docx', 'doc'] }],
    });
    if (!picked) return;
    quickFiles[side] = picked;
    renderQuick();
    if (quickFiles.old && quickFiles.new) commitQuickPair();
  });
}
$('quick-add-btn').addEventListener('click', commitQuickPair);

// ---------- run compare ----------
async function runAll() {
  const outputs = {
    word: $('out-word').checked,
    track_change: $('out-track').checked,
    pdf: $('out-pdf').checked,
    cpo: $('out-cpo').checked,
  };
  if (!outputs.word && !outputs.track_change && !outputs.pdf && !outputs.cpo) {
    toast('저장 형식을 하나 이상 선택하세요.', 'warning');
    return;
  }
  if (!currentPairs.length) {
    toast('비교할 짝이 없습니다.', 'warning');
    return;
  }
  const filledPairs = currentPairs.filter((p) => p.old && p.new);
  if (!filledPairs.length) {
    toast('원본·수정본이 모두 채워진 짝이 없습니다.', 'warning');
    return;
  }
  if (filledPairs.length < currentPairs.length) {
    toast(`${currentPairs.length - filledPairs.length}개의 빈 짝은 건너뜁니다.`, 'warning');
  }

  const author = settings.author.trim() || null;
  const runPairs = filledPairs.map((p, i) => ({
    old: p.old,
    new: p.new,
    out_base: `${p.base || 'pair'}_${i + 1}`,
    out_dir: currentOutDir,
  }));

  const btn = $('run-btn');
  btn.disabled = true;
  btn.querySelector('.btn-text').textContent = '처리 중…';
  btn.querySelector('.btn-spinner').classList.remove('hidden');
  setProgress('비교 진행 중…', `${runPairs.length}쌍 · 병렬 처리`);

  try {
    const report = await invoke('run_batch', { pairs: runPairs, outputs, author });
    renderResults(report);
    show('results-section');
    const okCount = report.results.filter((r) => !r.error).length;
    toast(
      `${okCount}/${report.results.length}쌍 완료 · ${report.total_elapsed_ms}ms`,
      okCount === report.results.length ? 'success' : 'warning'
    );
  } catch (e) {
    toast('오류: ' + e, 'error');
  } finally {
    btn.disabled = false;
    btn.querySelector('.btn-text').textContent = '비교 시작';
    btn.querySelector('.btn-spinner').classList.add('hidden');
    setProgress('', '', false);
  }
}

function renderResults(report) {
  $('results-summary').textContent =
    `${report.results.length}쌍 · ${report.total_elapsed_ms}ms · ${report.workers} workers`;
  const list = $('results-list');
  list.innerHTML = '';
  for (const r of report.results) {
    const card = document.createElement('div');
    card.className = 'result-card' + (r.error ? ' error' : '');

    if (r.error) {
      card.innerHTML = `
        <div class="result-header">
          <div class="result-pair">
            <span class="old" title="${r.old}">${basename(r.old)}</span>
            <span class="sep">→</span>
            <span class="new" title="${r.new}">${basename(r.new)}</span>
          </div>
          <div class="result-time">${r.elapsed_ms}ms</div>
        </div>
        <div class="result-error">${escapeHtml(r.error)}</div>
      `;
    } else {
      const s = r.stats;
      const statParts = [];
      if (s.words_inserted) statParts.push(`<span class="stat ins"><span class="stat-dot"></span>+${s.words_inserted}단어</span>`);
      if (s.words_deleted)  statParts.push(`<span class="stat del"><span class="stat-dot"></span>-${s.words_deleted}단어</span>`);
      if (s.words_moved)    statParts.push(`<span class="stat mov"><span class="stat-dot"></span>↔${s.words_moved}단어</span>`);
      if (s.paragraphs_modified) statParts.push(`<span class="stat">문단 ${s.paragraphs_modified}개 수정</span>`);
      if (s.rows_inserted || s.rows_deleted || s.rows_modified) {
        statParts.push(`<span class="stat">표 행 +${s.rows_inserted}/-${s.rows_deleted}/~${s.rows_modified}</span>`);
      }
      if (!statParts.length) statParts.push('<span class="stat muted">변경 없음</span>');

      const outBtns = r.outputs
        .map((o) =>
          `<button data-path="${o}" data-action="open">📄 ${basename(o)}</button>`
        )
        .join('');
      const folderBtn = r.outputs.length
        ? `<button data-path="${dirname(r.outputs[0])}" data-action="folder">📁 폴더 열기</button>`
        : '';

      card.innerHTML = `
        <div class="result-header">
          <div class="result-pair">
            <span class="old" title="${r.old}">${basename(r.old)}</span>
            <span class="sep">→</span>
            <span class="new" title="${r.new}">${basename(r.new)}</span>
          </div>
          <div class="result-time">${r.elapsed_ms}ms</div>
        </div>
        <div class="result-stats">${statParts.join('')}</div>
        <div class="result-outputs">${outBtns}${folderBtn}</div>
      `;
    }
    list.appendChild(card);
  }

  // Wire up open/folder buttons
  list.querySelectorAll('button[data-path]').forEach((btn) => {
    btn.addEventListener('click', async () => {
      try {
        const action = btn.dataset.action;
        if (action === 'open') {
          await invoke('open_path', { path: btn.dataset.path });
        } else if (action === 'folder') {
          await invoke('reveal_in_folder', { path: btn.dataset.path });
        }
      } catch (e) {
        toast('열기 실패: ' + e, 'error');
      }
    });
  });
}

function escapeHtml(s) {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ---------- drag and drop ----------
//
// In Tauri 2 native drag events fire only when `dragDropEnabled: true` in the
// window config. We use the webview's onDragDropEvent for broadest compat;
// falling back to global event listen if needed. HTML5 drag events are
// suppressed at document level only to prevent the browser navigating away
// when files are dropped outside our control.
const dz = $('quick-pair');

document.addEventListener('dragover', (e) => {
  e.preventDefault();
});
document.addEventListener('drop', (e) => {
  e.preventDefault();
});

(async () => {
  const webviewMod = window.__TAURI__.webview;
  const eventMod = window.__TAURI__.event;

  const onDrop = async (paths) => {
    dz.classList.remove('dragover');
    if (!paths || !paths.length) return;
    const docs = paths.filter((p) => /\.(docx?|DOCX?)$/.test(p));
    if (!docs.length) {
      toast('DOCX/DOC 파일만 지원합니다.', 'warning');
      return;
    }
    // Routing rules:
    //   1 file   → fill the next empty quick-zone slot
    //   2 files  → auto-commit as a pair (old=first, new=second)
    //   3+ files → batch-detection mode (populate the pair table)
    if (docs.length === 1) {
      fillNextQuickSlot(docs[0]);
    } else if (docs.length === 2) {
      quickFiles.old = docs[0];
      quickFiles.new = docs[1];
      renderQuick();
      commitQuickPair();
    } else {
      await handleFiles(docs);
    }
  };

  // Preferred: getCurrentWebview().onDragDropEvent (Tauri 2 idiomatic)
  if (webviewMod && typeof webviewMod.getCurrentWebview === 'function') {
    try {
      const wv = webviewMod.getCurrentWebview();
      await wv.onDragDropEvent((e) => {
        const kind = e.payload?.type;
        if (kind === 'enter' || kind === 'over') {
          dz.classList.add('dragover');
        } else if (kind === 'leave') {
          dz.classList.remove('dragover');
        } else if (kind === 'drop') {
          onDrop(e.payload.paths);
        }
      });
      return;
    } catch (err) {
      console.warn('onDragDropEvent failed, falling back to global listen', err);
    }
  }

  // Fallback: global event listen
  if (eventMod && typeof eventMod.listen === 'function') {
    await eventMod.listen('tauri://drag-enter', () => dz.classList.add('dragover'));
    await eventMod.listen('tauri://drag-over',  () => dz.classList.add('dragover'));
    await eventMod.listen('tauri://drag-leave', () => dz.classList.remove('dragover'));
    await eventMod.listen('tauri://drag-drop',  (e) => onDrop(e.payload?.paths || []));
  }
})();

// ---------- browse buttons ----------
$('browse-folder-btn').addEventListener('click', async () => {
  const selected = await open({ directory: true, multiple: false });
  if (!selected) return;
  try {
    const list = await invoke('list_docx_in_dir', { dir: selected });
    if (!list.length) {
      toast('폴더에 DOCX/DOC 파일이 없습니다.', 'warning');
      return;
    }
    await handleFiles(list);
  } catch (e) {
    toast('폴더 읽기 실패: ' + e, 'error');
  }
});

$('change-dir-btn').addEventListener('click', async () => {
  const selected = await open({ directory: true, multiple: false });
  if (!selected) return;
  currentOutDir = selected;
  updateOutDirDisplay();
});

// "＋ 짝 추가" in the pairs card header — appends an empty row so the user
// can drag-drop an unpaired file into each slot.
$('add-empty-pair-btn').addEventListener('click', () => {
  currentPairs.push({ old: null, new: null, base: '', reason: '직접 지정' });
  show('pairs-section');
  renderPairs();
});

$('clear-btn').addEventListener('click', () => {
  currentPairs = [];
  currentUnpaired = [];
  quickFiles.old = null;
  quickFiles.new = null;
  renderQuick();
  hide('pairs-section');
  hide('options-section');
  hide('results-section');
});

$('open-out-dir-btn').addEventListener('click', async () => {
  if (!currentOutDir) return;
  try {
    await invoke('reveal_in_folder', { path: currentOutDir });
  } catch (e) {
    toast('폴더 열기 실패: ' + e, 'error');
  }
});

$('run-btn').addEventListener('click', runAll);

// ---------- settings modal ----------
function openSettings() {
  $('settings-author').value = settings.author;
  $('settings-out-word').checked = settings.outputs.word;
  $('settings-out-track').checked = settings.outputs.track;
  $('settings-out-pdf').checked = settings.outputs.pdf;
  $('settings-out-cpo').checked = !!settings.outputs.cpo;
  show('settings-modal');
}
function closeSettings() { hide('settings-modal'); }

$('settings-btn').addEventListener('click', openSettings);
document.querySelectorAll('#settings-modal [data-close]').forEach((el) =>
  el.addEventListener('click', closeSettings)
);
$('settings-save').addEventListener('click', () => {
  settings.author = $('settings-author').value;
  settings.outputs.word  = $('settings-out-word').checked;
  settings.outputs.track = $('settings-out-track').checked;
  settings.outputs.pdf   = $('settings-out-pdf').checked;
  settings.outputs.cpo   = $('settings-out-cpo').checked;
  saveSettings();
  applySettingsToOptions();
  toast('설정 저장됨', 'success');
  closeSettings();
});

function applySettingsToOptions() {
  $('out-word').checked  = settings.outputs.word;
  $('out-track').checked = settings.outputs.track;
  $('out-pdf').checked   = settings.outputs.pdf;
  $('out-cpo').checked   = !!settings.outputs.cpo;
}

// Apply settings on load
applySettingsToOptions();
updateOutDirDisplay();
renderQuick();

