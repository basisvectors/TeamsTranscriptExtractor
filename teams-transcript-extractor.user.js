// ==UserScript==
// @name         Teams / Stream Transcript Extractor
// @namespace    teams-transcript-extractor
// @version      1.1
// @description  Captures the virtualized Microsoft Teams meeting recap transcript as it renders and downloads it as text.
// @license      MIT
// @match        https://*.sharepoint.com/*
// @match        https://teams.cloud.microsoft/*
// @match        https://*.cloud.microsoft/*
// @match        https://teams.microsoft.com/*
// @match        https://*.office.com/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  // Transcript rows, keyed by data-list-index so re-renders don't duplicate.
  const rows = new Map();
  let total = 0;
  let panel, statusEl, scrolling = false;

  const text = (el) => (el ? el.textContent.replace(/\s+/g, ' ').trim() : '');

  function captureCell(cell) {
    const idx = Number(cell.getAttribute('data-list-index'));
    if (Number.isNaN(idx)) return;

    const entries = cell.querySelectorAll('[id^="sub-entry-"]');
    if (!entries.length) return;

    const setsize = Number(entries[0].getAttribute('aria-setsize'));
    if (setsize > total) total = setsize;

    const isEvent = !!cell.querySelector('[class*="meetingEvent"]');
    const body = [...entries].map(text).filter(Boolean).join('\n');
    if (!body) return;

    rows.set(idx, {
      idx,
      isEvent,
      speaker: text(cell.querySelector('[class*="itemDisplayName"]')),
      time: text(cell.querySelector('[id^="Header-timestamp-"]')),
      body,
    });
  }

  function scan() {
    document.querySelectorAll('.ms-List-cell[data-list-index]').forEach(captureCell);
    updateStatus();
  }

  function buildTranscript() {
    const sorted = [...rows.values()].sort((a, b) => a.idx - b.idx);
    const out = [];
    let lastSpeaker = '';
    for (const r of sorted) {
      if (r.isEvent) {
        out.push(`[${r.body}]`, '');
        continue;
      }
      // Consecutive messages from the same speaker can render without a header.
      const speaker = r.speaker || lastSpeaker || 'Unknown';
      lastSpeaker = speaker;
      out.push(`${speaker}${r.time ? '  ' + r.time : ''}`, r.body, '');
    }
    return out.join('\n');
  }

  function missingIndices() {
    const missing = [];
    for (let i = 0; i < total; i++) if (!rows.has(i)) missing.push(i);
    return missing;
  }

  function download() {
    const blob = new Blob([buildTranscript()], { type: 'text/plain;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `transcript-${new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-')}.txt`;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  function lastRenderedCell() {
    let best = null, bestIdx = -1;
    document.querySelectorAll('.ms-List-cell[data-list-index]').forEach((c) => {
      const i = Number(c.getAttribute('data-list-index'));
      if (i > bestIdx) { bestIdx = i; best = c; }
    });
    return best;
  }

  async function autoScroll() {
    if (scrolling) { scrolling = false; return; }
    scrolling = true;
    setButtonLabel();

    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
    // Scrolling the last rendered row into view works whichever ancestor is
    // the real scroll container, and makes the virtualized list render more.
    let stuck = 0, lastMax = -1;
    while (scrolling) {
      scan();
      const cell = lastRenderedCell();
      if (!cell) break;
      cell.scrollIntoView({ block: 'start' });
      await sleep(500);
      scan();

      const max = rows.size ? Math.max(...rows.keys()) : -1;
      if (total && max >= total - 1) break;
      if (max === lastMax) {
        if (++stuck >= 6) break; // nothing new is loading
      } else {
        stuck = 0;
        lastMax = max;
      }
    }
    scrolling = false;
    setButtonLabel();
    updateStatus();

    const missing = missingIndices();
    if (missing.length) {
      console.warn('[transcript-extractor] missing indices:', missing);
      statusEl.textContent += ` — ${missing.length} missing, run again or scroll manually`;
    }
  }

  function setButtonLabel() {
    const btn = panel && panel.querySelector('#tte-scroll');
    if (btn) btn.textContent = scrolling ? 'Stop' : 'Auto-scroll & capture';
  }

  function updateStatus() {
    if (!statusEl) return;
    const s = `Captured ${rows.size}${total ? ' / ' + total : ''}`;
    // Only touch the DOM on change, otherwise the MutationObserver feeds itself.
    if (statusEl.textContent !== s) statusEl.textContent = s;
  }

  let scanQueued = false;
  function scheduleScan() {
    if (scanQueued) return;
    scanQueued = true;
    setTimeout(() => { scanQueued = false; scan(); }, 200);
  }

  function makePanel() {
    panel = document.createElement('div');
    panel.style.cssText =
      'position:fixed;left:50%;top:8px;transform:translateX(-50%);z-index:2147483647;pointer-events:auto;' +
      'background:#1f1f1f;color:#fff;' +
      'font:12px/1.4 Segoe UI,sans-serif;padding:8px 10px;border-radius:8px;box-shadow:0 2px 10px #0008;' +
      'display:flex;gap:6px;align-items:center;';
    panel.innerHTML =
      '<span id="tte-status"></span>' +
      '<button id="tte-scroll">Auto-scroll & capture</button>' +
      '<button id="tte-dl">Download .txt</button>' +
      '<button id="tte-copy">Copy</button>' +
      '<button id="tte-reset" title="Clear captured rows">Reset</button>';
    panel.querySelectorAll('button').forEach((b) => {
      b.style.cssText = 'background:#444;color:#fff;border:0;border-radius:4px;padding:3px 8px;cursor:pointer;';
    });
    document.documentElement.appendChild(panel);
    statusEl = panel.querySelector('#tte-status');

    const actions = {
      'tte-scroll': autoScroll,
      'tte-dl': download,
      'tte-copy': () => navigator.clipboard.writeText(buildTranscript()),
      'tte-reset': () => { rows.clear(); total = 0; scan(); },
    };
    // Overlays can sit on top of the panel and swallow clicks, so hit-test the
    // buttons ourselves in the capture phase instead of relying on onclick.
    window.addEventListener('pointerdown', (e) => {
      for (const id of Object.keys(actions)) {
        const r = panel.querySelector('#' + id).getBoundingClientRect();
        if (e.clientX >= r.left && e.clientX <= r.right && e.clientY >= r.top && e.clientY <= r.bottom) {
          e.preventDefault();
          e.stopPropagation();
          actions[id]();
          return;
        }
      }
    }, true);

    // Keyboard fallback: Alt+Shift+S scroll/stop, Alt+Shift+D download.
    window.addEventListener('keydown', (e) => {
      if (!e.altKey || !e.shiftKey) return;
      if (e.code === 'KeyS') { e.preventDefault(); autoScroll(); }
      if (e.code === 'KeyD') { e.preventDefault(); download(); }
    }, true);
    updateStatus();
  }

  console.log('[transcript-extractor] loaded in', location.href);

  // Only activate in the frame that actually hosts the transcript list.
  const boot = setInterval(() => {
    if (!document.querySelector('[id^="sub-entry-"]')) return;
    clearInterval(boot);
    console.log('[transcript-extractor] transcript found, panel added');
    makePanel();
    scan();
    const list = document.querySelector('.ms-List') || document.body;
    new MutationObserver(scheduleScan).observe(list, { childList: true, subtree: true });
    window.__transcriptExtractor = { rows, buildTranscript, download, missingIndices, autoScroll };
  }, 1000);
})();
