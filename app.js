/* ════════════════════════════════════════════════
   Excel 公式圖卡 Pro — App Logic
   ════════════════════════════════════════════════ */

'use strict';

// ── Constants ─────────────────────────────────
const CAT_LABELS = {
  math:     '數學統計',
  lookup:   '查找引用',
  text:     '文字處理',
  logic:    '邏輯判斷',
  datetime: '日期時間',
  advanced: '進階 365',
};

const DIFF_LABELS = { b: '入門', i: '進階', a: '高手' };
const VER_LABELS  = { all: '全版本', '2016': '2016+', '2019': '2019+', '365': '365 限定' };

// ── State ─────────────────────────────────────
let state = {
  cat:    'all',
  search: '',
  diff:   'all',
  ver:    'all',
};

// ── DOM refs ──────────────────────────────────
const $cardGrid   = document.getElementById('cardGrid');
const $comboGrid  = document.getElementById('comboGrid');
const $emptyState = document.getElementById('emptyState');
const $resultsMeta= document.getElementById('resultsMeta');
const $searchInput= document.getElementById('searchInput');
const $searchKbd  = document.getElementById('searchKbd');
const $themeToggle= document.getElementById('themeToggle');
const $cats       = document.querySelectorAll('.cat-btn');
const $diffFilter = document.getElementById('diffFilter');
const $verFilter  = document.getElementById('verFilter');
const $toast      = document.getElementById('toast');

// ── Theme ─────────────────────────────────────
function initTheme() {
  const saved = localStorage.getItem('excel-theme');
  const pref  = saved || (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
  document.documentElement.setAttribute('data-theme', pref);
}

function toggleTheme() {
  const cur  = document.documentElement.getAttribute('data-theme');
  const next = cur === 'dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', next);
  localStorage.setItem('excel-theme', next);
}

// ── Toast ─────────────────────────────────────
let toastTimer = null;
function showToast(msg) {
  $toast.textContent = msg;
  $toast.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => $toast.classList.remove('show'), 2000);
}

// ── Copy ──────────────────────────────────────
function copyText(text, btn) {
  navigator.clipboard.writeText(text).then(() => {
    showToast('✓ 已複製到剪貼簿');
    if (btn) {
      btn.textContent = '✓ 已複製';
      btn.classList.add('copied');
      setTimeout(() => {
        btn.textContent = '複製';
        btn.classList.remove('copied');
      }, 2000);
    }
  }).catch(() => showToast('複製失敗，請手動選取'));
}

// ── Syntax highlighter ────────────────────────
function highlightSyntax(syntax) {
  return syntax
    .replace(/=(\w+)\(/g,          '<span style="color:#7ee787">=</span><span style="color:#d2a8ff">$1</span>(')
    .replace(/\[([^\]]+)\]/g,       '[<span class="arg-opt">$1</span>]')
    .replace(/\b([a-zA-Z_]+)(?=,|\)|\s*$)/g, (m, p1) => {
      if (/^(TRUE|FALSE|AND|OR|NOT)$/.test(p1)) return `<span style="color:#ffa657">${p1}</span>`;
      return m;
    });
}

// ── Filter logic ──────────────────────────────
function matchesFilter(f) {
  if (state.cat !== 'all' && state.cat !== 'combos' && f.cat !== state.cat) return false;
  if (state.diff !== 'all' && f.diff !== state.diff) return false;
  if (state.ver !== 'all' && f.ver !== state.ver) return false;
  if (state.search) {
    const q = state.search.toLowerCase();
    const hay = [f.id, f.zh, f.desc, f.syntax, ...(f.tips || []), ...(f.rel || [])].join(' ').toLowerCase();
    if (!hay.includes(q)) return false;
  }
  return true;
}

// ── Build interactive example ─────────────────
function buildInteractive(f) {
  const fns = window.FORMULA_COMPUTE;

  if (f.itype === 'nums') {
    const vals = [...f.ivals];
    const compute = fns[f.ifn];

    const inputsHtml = vals.map((v, i) =>
      `<label class="ex-input-label">
        <input type="number" class="ex-input num-input" value="${v}" data-index="${i}" />
      </label>`
    ).join('');

    const initial = compute(vals.map(Number));

    return `
      <div class="interactive-ex" data-itype="nums">
        <span class="interactive-label">互動範例</span>
        <div class="inputs-row">${inputsHtml}</div>
        <div class="result-row">
          <span class="result-formula">${buildFormulaPreview(f.id, vals)}</span>
          <span class="result-arrow">→</span>
          <span class="result-value">${formatNum(initial)}</span>
        </div>
      </div>`;
  }

  if (f.itype === 'text1') {
    return `
      <div class="interactive-ex" data-itype="text1">
        <span class="interactive-label">互動範例</span>
        <div class="inputs-row">
          <input type="text" class="ex-input wide text-input" value="${escHtml(f.ival)}" />
        </div>
        <div class="result-row">
          <span class="result-formula">=<b>${f.id}</b>(...)</span>
          <span class="result-arrow">→</span>
          <span class="result-value">${escHtml(String(fns[f.ifn]([f.ival])))}</span>
        </div>
      </div>`;
  }

  if (f.itype === 'text2') {
    const [v0, v1] = f.ivals;
    const p = f.iparams || ['文字', '參數'];
    const result = v1 !== '' ? fns[f.ifn]([v0, v1]) : fns[f.ifn]([v0]);
    const safeResult = result === undefined ? '' : result;
    return `
      <div class="interactive-ex" data-itype="text2">
        <span class="interactive-label">互動範例</span>
        <div class="inputs-row">
          <label class="ex-input-label">
            <span>${p[0]}</span>
            <input type="text" class="ex-input wide text-input" value="${escHtml(String(v0))}" data-param="0" />
          </label>
          ${v1 !== '' ? `<label class="ex-input-label">
            <span>${p[1]}</span>
            <input type="text" class="ex-input text-input" style="width:66px" value="${escHtml(String(v1))}" data-param="1" />
          </label>` : ''}
        </div>
        <div class="result-row">
          <span class="result-formula">=<b>${f.id}</b>(...)</span>
          <span class="result-arrow">→</span>
          <span class="result-value">${escHtml(String(safeResult))}</span>
        </div>
      </div>`;
  }

  if (f.itype === 'text3') {
    const [v0, v1, v2] = f.ivals;
    const p = f.iparams || ['文字', '參數1', '參數2'];
    const result = fns[f.ifn]([v0, v1, v2]);
    return `
      <div class="interactive-ex" data-itype="text3">
        <span class="interactive-label">互動範例</span>
        <div class="inputs-row">
          <label class="ex-input-label">
            <span>${p[0]}</span>
            <input type="text" class="ex-input wide text-input" value="${escHtml(v0)}" data-param="0" />
          </label>
          <label class="ex-input-label">
            <span>${p[1]}</span>
            <input type="text" class="ex-input text-input" style="width:55px" value="${escHtml(v1)}" data-param="1" />
          </label>
          <label class="ex-input-label">
            <span>${p[2]}</span>
            <input type="text" class="ex-input text-input" style="width:55px" value="${escHtml(v2)}" data-param="2" />
          </label>
        </div>
        <div class="result-row">
          <span class="result-formula">=<b>${f.id}</b>(...)</span>
          <span class="result-arrow">→</span>
          <span class="result-value">${escHtml(String(result))}</span>
        </div>
      </div>`;
  }

  if (f.itype === 'date') {
    const now = new Date();
    const examples = {
      TODAY:      now.toLocaleDateString('zh-TW'),
      NOW:        now.toLocaleString('zh-TW'),
      YEAR:       now.getFullYear(),
      MONTH:      now.getMonth() + 1,
      DAY:        now.getDate(),
      WEEKDAY:    ['日','一','二','三','四','五','六'][now.getDay()],
    };
    const val = examples[f.id] || now.toLocaleDateString('zh-TW');
    return `
      <div class="interactive-ex">
        <span class="interactive-label">今日結果</span>
        <div class="result-row">
          <span class="result-formula">=${f.id}()</span>
          <span class="result-arrow">→</span>
          <span class="date-ex">${val}</span>
        </div>
      </div>`;
  }

  if (f.itype === 'table' && f.tableEx) {
    const t = f.tableEx;
    const headerCells = t.headers.map(h => `<th>${escHtml(h)}</th>`).join('');
    const bodyRows = t.rows.map((row, ri) => {
      const cls = ri === 1 ? ' class="highlight-row"' : '';
      return `<tr${cls}>${row.map(c => `<td>${escHtml(String(c))}</td>`).join('')}</tr>`;
    }).join('');
    return `
      <div class="interactive-ex">
        <span class="interactive-label">查找示意</span>
        <div class="mini-table-wrap">
          <table class="mini-table">
            <thead><tr>${headerCells}</tr></thead>
            <tbody>${bodyRows}</tbody>
          </table>
        </div>
        <div class="table-formula">
          <code>${escHtml(t.formula)}</code>
          <span class="result-arrow"> → </span>
          <span class="table-result">${escHtml(t.result)}</span>
        </div>
        <div class="table-note">${escHtml(t.note)}</div>
      </div>`;
  }

  return '';
}

function buildFormulaPreview(id, vals) {
  return `=${id}(${vals.slice(0, 4).join(', ')}${vals.length > 4 ? ', …' : ''})`;
}

function formatNum(n) {
  if (typeof n === 'number' && !isNaN(n)) {
    return Number.isInteger(n) ? n.toString() : n.toFixed(4).replace(/\.?0+$/, '');
  }
  return String(n);
}

function escHtml(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ── Card builder ──────────────────────────────
function buildCard(f) {
  const diffLabel = DIFF_LABELS[f.diff] || f.diff;
  const verLabel  = VER_LABELS[f.ver]   || f.ver;
  const verClass  = f.ver === 'all' ? 'badge-all' : `badge-${f.ver}`;
  const diffClass = `badge-${f.diff}`;

  const argsHtml = (f.args || []).map(a => `
    <div class="arg-item">
      <span class="arg-name ${a.req ? 'req' : 'opt'}">${escHtml(a.n)}</span>
      <span class="arg-desc">${escHtml(a.d)} ${!a.req ? '<span class="arg-req-label">（選填）</span>' : ''}</span>
    </div>`).join('');

  const tipsHtml = (f.tips || []).length
    ? `<ul class="tips-list">${f.tips.map(t => `<li>${escHtml(t)}</li>`).join('')}</ul>`
    : '';

  const relHtml = (f.rel || []).length
    ? `<div class="related-row">
        <span>相關：</span>
        ${f.rel.map(r => `<button class="related-link" data-jump="${r}">${r}</button>`).join('')}
      </div>`
    : '';

  return `
    <div class="formula-card" data-cat="${f.cat}" data-id="${f.id}" role="listitem" tabindex="0">
      <div class="card-header">
        <div class="card-title">
          <span class="formula-name">${escHtml(f.id)}</span>
          <span class="formula-zh">${escHtml(f.zh)}</span>
        </div>
        <div class="card-badges">
          <span class="badge ${diffClass}">${diffLabel}</span>
          <span class="badge ${verClass}">${verLabel}</span>
        </div>
      </div>
      <p class="card-desc">${escHtml(f.desc)}</p>
      <div class="syntax-block">
        <span class="syntax-code">${escHtml(f.syntax)}</span>
        <button class="copy-btn" data-copy="${escHtml(f.syntax)}">複製</button>
      </div>
      ${argsHtml ? `<div class="args-list">${argsHtml}</div>` : ''}
      ${buildInteractive(f)}
      ${tipsHtml}
      ${relHtml}
    </div>`;
}

// ── Combo card builder ────────────────────────
function buildComboCard(c) {
  const diffLabel = DIFF_LABELS[c.diff] || c.diff;
  const verLabel  = VER_LABELS[c.ver]   || c.ver;
  const diffClass = `badge-${c.diff}`;
  const verClass  = c.ver === 'all' ? 'badge-all' : `badge-${c.ver}`;

  return `
    <div class="combo-card" data-cat="combos" role="listitem" tabindex="0">
      <div>
        <span class="combo-tag">🔗 實戰組合</span>
      </div>
      <div>
        <div class="combo-title">${escHtml(c.title)}</div>
        <div class="combo-subtitle">${escHtml(c.subtitle)}</div>
      </div>
      <div class="combo-code">
        <button class="combo-copy-btn" data-copy="${escHtml(c.code)}">複製</button>${escHtml(c.code)}
      </div>
      <div>
        <div class="combo-section-label">說明</div>
        <div class="combo-section-text">${escHtml(c.explanation)}</div>
      </div>
      <div>
        <div class="combo-section-label">使用時機</div>
        <div class="combo-section-text">${escHtml(c.whenToUse)}</div>
      </div>
      ${c.upgrade ? `<div>
        <div class="combo-section-label">升級建議</div>
        <div class="combo-section-text" style="color:var(--primary)">${escHtml(c.upgrade)}</div>
      </div>` : ''}
      <div class="combo-badge-row">
        <span class="badge ${diffClass}">${diffLabel}</span>
        <span class="badge ${verClass}">${verLabel}</span>
      </div>
    </div>`;
}

// ── Render ────────────────────────────────────
function render() {
  const showCombos  = state.cat === 'all' || state.cat === 'combos';
  const showFormulas = state.cat !== 'combos';

  // Filter formulas
  const filtered = showFormulas
    ? window.FORMULAS.filter(matchesFilter)
    : [];

  // Filter combos
  const filteredCombos = showCombos
    ? window.COMBOS.filter(c => {
        if (state.diff !== 'all' && c.diff !== state.diff) return false;
        if (state.ver  !== 'all' && c.ver  !== state.ver)  return false;
        if (state.search) {
          const q   = state.search.toLowerCase();
          const hay = [c.title, c.subtitle, c.explanation, c.whenToUse, c.code].join(' ').toLowerCase();
          if (!hay.includes(q)) return false;
        }
        return true;
      })
    : [];

  const total = filtered.length + filteredCombos.length;

  // Cards
  $cardGrid.innerHTML  = filtered.map(buildCard).join('');
  $comboGrid.innerHTML = filteredCombos.map(buildComboCard).join('');

  // Visibility
  $cardGrid.classList.toggle('hidden',  filtered.length === 0);
  $comboGrid.classList.toggle('hidden', filteredCombos.length === 0);
  $emptyState.classList.toggle('hidden', total > 0);

  // Meta
  const catLabel = state.cat === 'all' ? '全部' :
                   state.cat === 'combos' ? '實戰組合' :
                   CAT_LABELS[state.cat] || state.cat;
  $resultsMeta.textContent = state.search
    ? `搜尋「${state.search}」— 找到 ${total} 個結果`
    : `${catLabel} — 共 ${total} 個`;

  // Bind interactive inputs
  bindInteractiveInputs();
  bindCopyButtons();
  bindRelatedLinks();
  bindCardKeyboard();
}

// ── Interactive input bindings ────────────────
function bindInteractiveInputs() {
  // Number inputs
  $cardGrid.querySelectorAll('.interactive-ex[data-itype="nums"]').forEach(ex => {
    const card = ex.closest('.formula-card');
    const id   = card.dataset.id;
    const f    = window.FORMULAS.find(x => x.id === id);
    if (!f) return;
    const compute = window.FORMULA_COMPUTE[f.ifn];
    const resVal  = ex.querySelector('.result-value');
    const resForm = ex.querySelector('.result-formula');

    ex.querySelectorAll('.num-input').forEach(inp => {
      inp.addEventListener('input', () => {
        const vals = [...ex.querySelectorAll('.num-input')].map(i => parseFloat(i.value) || 0);
        try {
          const result = compute(vals);
          resVal.textContent  = formatNum(result);
          resForm.textContent = buildFormulaPreview(id, vals);
        } catch (e) { resVal.textContent = '#ERR'; }
      });
    });
  });

  // Text1 inputs
  $cardGrid.querySelectorAll('.interactive-ex[data-itype="text1"]').forEach(ex => {
    const card = ex.closest('.formula-card');
    const id   = card.dataset.id;
    const f    = window.FORMULAS.find(x => x.id === id);
    if (!f) return;
    const compute = window.FORMULA_COMPUTE[f.ifn];
    const resVal  = ex.querySelector('.result-value');

    ex.querySelector('.text-input').addEventListener('input', function() {
      try {
        resVal.textContent = escHtml(String(compute([this.value])));
      } catch (e) { resVal.textContent = '#ERR'; }
    });
  });

  // Text2 inputs
  $cardGrid.querySelectorAll('.interactive-ex[data-itype="text2"]').forEach(ex => {
    const card = ex.closest('.formula-card');
    const id   = card.dataset.id;
    const f    = window.FORMULAS.find(x => x.id === id);
    if (!f) return;
    const compute = window.FORMULA_COMPUTE[f.ifn];
    const resVal  = ex.querySelector('.result-value');

    ex.querySelectorAll('.text-input').forEach(inp => {
      inp.addEventListener('input', () => {
        const inputs = [...ex.querySelectorAll('.text-input')].map(i => i.value);
        try {
          const result = compute(inputs);
          resVal.textContent = escHtml(String(result === undefined ? '' : result));
        } catch (e) { resVal.textContent = '#ERR'; }
      });
    });
  });

  // Text3 inputs
  $cardGrid.querySelectorAll('.interactive-ex[data-itype="text3"]').forEach(ex => {
    const card = ex.closest('.formula-card');
    const id   = card.dataset.id;
    const f    = window.FORMULAS.find(x => x.id === id);
    if (!f) return;
    const compute = window.FORMULA_COMPUTE[f.ifn];
    const resVal  = ex.querySelector('.result-value');

    ex.querySelectorAll('.text-input').forEach(inp => {
      inp.addEventListener('input', () => {
        const inputs = [...ex.querySelectorAll('.text-input')].map(i => i.value);
        try {
          resVal.textContent = escHtml(String(compute(inputs)));
        } catch (e) { resVal.textContent = '#ERR'; }
      });
    });
  });
}

// ── Copy button bindings ──────────────────────
function bindCopyButtons() {
  document.querySelectorAll('[data-copy]').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      copyText(btn.dataset.copy, btn);
    };
  });
}

// ── Related link jump ─────────────────────────
function bindRelatedLinks() {
  document.querySelectorAll('.related-link').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      const target = btn.dataset.jump;
      // search for it
      $searchInput.value = target;
      state.search = target;
      state.cat    = 'all';
      $cats.forEach(b => b.classList.toggle('active', b.dataset.cat === 'all'));
      render();
      window.scrollTo({ top: 0, behavior: 'smooth' });
      $searchInput.focus();
    };
  });
}

// ── Keyboard: Enter on card to copy syntax ────
function bindCardKeyboard() {
  document.querySelectorAll('.formula-card[tabindex]').forEach(card => {
    card.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        const f = window.FORMULAS.find(x => x.id === card.dataset.id);
        if (f) copyText(f.syntax);
      }
    });
  });
}

// ── Event listeners ───────────────────────────

// Search (debounced)
let searchTimer;
$searchInput.addEventListener('input', () => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => {
    state.search = $searchInput.value.trim();
    // Searching should always look across the whole library, not just
    // whatever category tab happens to still be active.
    if (state.search && state.cat !== 'all') {
      state.cat = 'all';
      $cats.forEach(b => {
        b.classList.toggle('active', b.dataset.cat === 'all');
        b.setAttribute('aria-selected', b.dataset.cat === 'all' ? 'true' : 'false');
      });
    }
    render();
  }, 160);
});

// Category tabs
$cats.forEach(btn => {
  btn.addEventListener('click', () => {
    state.cat = btn.dataset.cat;
    state.search = '';
    $searchInput.value = '';
    $cats.forEach(b => {
      b.classList.toggle('active', b === btn);
      b.setAttribute('aria-selected', b === btn ? 'true' : 'false');
    });
    render();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });
});

// Filters
$diffFilter.addEventListener('change', () => { state.diff = $diffFilter.value; render(); });
$verFilter.addEventListener('change',  () => { state.ver  = $verFilter.value;  render(); });

// Theme
$themeToggle.addEventListener('click', toggleTheme);

// Keyboard shortcuts
document.addEventListener('keydown', (e) => {
  // '/' to focus search (unless already in input)
  if (e.key === '/' && document.activeElement !== $searchInput && !e.ctrlKey && !e.metaKey) {
    e.preventDefault();
    $searchInput.focus();
    $searchInput.select();
  }
  // Esc to clear search
  if (e.key === 'Escape' && document.activeElement === $searchInput) {
    $searchInput.value = '';
    state.search = '';
    render();
    $searchInput.blur();
  }
});

// Hide / kbd when search focused
$searchInput.addEventListener('focus', ()  => $searchKbd.style.opacity = '0');
$searchInput.addEventListener('blur',  ()  => $searchKbd.style.opacity = '1');

// ── Init ──────────────────────────────────────
initTheme();
render();

// Register service worker for PWA
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  });
}
