/**
 * app.js — SPA Controller for PrintMore
 */

'use strict';

// ===== Constants =====
const STORAGE_KEY = 'printLayouts';

const PAGE_SIZES = {
  A3:     { width: 297,   height: 420 },
  A4:     { width: 210,   height: 297 },
  A5:     { width: 148,   height: 210 },
  Letter: { width: 215.9, height: 279.4 },
  Legal:  { width: 215.9, height: 355.6 },
};

function getPageSizeMm(page) {
  if ((page?.size || 'A4') === 'custom') {
    let w = Math.max(20, parseFloat(page?.customWidthMm ?? page?.customWidth) || 210);
    let h = Math.max(20, parseFloat(page?.customHeightMm ?? page?.customHeight) || 297);
    if ((page?.orientation || 'portrait') === 'landscape') [w, h] = [h, w];
    return { width: w, height: h };
  }
  let { width, height } = PAGE_SIZES[page?.size] || PAGE_SIZES.A4;
  if ((page?.orientation || 'portrait') === 'landscape') [width, height] = [height, width];
  return { width, height };
}

// ===== State =====
let currentView = 'home';
let currentLayoutId = null;
let designerInstance = null;
let shareLayoutModalId = null;
let currentSessionPassword = '';
let selectedManagedUser = null;

// Parsed data storage for Run view
let runParsedData = null;
let manualRowCount = 1;
let smartRulesDraft = null;
let homeSearchQuery = '';
let managedUsersCache = [];
let managedLayoutsUser = null;
let managedUsersSearchQuery = '';

function getCurrentUser() {
  return window.AuthStore ? window.AuthStore.currentUser() : null;
}

function updateUserChrome() {
  const user = getCurrentUser();
  const label = document.getElementById('current-user-label');
  const addUserBtn = document.getElementById('btn-add-user');
  const smartUiBtn = document.getElementById('btn-smart-ui');
  const canDesign = user?.role === 'designer' || user?.role === 'super';

  if (label) label.textContent = user ? `User: ${user.username}` : '';
  if (addUserBtn) addUserBtn.classList.toggle('hidden', !user?.isSuperUser);
  if (smartUiBtn) smartUiBtn.classList.toggle('hidden', !user?.isSuperUser);
  document.getElementById('btn-new-layout')?.classList.toggle('hidden', !canDesign);
}

function showLoginView(message) {
  showView('login');
  const error = document.getElementById('login-error');
  if (error) {
    error.textContent = message || '';
    error.classList.toggle('hidden', !message);
  }
}

async function loadCurrentUserLayouts() {
  const user = getCurrentUser();
  if (!user) {
    showLoginView();
    return false;
  }

  if (window.LayoutStore) {
    await window.LayoutStore.init(user);
    await window.LayoutStore.loadSmartRules?.();
    updateStorageStatus();
  }
  updateUserChrome();
  return true;
}

function updateStorageStatus() {
  const el = document.getElementById('storage-status');
  if (!el || !window.LayoutStore) return;

  const status = window.LayoutStore.status();
  const labels = {
    loading: 'Loading layouts...',
    supabase: 'Supabase connected',
    offline: 'Supabase offline - local copy active',
    local: 'Local browser storage',
  };

  el.textContent = labels[status] || labels.local;
  el.dataset.status = status;
}

// ===== Utility =====
function generateId() {
  return 'lay-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7);
}

function formatDate(isoString) {
  try {
    const d = new Date(isoString);
    return d.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
  } catch {
    return isoString;
  }
}

// ===== LocalStorage =====
function loadLayouts() {
  return window.LayoutStore ? window.LayoutStore.getAll() : [];
}

function saveLayouts(layouts) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(layouts));
}

function getLayoutById(id) {
  return window.LayoutStore ? window.LayoutStore.getById(id) : null;
}

function saveLayout(layout) {
  if (!window.LayoutStore) return;
  const nowIso = new Date().toISOString();
  layout.createdAt = layout.createdAt || nowIso;
  layout.updatedAt = nowIso;
  window.LayoutStore.save(layout).then(result => {
    updateStorageStatus();
    if (result && !result.ok) showToast('Saved locally. Supabase sync failed.');
  });
}

function deleteLayout(id) {
  if (!window.LayoutStore) return;
  window.LayoutStore.remove(id).then(result => {
    updateStorageStatus();
    if (result && !result.ok) showToast('Deleted locally. Supabase sync failed.');
  });
}

// ===== Export / Import =====
function exportLayout(id) {
  const layout = getLayoutById(id);
  if (!layout) return;
  exportLayoutData(layout);
}

function exportLayoutData(layout) {
  if (!layout) return;
  const json = JSON.stringify(layout, null, 2);
  const blob = new Blob([json], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = (layout.name || 'layout').replace(/[^a-z0-9_\-]/gi, '_') + '.json';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function importLayoutFromFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = JSON.parse(e.target.result);
      if (!data || typeof data !== 'object') throw new Error('Invalid format');
      if (!data.name || !data.page) throw new Error('Missing required layout properties');

      // Check for name collision
      const existing = loadLayouts();
      const nameExists = (n) => existing.some(l => l.name.toLowerCase() === n.trim().toLowerCase());
      if (nameExists(data.name)) {
        const newName = prompt(
          `A layout named "${data.name}" already exists.\nEnter a different name for the imported layout:`,
          data.name + ' (imported)'
        );
        if (!newName || !newName.trim()) {
          return; // user cancelled
        }
        if (nameExists(newName)) {
          alert(`A layout named "${newName.trim()}" also already exists. Import cancelled.`);
          return;
        }
        data.name = newName.trim();
      }

      // Assign a fresh ID to avoid collisions
      data.id = generateId();
      data.importedAt = new Date().toISOString();
      saveLayout(data);
      renderHomeView();
      showToast('Layout "' + data.name + '" imported!');
    } catch (err) {
      alert('Import failed: ' + err.message + '\nMake sure you selected a valid layout JSON file.');
    }
  };
  reader.readAsText(file);
}

// ===== Field name parser =====
function parseFieldNames(text) {
  const seen = new Set();
  const result = [];
  const rows = text.split(/\r?\n/);
  for (const row of rows) {
    const cells = row.split('\t');
    for (const cell of cells) {
      const name = cell.trim();
      if (name && !seen.has(name)) {
        seen.add(name);
        result.push(name);
      }
    }
  }
  return result;
}

// ===== View Switching =====
function showView(viewName) {
  document.querySelectorAll('.view').forEach(v => {
    v.classList.remove('active');
    v.classList.add('hidden');
  });
  const target = document.getElementById('view-' + viewName);
  if (target) {
    target.classList.remove('hidden');
    target.classList.add('active');
  }
  currentView = viewName;
}

// ===== Home View =====
function renderHomeView() {
  showView('home');
  updateStorageStatus();
  const grid = document.getElementById('layouts-grid');
  const noMsg = document.getElementById('no-layouts-msg');
  const searchInput = document.getElementById('layout-search-input');
  const layouts = loadLayouts();
  const user = getCurrentUser();
  const canDesign = user?.role === 'designer' || user?.role === 'super';
  const search = (homeSearchQuery || '').trim().toLowerCase();
  if (searchInput && searchInput.value !== homeSearchQuery) searchInput.value = homeSearchQuery;
  const visibleLayouts = [...layouts]
    .sort((a, b) => {
      const ta = Date.parse(a?.updatedAt || a?.createdAt || 0) || 0;
      const tb = Date.parse(b?.updatedAt || b?.createdAt || 0) || 0;
      return tb - ta;
    })
    .filter(layout => !search || String(layout?.name || '').toLowerCase().includes(search));

  grid.innerHTML = '';
  if (visibleLayouts.length === 0) {
    noMsg.classList.remove('hidden');
    const noMsgText = noMsg.querySelector('p');
    if (noMsgText) {
      noMsgText.innerHTML = search
        ? 'No layouts match your search.'
        : 'No layouts yet. Click <strong>+ New Layout</strong> to get started.';
    }
    return;
  }
  noMsg.classList.add('hidden');

  visibleLayouts.forEach(layout => {
    const card = document.createElement('div');
    card.className = 'layout-card';
    const page = layout.page || {};
    const orient = page.orientation === 'landscape' ? 'Landscape' : 'Portrait';
    card.innerHTML = `
      <div class="layout-card-header">
        <div class="layout-card-name">${escapeHtml(layout.name)}</div>
        <div class="layout-card-badge">${page.size || 'A4'} · ${orient}</div>
      </div>
      <div class="layout-card-meta">
        <span>${(layout.fields || []).length} field(s) · ${(layout.elements || []).length} element(s)</span>
        <span>Updated ${formatDate(layout.updatedAt || layout.createdAt)}</span>
      </div>
      <div class="layout-card-actions">
        ${canDesign ? `<button class="btn btn-primary btn-sm" data-action="edit" data-id="${layout.id}">Edit</button>` : ''}
        <button class="btn btn-secondary btn-sm" data-action="run" data-id="${layout.id}">Run</button>
        <button class="btn btn-ghost btn-sm" data-action="share" data-id="${layout.id}" title="Share or download">Share</button>
        ${canDesign ? `<button class="btn btn-danger btn-sm" data-action="delete" data-id="${layout.id}">Delete</button>` : ''}
      </div>
    `;
    grid.appendChild(card);
  });
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

// ===== Setup View =====
let setupStep = 1;

function showSetupView(prefill) {
  setupStep = 1;
  showView('setup');
  showSetupStep(1);
  const customWidth = prefill?.page?.customWidthMm ?? prefill?.page?.customWidth ?? 210;
  const customHeight = prefill?.page?.customHeightMm ?? prefill?.page?.customHeight ?? 297;
  document.getElementById('layout-name').value = prefill?.name || '';
  document.getElementById('page-size').value = prefill?.page?.size || 'A4';
  document.getElementById('page-orientation').value = prefill?.page?.orientation || 'portrait';
  document.getElementById('custom-page-width').value = customWidth;
  document.getElementById('custom-page-height').value = customHeight;
  document.getElementById('margin-top').value = prefill?.page?.marginTop ?? 5;
  document.getElementById('margin-bottom').value = prefill?.page?.marginBottom ?? 5;
  document.getElementById('margin-left').value = prefill?.page?.marginLeft ?? 5;
  document.getElementById('margin-right').value = prefill?.page?.marginRight ?? 5;
  document.getElementById('default-font-family').value = prefill?.defaultStyle?.fontFamily || 'Arial';
  document.getElementById('default-font-size').value = prefill?.defaultStyle?.fontSize || 12;
  document.getElementById('default-font-color').value = prefill?.defaultStyle?.color || '#000000';
  document.getElementById('field-names-input').value = (prefill?.fields || []).join('\t');
  document.getElementById('layout-template').value = 'blank';
  const ftList = document.getElementById('field-type-list');
  if (ftList) ftList.innerHTML = '';
  document.getElementById('parsed-fields-preview').classList.add('hidden');
  toggleSetupCustomSize();
  renderTemplateOptions();
}

function showSetupStep(n) {
  setupStep = n;
  document.getElementById('setup-step-1').classList.toggle('hidden', n !== 1);
  document.getElementById('setup-step-2').classList.toggle('hidden', n !== 2);
  document.getElementById('setup-step-3').classList.toggle('hidden', n !== 3);
  document.getElementById('step-ind-1').classList.toggle('active', n === 1);
  document.getElementById('step-ind-1').classList.toggle('done', n > 1);
  document.getElementById('step-ind-2').classList.toggle('active', n === 2);
  document.getElementById('step-ind-2').classList.toggle('done', n > 2);
  document.getElementById('step-ind-3').classList.toggle('active', n === 3);
  if (n === 3) renderTemplateOptions();
}

function toggleSetupCustomSize() {
  const size = document.getElementById('page-size')?.value || 'A4';
  document.getElementById('custom-page-size-group')?.classList.toggle('hidden', size !== 'custom');
}

function createNewLayout() {
  const name = document.getElementById('layout-name').value.trim();
  const finalName = name || 'Untitled Layout';
  const existing = loadLayouts();
  const hasDuplicate = existing.some(l => String(l.name || '').trim().toLowerCase() === finalName.toLowerCase());
  if (hasDuplicate) {
    alert(`Layout name "${finalName}" already exists. Please use a unique name.`);
    document.getElementById('layout-name').focus();
    return;
  }
  const size = document.getElementById('page-size').value;
  const orientation = document.getElementById('page-orientation').value;
  const customWidthMm = Math.max(20, parseFloat(document.getElementById('custom-page-width')?.value) || 210);
  const customHeightMm = Math.max(20, parseFloat(document.getElementById('custom-page-height')?.value) || 297);
  const marginTop = parseFloat(document.getElementById('margin-top').value) || 5;
  const marginBottom = parseFloat(document.getElementById('margin-bottom').value) || 5;
  const marginLeft = parseFloat(document.getElementById('margin-left').value) || 5;
  const marginRight = parseFloat(document.getElementById('margin-right').value) || 5;
  const defaultStyle = {
    fontFamily: document.getElementById('default-font-family').value || 'Arial',
    fontSize: parseFloat(document.getElementById('default-font-size').value) || 12,
    color: document.getElementById('default-font-color').value || '#000000',
    fontWeight: 'normal',
    fontStyle: 'normal',
    textDecoration: 'none',
    textAlign: 'left',
  };
  const fieldText = document.getElementById('field-names-input').value;
  const fields = parseFieldNames(fieldText);

  const nowIso = new Date().toISOString();
  const layout = {
    id: generateId(),
    name: finalName,
    createdAt: nowIso,
    updatedAt: nowIso,
    page: { size, orientation, marginTop, marginBottom, marginLeft, marginRight, customWidthMm, customHeightMm },
    fields,
    texts: fields.slice(),
    fieldMeta: {},
    defaultStyle,
    sharedWithUsernames: [],
    elements: [],
  };
  layout.elements = buildTemplateElements(document.getElementById('layout-template').value, layout);

  saveLayout(layout);
  openDesigner(layout.id);
}

const LAYOUT_TEMPLATES = [
  { id: 'blank', label: 'Blank Layout', meta: 'Start with an empty page.', pattern: 'blank' },
  { id: 'layout-1', label: 'Layout 1', meta: 'Header block with repeating line table.', pattern: 'receipt' },
  { id: 'layout-2', label: 'Layout 2', meta: 'Large item table with compact header.', pattern: 'pick' },
  { id: 'layout-3', label: 'Layout 3', meta: 'Shipping-style header and carrier section.', pattern: 'ship' },
  { id: 'layout-4', label: 'Layout 4', meta: 'Document summary with signature/footer area.', pattern: 'delivery' },
  { id: 'layout-5', label: 'Layout 5', meta: 'Wide operational checklist style.', pattern: 'checklist' },
  { id: 'smart', label: 'Layout 6', meta: 'Suggests a structure from your fields.', pattern: 'smart' },
];

function renderTemplateOptions() {
  const container = document.getElementById('template-options');
  if (!container) return;
  const selected = document.getElementById('layout-template').value || 'blank';
  container.innerHTML = '';
  LAYOUT_TEMPLATES.forEach(tpl => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'template-option' + (tpl.id === selected ? ' active' : '');
    btn.dataset.template = tpl.id;
    btn.innerHTML = `<div class="template-option-title">${escapeHtml(tpl.label)}</div><div class="template-option-meta">${escapeHtml(tpl.meta)}</div>`;
    btn.addEventListener('click', () => {
      document.getElementById('layout-template').value = tpl.id;
      renderTemplateOptions();
    });
    container.appendChild(btn);
  });
  renderTemplatePreview(selected);
}

function getSmartRules() {
  return window.LayoutStore?.getSmartRules?.() || {
    header: ['customer', 'date', 'document', 'order', 'reference'],
    table: ['item', 'sku', 'description', 'quantity', 'qty'],
    footer: ['signature', 'remarks', 'total', 'user'],
  };
}

function compactKey(value) {
  return String(value || '').toLowerCase().replace(/[\s_-]+/g, '');
}

function smartLabel(value) {
  const text = String(value || '').trim();
  return text ? text.charAt(0).toUpperCase() + text.slice(1) : '';
}

function matchesSmartRule(field, rules) {
  const key = compactKey(field);
  return (rules || []).some(rule => {
    const ruleKey = compactKey(rule);
    return ruleKey && (key.includes(ruleKey) || ruleKey.includes(key));
  });
}

function classifyLayoutFields(fields) {
  const sourceFields = fields || [];
  const rules = getSmartRules();
  const buckets = { header: [], table: [], footer: [] };
  sourceFields.forEach(field => {
    if (matchesSmartRule(field, rules.table)) {
      buckets.table.push(field);
    } else if (matchesSmartRule(field, rules.header)) {
      buckets.header.push(field);
    } else if (matchesSmartRule(field, rules.footer)) {
      buckets.footer.push(field);
    }
  });
  return buckets;
}

function templateContext(fields) {
  const hasInput = (fields || []).length > 0;
  const allFields = hasInput ? fields : ['Customer', 'Date', 'Item', 'Qty', 'Remarks'];
  const buckets = classifyLayoutFields(allFields);
  return {
    fields: allFields,
    header: hasInput ? buckets.header : (buckets.header.length ? buckets.header : allFields.slice(0, 3)),
    table: hasInput ? buckets.table : (buckets.table.length ? buckets.table : allFields.slice(0, 4)),
    footer: hasInput ? buckets.footer : (buckets.footer.length ? buckets.footer : ['Remarks', 'Signature']),
  };
}

function renderTemplatePreview(templateId) {
  const preview = document.getElementById('template-preview');
  if (!preview) return;
  const ctx = templateContext(parseFieldNames(document.getElementById('field-names-input').value));
  if (templateId === 'blank') {
    preview.innerHTML = '<div class="tpl-title">Blank Layout</div><div class="tpl-line"></div>';
    return;
  }
  const pattern = templateId === 'smart' ? smartTemplateKey({ name: document.getElementById('layout-name').value, fields: ctx.fields }) : templateId;
  const tableLabels = ctx.table.length ? ctx.table : [];
  const box = (label, left, top, width = 96, height = 11) => `<div class="tpl-box" style="left:${left}px;top:${top}px;width:${width}px;height:${height}px;"><span style="font-size:7px;position:absolute;left:4px;top:2px;">${escapeHtml(label)}</span></div>`;
  const table = (top, rows = 4, left = 16, width = 223) => `
    <table class="tpl-table" style="top:${top}px;left:${left}px;width:${width}px;">
      <tr>${tableLabels.map(l => `<th>${escapeHtml(smartLabel(l))}</th>`).join('')}</tr>
      ${Array.from({ length: rows }).map(() => `<tr>${tableLabels.map(() => '<td>&nbsp;</td>').join('')}</tr>`).join('')}
    </table>`;
  const tableHtml = (top, rows = 4, left = 16, width = 223) => tableLabels.length ? table(top, rows, left, width) : '';
  let body = '';
  if (pattern === 'layout-1') {
    body = ctx.header.slice(0, 4).map((l, i) => box(l, 16 + (i % 2) * 112, 46 + Math.floor(i / 2) * 13, 96, 11)).join('') + tableHtml(76, 4);
  } else if (pattern === 'layout-2') {
    body = ctx.header.slice(0, 3).map((l, i) => box(l, 16 + i * 74, 46, 65, 11)).join('') + tableHtml(62, 6);
  } else if (pattern === 'layout-3') {
    body = (ctx.header[0] ? box(ctx.header[0], 16, 46, 104, 11) : '') + (ctx.header[1] ? box(ctx.header[1], 135, 46, 104, 11) : '') + (ctx.header[2] ? box(ctx.header[2], 16, 61, 223, 11) : '') + tableHtml(78, 4);
  } else if (pattern === 'layout-4') {
    const footer = ctx.footer.slice(0, 2);
    body =
      ctx.header.slice(0, 4).map((l, i) => box(l, 16 + (i % 2) * 112, 44 + Math.floor(i / 2) * 12, 96, 10)).join('') +
      tableHtml(70, 3) +
      (footer[0] ? box(footer[0], 16, 114, 223, 16) : '') +
      (footer[1] ? box(footer[1], 150, 134, 89, 13) : '');
  } else {
    body = ctx.header.slice(0, 8).map((l, i) => box(l, 16 + (i % 2) * 112, 46 + Math.floor(i / 2) * 13, 96, 11)).join('') +
      (ctx.footer[0] ? box(ctx.footer[0], 16, 104, 223, 20) : '') +
      (ctx.footer[1] ? box(ctx.footer[1], 16, 132, 100, 15) : '');
  }
  preview.innerHTML = `
    <div class="tpl-title">Layout Preview</div>
    <div class="tpl-line"></div>
    ${body}
    <div class="tpl-footer">Printed by user</div>
  `;
}

function buildTemplateElements(template, layout) {
  if (template === 'blank') return [];
  const fields = layout.fields || [];
  const style = layout.defaultStyle || {};
  const pageMm = getPageSizeMm(layout.page || { size: 'A4', orientation: 'portrait' });
  const sx = pageMm.width / 210;
  const sy = pageMm.height / 297;
  const sf = Math.max(0.45, Math.min(1, Math.min(sx, sy)));
  const mx = layout.page?.marginLeft ?? 5;
  const my = layout.page?.marginTop ?? 5;
  const contentW = Math.max(80, pageMm.width - (layout.page?.marginLeft ?? 5) - (layout.page?.marginRight ?? 5));
  const toX = v => +(mx + v * sx).toFixed(2);
  const toY = v => +(my + v * sy).toFixed(2);
  const toW = v => +(v * sx).toFixed(2);
  const toH = v => +(v * sy).toFixed(2);
  const fs = (value, min = 5) => Math.max(min, +(value * sf).toFixed(2));
  const picked = template === 'smart' ? smartTemplateKey(layout) : template;
  const ctx = templateContext(fields);
  const title = layout.name || 'DOCUMENT';
  const titleFont = fs(15, 6);
  const titleWidthMm = Math.max(
    10,
    Math.min(contentW * 0.8, (String(title).length * titleFont * 0.2) + 6)
  );
  const makeId = p => 'el-' + p + '-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 5);
  const baseStyle = { ...style, backgroundColor: 'transparent', borderWidth: 0, borderColor: '#000000', borderStyle: 'solid', opacity: 1 };
  const elements = [
    { id: makeId('title'), type: 'text', x: toX(0), y: toY(0), width: toW(titleWidthMm), height: toH(8), content: title, fieldName: '', imageData: '', style: { ...baseStyle, fontSize: titleFont, fontWeight: 'bold' } },
    { id: makeId('user'), type: 'user', x: toX(contentW - 45), y: toY(1), width: toW(45), height: toH(6), content: '', fieldName: '', imageData: '', style: { ...baseStyle, fontSize: fs(8, 5), textAlign: 'right' } },
    { id: makeId('line'), type: 'line', x: toX(0), y: toY(12), width: toW(contentW), height: toH(2), lineDirection: 'horizontal', style: { ...baseStyle, borderWidth: 1, borderColor: '#000000' } },
  ];
  const addPair = (field, x, y, w = 86) => {
    elements.push({ id: makeId('lbl'), type: 'text', x: toX(x), y: toY(y), width: toW(28), height: toH(5), content: smartLabel(field), fieldName: '', imageData: '', style: { ...baseStyle, fontSize: fs(7.5, 4.5), fontWeight: 'bold' } });
    elements.push({ id: makeId('fld'), type: 'field', x: toX(x + 29), y: toY(y), width: toW(Math.max(12, w - 29)), height: toH(5), content: '', fieldName: field, imageData: '', style: { ...baseStyle, fontSize: fs(7.5, 4.5) } });
  };
  const addFieldGrid = (fieldList, x, y, cols = 2, colWidth = 95, rowGap = 6, cellWidth = 86) => {
    fieldList.forEach((field, index) => {
      addPair(field, x + (index % cols) * colWidth, y + Math.floor(index / cols) * rowGap, cellWidth);
    });
    return y + Math.ceil(fieldList.length / cols) * rowGap;
  };
  const headerFields = ctx.header;
  const headerCols = picked === 'layout-2' ? 3 : 2;
  const headerColWidth = picked === 'layout-2' ? 64 : 95;
  const headerCellWidth = picked === 'layout-2' ? 58 : 86;
  let headerEndY = 24;
  if (picked === 'layout-2') {
    headerEndY = addFieldGrid(headerFields, 0, 18, headerCols, headerColWidth, 6, headerCellWidth);
  } else if (picked === 'layout-3') {
    headerEndY = addFieldGrid(headerFields, 0, 18, 2, 95, 6, 86);
  } else if (picked === 'layout-4') {
    headerEndY = addFieldGrid(headerFields, 0, 18, 2, 95, 6, 86);
  } else {
    headerEndY = addFieldGrid(headerFields, 0, 18, 2, 95, 6, 86);
  }

  if (picked === 'layout-5') {
    const remarksY = Math.max(34, headerEndY + 4);
    if (ctx.footer.length) {
      elements.push({ id: makeId('remarks'), type: 'rect', x: toX(0), y: toY(remarksY), width: toW(contentW), height: toH(18), content: '', fieldName: '', imageData: '', style: { ...baseStyle, borderWidth: 1 } });
      elements.push({ id: makeId('remarkslbl'), type: 'text', x: toX(2), y: toY(remarksY + 2), width: toW(60), height: toH(5), content: smartLabel(ctx.footer[0]), fieldName: '', imageData: '', style: { ...baseStyle, fontSize: fs(7.5, 4.5), fontWeight: 'bold' } });
      elements.push({ id: makeId('remarksfld'), type: 'field', x: toX(2), y: toY(remarksY + 8), width: toW(contentW - 4), height: toH(7), content: '', fieldName: ctx.footer[0], imageData: '', style: { ...baseStyle, fontSize: fs(7.5, 4.5) } });
    }
    return elements;
  }

  const tableFields = ctx.table;
  const tableY = Math.max(picked === 'layout-2' ? 28 : (picked === 'layout-3' ? 36 : 32), headerEndY + 4);
  if (tableFields.length) {
    const tableCols = tableFields.length;
    const cells = tableFields.flatMap((field, col) => [
      { row: 0, col, fieldName: '', content: smartLabel(field), style: { fontWeight: 'bold' } },
      { row: 1, col, fieldName: field, content: '', style: {} },
    ]);
    elements.push({
      id: makeId('tbl'), type: 'table', x: toX(0), y: toY(tableY), width: toW(contentW), height: toH(10),
      content: '', fieldName: '', imageData: '',
      style: { ...baseStyle, fontSize: fs(tableCols > 6 ? 7 : 8, 4.5), borderWidth: 1, borderColor: '#000000', borderStyle: 'solid' },
      table: { rows: 2, cols: tableCols, cells, theme: 'plain', borderMode: 'all', colWidths: Array(tableCols).fill(1), rowHeights: [5, 5], detailMode: true, colProps: [] },
    });
  }
  if (ctx.footer.length) {
    addFieldGrid(ctx.footer, 0, (tableFields.length ? tableY + 16 : Math.max(headerEndY + 6, 48)), 2, 95, 6, 86);
  }
  return elements;
}

function smartTemplateKey(layout) {
  const haystack = `${layout.name || ''} ${(layout.fields || []).join(' ')}`.toLowerCase();
  if (haystack.includes('pick')) return 'layout-2';
  if (haystack.includes('ship') || haystack.includes('carrier') || haystack.includes('bol')) return 'layout-3';
  if (haystack.includes('delivery') || haystack.includes('dispatch')) return 'layout-4';
  if (haystack.includes('check') || haystack.includes('load')) return 'layout-5';
  return 'layout-1';
}

// ===== Designer View =====
function openDesigner(layoutId) {
  const user = getCurrentUser();
  if (!(user?.role === 'designer' || user?.role === 'super')) {
    alert('You do not have access to edit layouts.');
    renderHomeView();
    return;
  }
  currentLayoutId = layoutId;
  showView('designer');

  const layout = getLayoutById(layoutId);
  if (!layout) {
    alert('Layout not found.');
    renderHomeView();
    return;
  }

  document.getElementById('toolbar-layout-name').textContent = layout.name;

  if (designerInstance) {
    designerInstance.destroy();
  }
  designerInstance = new Designer(layoutId);
  designerInstance.init();
}

// ===== Run View =====
function openRunView(layoutId) {
  currentLayoutId = layoutId;
  const layout = getLayoutById(layoutId);
  if (!layout) {
    alert('Layout not found.');
    return;
  }

  showView('run');
  document.getElementById('run-layout-name').textContent = layout.name;

  // Reset state
  runParsedData = null;
  const pasteArea = document.getElementById('run-paste-area');
  if (pasteArea) pasteArea.value = '';
  const parsedPreview = document.getElementById('run-parsed-preview');
  if (parsedPreview) parsedPreview.classList.add('hidden');
  const pasteStatus = document.getElementById('paste-status');
  if (pasteStatus) pasteStatus.textContent = '';

  // Reset tabs to show paste tab
  document.querySelectorAll('.run-tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.run-tab-content').forEach(c => c.classList.remove('active'));
  const pasteTab = document.querySelector('.run-tab[data-tab="paste"]');
  if (pasteTab) pasteTab.classList.add('active');
  const pasteTabContent = document.getElementById('run-tab-paste');
  if (pasteTabContent) pasteTabContent.classList.add('active');
  document.getElementById('btn-add-manual-row')?.classList.add('hidden');

  // Populate manual form
  const form = document.getElementById('run-fields-form');
  form.innerHTML = '';
  const fields = layout.fields || [];
  if (fields.length === 0) {
    form.innerHTML = '<p class="hint">This layout has no fields defined.</p>';
  } else {
    fields.forEach(fieldName => {
      const group = document.createElement('div');
      group.className = 'form-group';
      group.innerHTML = `
        <label>${escapeHtml(fieldName)}</label>
        <input type="text" data-field="${escapeHtml(fieldName)}" placeholder="Enter ${escapeHtml(fieldName)}…" />
      `;
      form.appendChild(group);
    });
  }
  manualRowCount = 1;
  renderManualRows(layout);
}

function renderManualRows(layout) {
  const form = document.getElementById('run-fields-form');
  const existingValues = [];
  form.querySelectorAll('input[data-field]').forEach(input => {
    const rowIndex = parseInt(input.dataset.row || '0', 10);
    existingValues[rowIndex] = existingValues[rowIndex] || {};
    existingValues[rowIndex][input.dataset.field] = input.value;
  });
  form.innerHTML = '';
  const fields = layout.fields || [];
  if (fields.length === 0) {
    form.innerHTML = '<p class="hint">This layout has no fields defined.</p>';
    return;
  }

  for (let rowIndex = 0; rowIndex < manualRowCount; rowIndex++) {
    const row = document.createElement('div');
    row.className = 'manual-entry-row';
    row.innerHTML = `<div class="manual-row-title">Line ${rowIndex + 1}</div>`;
    fields.forEach(fieldName => {
      const group = document.createElement('div');
      group.className = 'form-group';
      group.innerHTML = `
        <label>${escapeHtml(fieldName)}</label>
        <input type="text" data-row="${rowIndex}" data-field="${escapeHtml(fieldName)}" value="${escapeHtml(existingValues[rowIndex]?.[fieldName] || '')}" placeholder="Enter ${escapeHtml(fieldName)}" />
      `;
      row.appendChild(group);
    });
    form.appendChild(row);
  }
}

function parsePasteData(text, fields) {
  const rows = text.trim().split(/\r?\n/).map(r => r.split('\t'));
  if (rows.length === 0) return null;

  let headers, dataRows;
  // Check if first row matches field names (case-insensitive)
  const firstRow = rows[0].map(c => c.trim());
  const fieldLower = fields.map(f => f.toLowerCase());
  const firstRowMatchesFields = firstRow.some(h => fieldLower.includes(h.toLowerCase()));

  if (firstRowMatchesFields) {
    headers = firstRow;
    dataRows = rows.slice(1);
  } else {
    // Assume same order as fields
    headers = fields;
    dataRows = rows;
  }

  return dataRows.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      // Find matching field (case-insensitive)
      const matchingField = fields.find(f => f.toLowerCase() === h.toLowerCase()) || h;
      obj[matchingField] = (row[i] || '').trim();
    });
    return obj;
  }).filter(row => Object.values(row).some(v => v !== ''));
}

function renderParsedPreview(data, fields) {
  const preview = document.getElementById('run-parsed-preview');
  if (!data || data.length === 0) {
    if (preview) preview.classList.add('hidden');
    return;
  }

  const displayFields = fields.filter(f => data[0].hasOwnProperty(f));
  let html = `<p class="hint" style="margin-bottom:8px;">${data.length} row(s) parsed. First 5 shown:</p>`;
  html += '<table class="run-parsed-table"><thead><tr>';
  displayFields.forEach(f => { html += `<th>${escapeHtml(f)}</th>`; });
  html += '</tr></thead><tbody>';
  data.slice(0, 5).forEach(row => {
    html += '<tr>';
    displayFields.forEach(f => { html += `<td>${escapeHtml(row[f] || '')}</td>`; });
    html += '</tr>';
  });
  html += '</tbody></table>';

  if (preview) {
    preview.innerHTML = html;
    preview.classList.remove('hidden');
  }
}

async function shareLayoutToUser(layoutId, enteredUserId) {
  const current = getCurrentUser();
  if (!(current?.role === 'designer' || current?.role === 'super')) {
    return { ok: false, message: 'Only designer or super user can share layouts.' };
  }

  const source = getLayoutById(layoutId);
  if (!source) {
    return { ok: false, message: 'Layout not found.' };
  }

  const targetUsername = String(enteredUserId || '').trim().toUpperCase();
  if (!targetUsername) return { ok: false, message: 'Enter user id.' };

  try {
    const targetUser = await window.AuthStore?.findUser?.(targetUsername);
    if (!targetUser) {
      return { ok: false, message: 'Invalid user id.' };
    }

    const copy = JSON.parse(JSON.stringify(source));
    copy.id = generateId();
    copy.createdAt = new Date().toISOString();
    copy.updatedAt = new Date().toISOString();
    copy.userId = targetUser.id;
    copy.username = targetUser.username;
    delete copy.sharedWithUsernames;

    const result = await window.LayoutStore?.saveForUser?.(copy, targetUser);
    if (!result?.ok) throw (result?.error || new Error('Could not share layout.'));
    return { ok: true, message: `Layout shared to ${targetUser.username}.` };
  } catch (err) {
    return { ok: false, message: err.message || 'Could not share layout.' };
  }
}

function setShareLayoutTab(tabName) {
  document.querySelectorAll('.modal-tab[data-share-tab]').forEach(tab => {
    tab.classList.toggle('active', tab.dataset.shareTab === tabName);
  });
  document.getElementById('share-layout-tab-share')?.classList.toggle('active', tabName === 'share');
  document.getElementById('share-layout-tab-download')?.classList.toggle('active', tabName === 'download');
  document.getElementById('btn-share-layout-send')?.classList.toggle('hidden', tabName !== 'share');
  document.getElementById('btn-share-layout-download')?.classList.toggle('hidden', tabName !== 'download');
}

function closeShareLayoutModal() {
  document.getElementById('modal-share-layout')?.classList.add('hidden');
  shareLayoutModalId = null;
}

function openShareLayoutModal(layoutId) {
  shareLayoutModalId = layoutId;
  document.getElementById('share-layout-user-id').value = '';
  document.getElementById('share-layout-error')?.classList.add('hidden');
  const user = getCurrentUser();
  const canShare = user?.role === 'designer' || user?.role === 'super';
  document.getElementById('tab-share-layout-share')?.classList.toggle('hidden', !canShare);
  setShareLayoutTab(canShare ? 'share' : 'download');
  document.getElementById('modal-share-layout')?.classList.remove('hidden');
}

// ===== Event wiring =====
function initHomeEvents() {
  document.getElementById('btn-logout').addEventListener('click', () => {
    if (designerInstance) {
      designerInstance.destroy();
      designerInstance = null;
    }
    if (window.AuthStore) window.AuthStore.logout();
    currentSessionPassword = '';
    selectedManagedUser = null;
    showLoginView();
  });

  document.getElementById('btn-add-user').addEventListener('click', () => {
    openAddUserModal();
  });

  document.getElementById('btn-smart-ui')?.addEventListener('click', () => {
    openSmartUiModal();
  });

  document.getElementById('btn-new-layout').addEventListener('click', () => {
    showSetupView(null);
  });

  document.getElementById('btn-import-layout').addEventListener('click', () => {
    document.getElementById('import-layout-file').value = '';
    document.getElementById('import-layout-file').click();
  });

  document.getElementById('import-layout-file').addEventListener('change', (e) => {
    importLayoutFromFile(e.target.files[0]);
  });

  document.getElementById('layout-search-input')?.addEventListener('input', (e) => {
    homeSearchQuery = e.target.value || '';
    renderHomeView();
  });

  document.getElementById('layouts-grid').addEventListener('click', (e) => {
    const btn = e.target.closest('[data-action]');
    if (!btn) return;
    const action = btn.dataset.action;
    const id = btn.dataset.id;
    if (action === 'edit') {
      const user = getCurrentUser();
      if (!(user?.role === 'designer' || user?.role === 'super')) {
        alert('You do not have access to edit layouts.');
        return;
      }
      openDesigner(id);
    } else if (action === 'share') {
      openShareLayoutModal(id);
    } else if (action === 'run') {
      openRunView(id);
    } else if (action === 'delete') {
      if (confirm('Delete this layout? This cannot be undone.')) {
        deleteLayout(id);
        renderHomeView();
      }
    }
  });
}

function initShareLayoutEvents() {
  document.querySelectorAll('.modal-tab[data-share-tab]').forEach(tab => {
    tab.addEventListener('click', () => setShareLayoutTab(tab.dataset.shareTab));
  });

  document.getElementById('btn-share-layout-close')?.addEventListener('click', closeShareLayoutModal);
  document.getElementById('btn-share-layout-cancel')?.addEventListener('click', closeShareLayoutModal);
  document.getElementById('modal-share-layout')?.addEventListener('click', (e) => {
    if (e.target === document.getElementById('modal-share-layout')) closeShareLayoutModal();
  });

  document.getElementById('btn-share-layout-download')?.addEventListener('click', () => {
    if (!shareLayoutModalId) return;
    exportLayout(shareLayoutModalId);
    closeShareLayoutModal();
  });

  document.getElementById('btn-share-layout-send')?.addEventListener('click', async () => {
    if (!shareLayoutModalId) return;
    const userId = document.getElementById('share-layout-user-id')?.value || '';
    const btn = document.getElementById('btn-share-layout-send');
    const error = document.getElementById('share-layout-error');
    btn.disabled = true;
    btn.textContent = 'Sharing...';
    const result = await shareLayoutToUser(shareLayoutModalId, userId);
    btn.disabled = false;
    btn.textContent = 'Share Layout';
    if (!result.ok) {
      error.textContent = result.message || 'Could not share layout.';
      error.classList.remove('hidden');
      return;
    }
    showToast(result.message || 'Layout shared.');
    closeShareLayoutModal();
  });
}

function initLoginEvents() {
  ['login-user-id', 'new-user-id', 'share-layout-user-id'].forEach(id => {
    document.getElementById(id)?.addEventListener('input', (e) => {
      const pos = e.target.selectionStart;
      e.target.value = e.target.value.toUpperCase();
      e.target.setSelectionRange(pos, pos);
    });
  });

  document.getElementById('login-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const username = document.getElementById('login-user-id').value.trim().toUpperCase();
    const password = document.getElementById('login-password').value;
    const error = document.getElementById('login-error');
    const btn = document.getElementById('btn-login');

    error.classList.add('hidden');
    btn.disabled = true;
    btn.textContent = 'Signing in...';

    try {
      await window.AuthStore.login(username, password);
      currentSessionPassword = password;
      document.getElementById('login-password').value = '';
      await loadCurrentUserLayouts();
      const livePreviewId = new URLSearchParams(window.location.search).get('livepreview');
      if (livePreviewId) {
        initLivePreviewMode(livePreviewId);
        return;
      }
      renderHomeView();
    } catch (err) {
      error.textContent = err.message || 'Sign in failed.';
      error.classList.remove('hidden');
    } finally {
      btn.disabled = false;
      btn.textContent = 'Sign In';
    }
  });
}

function openAddUserModal() {
  const current = getCurrentUser();
  if (!current?.isSuperUser) {
    alert('Only the super user can maintain users.');
    return;
  }
  selectedManagedUser = null;
  document.getElementById('users-directory-body').innerHTML = '<tr><td colspan="5" class="hint">Loading users...</td></tr>';
  document.getElementById('new-user-id').value = '';
  document.getElementById('new-user-name').value = '';
  document.getElementById('new-user-role').value = 'user';
  document.getElementById('new-user-active').value = 'true';
  document.getElementById('new-user-password').value = '';
  document.getElementById('users-search-input').value = '';
  managedUsersSearchQuery = '';
  document.getElementById('users-directory-error').classList.add('hidden');
  document.getElementById('add-user-error').classList.add('hidden');
  document.getElementById('change-user-error')?.classList.add('hidden');
  document.getElementById('modal-add-user').classList.remove('hidden');
  refreshUsersDirectory();
}

function closeAddUserModal() {
  document.getElementById('modal-add-user').classList.add('hidden');
  closeUserLayoutsModal();
}

function openCreateUserModal() {
  document.getElementById('add-user-error').classList.add('hidden');
  document.getElementById('modal-create-user').classList.remove('hidden');
}

function closeCreateUserModal() {
  document.getElementById('modal-create-user').classList.add('hidden');
}

function openChangeUserModal(user) {
  if (!user) return;
  document.getElementById('change-user-id').value = user.username || '';
  document.getElementById('change-user-password').value = '';
  document.getElementById('change-user-active').value = user.active ? 'true' : 'false';
  document.getElementById('change-user-error').classList.add('hidden');
  document.getElementById('modal-change-user').classList.remove('hidden');
}

function closeChangeUserModal() {
  document.getElementById('modal-change-user').classList.add('hidden');
}

function closeUserLayoutsModal() {
  document.getElementById('modal-user-layouts')?.classList.add('hidden');
  managedLayoutsUser = null;
}

function renderManagedUserLayouts(layouts) {
  const body = document.getElementById('user-layouts-body');
  if (!body) return;
  const items = Array.isArray(layouts) ? layouts : [];
  if (!items.length) {
    body.innerHTML = '<tr><td colspan="3" class="hint">No layouts found for this user.</td></tr>';
    return;
  }
  body.innerHTML = items.map(layout => `
    <tr data-layout-id="${escapeHtml(layout.id || '')}">
      <td>${escapeHtml(layout.name || 'Untitled Layout')}</td>
      <td>${escapeHtml(formatDate(layout.updatedAt || layout.createdAt || ''))}</td>
      <td>
        <div class="user-layout-actions">
          <button class="btn btn-secondary btn-sm" data-user-layout-action="export" data-layout-id="${escapeHtml(layout.id || '')}">Export</button>
          <button class="btn btn-danger btn-sm" data-user-layout-action="delete" data-layout-id="${escapeHtml(layout.id || '')}">Delete</button>
        </div>
      </td>
    </tr>
  `).join('');
}

async function refreshManagedUserLayouts() {
  const error = document.getElementById('user-layouts-error');
  const btn = document.getElementById('btn-user-layouts-refresh');
  error?.classList.add('hidden');
  if (!managedLayoutsUser?.id) {
    if (error) {
      error.textContent = 'Select a valid user.';
      error.classList.remove('hidden');
    }
    return;
  }
  btn.disabled = true;
  btn.textContent = 'Loading...';
  try {
    const layouts = await window.LayoutStore?.listForUser?.(managedLayoutsUser.id);
    renderManagedUserLayouts(layouts || []);
  } catch (err) {
    if (error) {
      error.textContent = err.message || 'Could not load user layouts.';
      error.classList.remove('hidden');
    }
  } finally {
    btn.disabled = false;
    btn.textContent = 'Refresh';
  }
}

function openUserLayoutsModal() {
  if (!selectedManagedUser?.id) {
    alert('Select a user first.');
    return;
  }
  managedLayoutsUser = selectedManagedUser;
  document.getElementById('user-layouts-title').textContent = `User Layouts - ${selectedManagedUser.username}`;
  document.getElementById('user-layouts-error')?.classList.add('hidden');
  document.getElementById('user-layouts-body').innerHTML = '<tr><td colspan="3" class="hint">Loading layouts...</td></tr>';
  document.getElementById('modal-user-layouts')?.classList.remove('hidden');
  refreshManagedUserLayouts();
}

function formatLastLogin(lastLoginAt) {
  if (!lastLoginAt) return 'Never';
  try {
    return new Date(lastLoginAt).toLocaleString();
  } catch {
    return String(lastLoginAt);
  }
}

function renderUsersDirectory(users) {
  const body = document.getElementById('users-directory-body');
  if (!body) return;
  const q = (managedUsersSearchQuery || '').trim().toLowerCase();
  const allItems = Array.isArray(users) ? users : [];
  const items = q
    ? allItems.filter(user => {
      const hay = `${user.username || ''} ${user.fullName || ''} ${user.role || ''} ${user.active ? 'active' : 'inactive'}`.toLowerCase();
      return hay.includes(q);
    })
    : allItems;
  if (!items.length) {
    body.innerHTML = `<tr><td colspan="5" class="hint">${q ? 'No users match your search.' : 'No users found.'}</td></tr>`;
    selectedManagedUser = null;
    return;
  }
  body.innerHTML = items.map(user => `
    <tr data-user-id="${escapeHtml(user.id || '')}" class="${selectedManagedUser?.id === user.id ? 'selected' : ''}">
      <td>${escapeHtml(user.username || '')}</td>
      <td>${escapeHtml(user.fullName || '-')}</td>
      <td>${escapeHtml(smartLabel(user.role || 'user'))}</td>
      <td>${escapeHtml(formatLastLogin(user.lastLoginAt))}</td>
      <td><span class="user-status-pill ${user.active ? 'active' : 'inactive'}">${user.active ? 'Active' : 'Inactive'}</span></td>
    </tr>
  `).join('');
}

async function refreshUsersDirectory() {
  const current = getCurrentUser();
  const error = document.getElementById('users-directory-error');
  const btn = document.getElementById('btn-list-users');
  error.classList.add('hidden');

  if (!current?.isSuperUser) {
    error.textContent = 'Only the super user can view users.';
    error.classList.remove('hidden');
    return;
  }
  if (!currentSessionPassword) {
    error.textContent = 'Session expired. Please sign in again.';
    error.classList.remove('hidden');
    return;
  }

  btn.disabled = true;
  btn.textContent = 'Loading...';
  try {
    const users = await window.AuthStore.listUsers(current.username, currentSessionPassword);
    users.sort((a, b) => String(a.username || '').localeCompare(String(b.username || '')));
    managedUsersCache = users;
    if (selectedManagedUser?.id) {
      selectedManagedUser = users.find(u => u.id === selectedManagedUser.id) || null;
    }
    renderUsersDirectory(users);
  } catch (err) {
    error.textContent = err.message || 'Could not load users.';
    error.classList.remove('hidden');
  } finally {
    btn.disabled = false;
    btn.textContent = 'Refresh';
  }
}

function initAddUserEvents() {
  document.getElementById('btn-add-user-close').addEventListener('click', closeAddUserModal);
  document.getElementById('btn-add-user-cancel').addEventListener('click', closeAddUserModal);
  document.getElementById('modal-add-user').addEventListener('click', (e) => {
    if (e.target === document.getElementById('modal-add-user')) closeAddUserModal();
  });
  document.getElementById('btn-user-add').addEventListener('click', openCreateUserModal);
  document.getElementById('btn-list-users').addEventListener('click', refreshUsersDirectory);
  document.getElementById('users-search-input')?.addEventListener('input', (e) => {
    managedUsersSearchQuery = e.target.value || '';
    renderUsersDirectory(managedUsersCache);
  });
  document.getElementById('btn-user-layouts').addEventListener('click', openUserLayoutsModal);

  document.getElementById('btn-create-user-close').addEventListener('click', closeCreateUserModal);
  document.getElementById('btn-create-user-cancel').addEventListener('click', closeCreateUserModal);
  document.getElementById('modal-create-user').addEventListener('click', (e) => {
    if (e.target === document.getElementById('modal-create-user')) closeCreateUserModal();
  });
  document.getElementById('users-directory-body').addEventListener('click', (e) => {
    const row = e.target.closest('tr[data-user-id]');
    if (!row) return;
    const id = row.dataset.userId;
    const rows = Array.from(document.querySelectorAll('#users-directory-body tr[data-user-id]'));
    rows.forEach(r => r.classList.toggle('selected', r === row));
    selectedManagedUser = managedUsersCache.find(u => u.id === id) || {
      id,
      username: row.children[0]?.textContent?.trim() || '',
      active: (row.children[4]?.textContent || '').trim().toLowerCase() === 'active',
    };
  });

  document.getElementById('btn-user-change').addEventListener('click', () => {
    if (!selectedManagedUser?.username) {
      alert('Select a user first.');
      return;
    }
    openChangeUserModal(selectedManagedUser);
  });

  document.getElementById('btn-user-delete').addEventListener('click', async () => {
    const current = getCurrentUser();
    const error = document.getElementById('users-directory-error');
    error.classList.add('hidden');

    if (!selectedManagedUser?.username) {
      alert('Select a user first.');
      return;
    }
    if (!currentSessionPassword) {
      error.textContent = 'Session expired. Please sign in again.';
      error.classList.remove('hidden');
      return;
    }
    if (selectedManagedUser.username === current?.username) {
      alert('You cannot delete your own user id.');
      return;
    }
    if (!confirm(`Delete user "${selectedManagedUser.username}"?`)) return;

    try {
      await window.AuthStore.deleteUser(current.username, currentSessionPassword, selectedManagedUser.username);
      showToast(`User "${selectedManagedUser.username}" deleted.`);
      selectedManagedUser = null;
      await refreshUsersDirectory();
    } catch (err) {
      error.textContent = err.message || 'Could not delete user.';
      error.classList.remove('hidden');
    }
  });

  document.getElementById('btn-add-user-save').addEventListener('click', async () => {
    const current = getCurrentUser();
    const username = document.getElementById('new-user-id').value.trim().toUpperCase();
    const fullName = document.getElementById('new-user-name').value.trim();
    const password = document.getElementById('new-user-password').value;
    const role = document.getElementById('new-user-role').value || 'user';
    const error = document.getElementById('add-user-error');
    const btn = document.getElementById('btn-add-user-save');

    error.classList.add('hidden');

    if (!current?.isSuperUser) {
      error.textContent = 'Only the super user can add users.';
      error.classList.remove('hidden');
      return;
    }
    if (!currentSessionPassword) {
      error.textContent = 'Session expired. Please sign in again.';
      error.classList.remove('hidden');
      return;
    }
    if (!username || !fullName || !password) {
      error.textContent = 'Enter user id, user name, and password.';
      error.classList.remove('hidden');
      return;
    }

    btn.disabled = true;
    btn.textContent = 'Creating...';
    try {
      await window.AuthStore.addUser(current.username, currentSessionPassword, username, fullName, password, role);
      if (document.getElementById('new-user-active').value === 'false') {
        await window.AuthStore.setUserActive(current.username, currentSessionPassword, username, false);
      }
      document.getElementById('new-user-id').value = '';
      document.getElementById('new-user-name').value = '';
      document.getElementById('new-user-password').value = '';
      showToast(`User "${username}" created.`);
      closeCreateUserModal();
      await refreshUsersDirectory();
    } catch (err) {
      error.textContent = err.message || 'Could not create user.';
      error.classList.remove('hidden');
    } finally {
      btn.disabled = false;
      btn.textContent = 'Create User';
    }
  });

  document.getElementById('btn-change-user-close').addEventListener('click', closeChangeUserModal);
  document.getElementById('btn-change-user-cancel').addEventListener('click', closeChangeUserModal);
  document.getElementById('modal-change-user').addEventListener('click', (e) => {
    if (e.target === document.getElementById('modal-change-user')) closeChangeUserModal();
  });

  document.getElementById('btn-change-user-save').addEventListener('click', async () => {
    const current = getCurrentUser();
    const username = document.getElementById('change-user-id').value.trim().toUpperCase();
    const newPassword = document.getElementById('change-user-password').value;
    const active = document.getElementById('change-user-active').value === 'true';
    const error = document.getElementById('change-user-error');
    const btn = document.getElementById('btn-change-user-save');

    error.classList.add('hidden');

    if (!current?.isSuperUser) {
      error.textContent = 'Only the super user can change passwords.';
      error.classList.remove('hidden');
      return;
    }
    if (!currentSessionPassword) {
      error.textContent = 'Session expired. Please sign in again.';
      error.classList.remove('hidden');
      return;
    }
    const statusChanged = selectedManagedUser && (Boolean(selectedManagedUser.active) !== active);
    if (!username) {
      error.textContent = 'Invalid user selected.';
      error.classList.remove('hidden');
      return;
    }
    if (!newPassword && !statusChanged) {
      error.textContent = 'Change password or update status.';
      error.classList.remove('hidden');
      return;
    }

    btn.disabled = true;
    btn.textContent = 'Saving...';
    try {
      if (newPassword) {
        await window.AuthStore.resetPassword(current.username, currentSessionPassword, username, newPassword);
      }
      if (statusChanged) {
        await window.AuthStore.setUserActive(current.username, currentSessionPassword, username, active);
      }
      document.getElementById('change-user-password').value = '';
      showToast(`User "${username}" updated.`);
      closeChangeUserModal();
      await refreshUsersDirectory();
    } catch (err) {
      error.textContent = err.message || 'Could not reset password.';
      error.classList.remove('hidden');
    } finally {
      btn.disabled = false;
      btn.textContent = 'Save';
    }
  });

  document.getElementById('btn-user-layouts-close')?.addEventListener('click', closeUserLayoutsModal);
  document.getElementById('btn-user-layouts-cancel')?.addEventListener('click', closeUserLayoutsModal);
  document.getElementById('btn-user-layouts-refresh')?.addEventListener('click', refreshManagedUserLayouts);
  document.getElementById('modal-user-layouts')?.addEventListener('click', (e) => {
    if (e.target === document.getElementById('modal-user-layouts')) closeUserLayoutsModal();
  });
  document.getElementById('user-layouts-body')?.addEventListener('click', async (e) => {
    const btn = e.target.closest('[data-user-layout-action]');
    if (!btn || !managedLayoutsUser?.id) return;
    const action = btn.dataset.userLayoutAction;
    const layoutId = btn.dataset.layoutId;
    const error = document.getElementById('user-layouts-error');
    error?.classList.add('hidden');
    try {
      const layouts = await window.LayoutStore?.listForUser?.(managedLayoutsUser.id);
      const layout = (layouts || []).find(l => String(l.id) === String(layoutId));
      if (!layout) throw new Error('Layout not found.');

      if (action === 'export') {
        exportLayoutData(layout);
        return;
      }
      if (action === 'delete') {
        if (!confirm(`Delete layout "${layout.name}" for user "${managedLayoutsUser.username}"?`)) return;
        await window.LayoutStore?.removeForUser?.(layout.id, managedLayoutsUser.id);
        showToast(`Deleted "${layout.name}".`);
        await refreshManagedUserLayouts();
      }
    } catch (err) {
      if (error) {
        error.textContent = err.message || 'Could not complete action.';
        error.classList.remove('hidden');
      }
    }
  });
}

function openSmartUiModal() {
  const current = getCurrentUser();
  if (!current?.isSuperUser) {
    alert('Only the super user can maintain Smart UI.');
    return;
  }
  smartRulesDraft = JSON.parse(JSON.stringify(getSmartRules()));
  document.getElementById('smart-ui-error')?.classList.add('hidden');
  setSmartUiTab('header');
  renderSmartRulesEditor();
  document.getElementById('modal-smart-ui')?.classList.remove('hidden');
}

function closeSmartUiModal() {
  document.getElementById('modal-smart-ui')?.classList.add('hidden');
}

function setSmartUiTab(tabName) {
  document.querySelectorAll('.modal-tab[data-smart-tab]').forEach(tab => {
    tab.classList.toggle('active', tab.dataset.smartTab === tabName);
  });
  document.querySelectorAll('.smart-tab-content').forEach(content => content.classList.remove('active'));
  document.getElementById('smart-tab-' + tabName)?.classList.add('active');
}

function renderSmartRulesEditor() {
  if (!smartRulesDraft) smartRulesDraft = JSON.parse(JSON.stringify(getSmartRules()));
  ['header', 'table', 'footer'].forEach(bucket => {
    const list = document.getElementById(`smart-${bucket}-list`);
    if (!list) return;
    const values = smartRulesDraft[bucket] || [];
    list.innerHTML = values.length
      ? values.map(value => `
        <span class="smart-rule-chip">
          ${escapeHtml(value)}
            <button type="button" title="Change" data-smart-edit="${bucket}" data-smart-value="${escapeHtml(value)}">Edit</button>
            <button type="button" data-smart-remove="${bucket}" data-smart-value="${escapeHtml(value)}">&times;</button>
        </span>
        `).join('')
      : '<span class="hint">No keywords added.</span>';
  });
}

function addSmartRule(bucket) {
  const input = document.getElementById(`smart-${bucket}-input`);
  const value = input?.value.trim().toLowerCase();
  if (!value) return;
  smartRulesDraft = smartRulesDraft || JSON.parse(JSON.stringify(getSmartRules()));
  smartRulesDraft[bucket] = smartRulesDraft[bucket] || [];
  if (!smartRulesDraft[bucket].includes(value)) smartRulesDraft[bucket].push(value);
  smartRulesDraft[bucket].sort((a, b) => a.localeCompare(b));
  input.value = '';
  renderSmartRulesEditor();
}

async function saveSmartUiRules() {
  const current = getCurrentUser();
  const error = document.getElementById('smart-ui-error');
  const btn = document.getElementById('btn-smart-ui-save');
  error?.classList.add('hidden');

  if (!current?.isSuperUser) {
    if (error) {
      error.textContent = 'Only the super user can save Smart UI.';
      error.classList.remove('hidden');
    }
    return;
  }

  btn.disabled = true;
  btn.textContent = 'Saving...';
  try {
    const result = await window.LayoutStore?.saveSmartRules?.(smartRulesDraft || getSmartRules());
    renderTemplatePreview(document.getElementById('layout-template')?.value || 'blank');
    closeSmartUiModal();
    showToast(result?.mode === 'local' || result?.ok === false ? 'Smart UI saved locally.' : 'Smart UI saved.');
  } catch (err) {
    if (error) {
      error.textContent = err.message || 'Could not save Smart UI.';
      error.classList.remove('hidden');
    }
  } finally {
    btn.disabled = false;
    btn.textContent = 'Save Smart UI';
  }
}

function initSmartUiEvents() {
  document.querySelectorAll('.modal-tab[data-smart-tab]').forEach(tab => {
    tab.addEventListener('click', () => setSmartUiTab(tab.dataset.smartTab));
  });
  ['header', 'table', 'footer'].forEach(bucket => {
    document.getElementById(`btn-smart-${bucket}-add`)?.addEventListener('click', () => addSmartRule(bucket));
    document.getElementById(`smart-${bucket}-input`)?.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        addSmartRule(bucket);
      }
    });
  });
  document.getElementById('modal-smart-ui')?.addEventListener('click', (e) => {
    if (e.target === document.getElementById('modal-smart-ui')) closeSmartUiModal();
    const editBtn = e.target.closest('[data-smart-edit]');
    if (editBtn) {
      const bucket = editBtn.dataset.smartEdit;
      const oldValue = editBtn.dataset.smartValue;
      const nextValue = prompt('Change keyword', oldValue)?.trim().toLowerCase();
      if (!nextValue) return;
      smartRulesDraft = smartRulesDraft || JSON.parse(JSON.stringify(getSmartRules()));
      smartRulesDraft[bucket] = (smartRulesDraft[bucket] || [])
        .map(item => item === oldValue ? nextValue : item)
        .filter((item, index, arr) => item && arr.indexOf(item) === index)
        .sort((a, b) => a.localeCompare(b));
      renderSmartRulesEditor();
      return;
    }
    const removeBtn = e.target.closest('[data-smart-remove]');
    if (!removeBtn) return;
    const bucket = removeBtn.dataset.smartRemove;
    const value = removeBtn.dataset.smartValue;
    smartRulesDraft = smartRulesDraft || JSON.parse(JSON.stringify(getSmartRules()));
    smartRulesDraft[bucket] = (smartRulesDraft[bucket] || []).filter(item => item !== value);
    renderSmartRulesEditor();
  });
  document.getElementById('btn-smart-ui-close')?.addEventListener('click', closeSmartUiModal);
  document.getElementById('btn-smart-ui-cancel')?.addEventListener('click', closeSmartUiModal);
  document.getElementById('btn-smart-ui-save')?.addEventListener('click', saveSmartUiRules);
}

function initSetupEvents() {
  document.getElementById('page-size')?.addEventListener('change', toggleSetupCustomSize);
  ['custom-page-width', 'custom-page-height'].forEach(id => {
    document.getElementById(id)?.addEventListener('input', () => {
      if (document.getElementById('page-size')?.value !== 'custom') return;
      renderTemplatePreview(document.getElementById('layout-template')?.value || 'blank');
    });
  });

  document.getElementById('btn-setup-back').addEventListener('click', () => {
    renderHomeView();
  });

  document.getElementById('btn-step1-next').addEventListener('click', () => {
    const name = document.getElementById('layout-name').value.trim();
    if (!name) {
      alert('Please enter a layout name.');
      document.getElementById('layout-name').focus();
      return;
    }
    showSetupStep(2);
  });

  document.getElementById('btn-step2-back').addEventListener('click', () => {
    showSetupStep(1);
  });

  document.getElementById('btn-step2-next').addEventListener('click', () => {
    showSetupStep(3);
  });

  document.getElementById('btn-step3-back').addEventListener('click', () => {
    showSetupStep(2);
  });

  document.getElementById('btn-create-layout').addEventListener('click', () => {
    createNewLayout();
  });

  document.getElementById('field-names-input').addEventListener('input', () => {
    const text = document.getElementById('field-names-input').value;
    const fields = parseFieldNames(text);
    const preview = document.getElementById('parsed-fields-preview');
    if (fields.length > 0) {
      preview.classList.remove('hidden');
      preview.innerHTML = `<label>Parsed Fields:</label><div class="parsed-fields-list">${fields.map(f => `<span>${escapeHtml(f)}</span>`).join('')}</div>`;
      renderTemplatePreview(document.getElementById('layout-template')?.value || 'blank');
    } else {
      preview.classList.add('hidden');
      preview.innerHTML = '';
      renderTemplatePreview(document.getElementById('layout-template')?.value || 'blank');
    }
  });
}

function initDesignerAppEvents() {
  document.getElementById('btn-designer-home').addEventListener('click', () => {
    if (designerInstance) {
      designerInstance.saveLayout();
      designerInstance.destroy();
      designerInstance = null;
    }
    renderHomeView();
  });

  document.getElementById('btn-save').addEventListener('click', () => {
    if (designerInstance) {
      designerInstance.saveLayout();
      showToast('Layout saved!');
    }
  });

  document.getElementById('btn-preview-pdf').addEventListener('click', async () => {
    if (!designerInstance) return;
    const layout = designerInstance.getLayout();
    // For preview, use empty field values
    const fieldValues = {};
    (layout.fields || []).forEach(f => { fieldValues[f] = `{${f}}`; });
    await generatePDF(layout, fieldValues);
  });

  document.getElementById('btn-run-preview').addEventListener('click', () => {
    if (!designerInstance) return;
    openRunPreviewModal(designerInstance.getLayout());
  });

  document.getElementById('btn-export-layout').addEventListener('click', () => {
    if (!designerInstance) return;
    designerInstance.saveLayout();
    exportLayout(currentLayoutId);
  });

  document.getElementById('btn-live-preview').addEventListener('click', () => {
    if (!designerInstance) return;
    designerInstance.saveLayout();
    window.open(window.location.pathname + '?livepreview=' + currentLayoutId, '_blank');
  });

  document.getElementById('btn-undo').addEventListener('click', () => {
    if (designerInstance) designerInstance.undo();
  });

  document.getElementById('btn-redo').addEventListener('click', () => {
    if (designerInstance) designerInstance.redo();
  });

  document.getElementById('btn-zoom-in').addEventListener('click', () => {
    if (designerInstance) designerInstance.zoomIn();
  });

  document.getElementById('btn-zoom-out').addEventListener('click', () => {
    if (designerInstance) designerInstance.zoomOut();
  });

  document.getElementById('btn-grid-toggle').addEventListener('click', () => {
    if (designerInstance) designerInstance.toggleGrid();
  });

  document.getElementById('btn-page-settings').addEventListener('click', () => {
    if (designerInstance) designerInstance.openPageSettings();
  });

  // Page settings modal
  document.getElementById('btn-page-settings-close').addEventListener('click', closePageSettingsModal);
  document.getElementById('btn-page-settings-cancel').addEventListener('click', closePageSettingsModal);
  document.getElementById('modal-page-size')?.addEventListener('change', (e) => {
    document.getElementById('modal-custom-size-group')?.classList.toggle('hidden', e.target.value !== 'custom');
  });
  document.getElementById('btn-page-settings-apply').addEventListener('click', () => {
    if (designerInstance) designerInstance.applyPageSettings();
  });

  // Keyboard shortcuts
  document.addEventListener('keydown', (e) => {
    if (currentView !== 'designer' || !designerInstance) return;
    const ctrl = e.ctrlKey || e.metaKey;
    if (ctrl && e.key === 'z' && !e.shiftKey) { e.preventDefault(); designerInstance.undo(); }
    if (ctrl && (e.key === 'y' || (e.key === 'z' && e.shiftKey))) { e.preventDefault(); designerInstance.redo(); }
    if (ctrl && e.key === 's') { e.preventDefault(); designerInstance.saveLayout(); showToast('Layout saved!'); }
    if (e.key === 'Delete' || e.key === 'Backspace') {
      const active = document.activeElement;
      const isEditing = active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA' || active.tagName === 'SELECT' || active.isContentEditable);
      if (!isEditing) { e.preventDefault(); designerInstance.deleteSelected(); }
    }
    if (ctrl && e.key === 'd') { e.preventDefault(); designerInstance.duplicateSelected(); }

    // Arrow key nudge
    if (['ArrowLeft','ArrowRight','ArrowUp','ArrowDown'].includes(e.key)) {
      const active = document.activeElement;
      const isEditing = active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA' || active.tagName === 'SELECT' || active.isContentEditable);
      if (!isEditing && designerInstance.selectedId) {
        e.preventDefault();
        const step = e.shiftKey ? 0.1 : (ctrl ? 5 : 1); // Shift=fine (0.1mm), Ctrl=coarse (5mm), default=1mm
        const el = designerInstance._findElement(designerInstance.selectedId);
        if (el) {
          if (e.key === 'ArrowLeft')  el.x = Math.max(0, el.x - step);
          if (e.key === 'ArrowRight') el.x = Math.max(0, el.x + step);
          if (e.key === 'ArrowUp')    el.y = Math.max(0, el.y - step);
          if (e.key === 'ArrowDown')  el.y = Math.max(0, el.y + step);
          const domEl = designerInstance.pageCanvas.querySelector(`[data-id="${designerInstance.selectedId}"]`);
          if (domEl) {
            domEl.style.left = designerInstance._mmToPx(el.x) + 'px';
            domEl.style.top  = designerInstance._mmToPx(el.y) + 'px';
          }
          const propX = document.getElementById('prop-x');
          const propY = document.getElementById('prop-y');
          if (propX) propX.value = el.x.toFixed(1);
          if (propY) propY.value = el.y.toFixed(1);
          designerInstance._debounceSave();
        }
      }
    }
  });
}

function closePageSettingsModal() {
  document.getElementById('modal-page-settings').classList.add('hidden');
}

// ===== Run Preview Modal =====
let runPreviewModalParsedData = null;

function openRunPreviewModal(layout) {
  const modal = document.getElementById('modal-run-preview');
  const fields = layout.fields || [];

  // Populate manual entry form
  const form = document.getElementById('modal-run-fields-form');
  form.innerHTML = '';
  if (fields.length === 0) {
    form.innerHTML = '<p class="hint">No fields defined for this layout.</p>';
  } else {
    fields.forEach(f => {
      const row = document.createElement('div');
      row.className = 'field-input-row';
      row.innerHTML = `<label class="field-input-label">${f}</label><input type="text" class="field-input-val" data-field="${f}" placeholder="${f}" />`;
      form.appendChild(row);
    });
  }

  // Reset paste state
  runPreviewModalParsedData = null;
  document.getElementById('modal-run-paste-area').value = '';
  document.getElementById('modal-paste-status').textContent = '';
  document.getElementById('modal-run-preview-area').innerHTML =
    '<div class="preview-placeholder"><div style="font-size:36px;margin-bottom:12px;">&#128196;</div><div>Enter data and click <strong>Preview</strong></div></div>';

  modal.classList.remove('hidden');
}

function closeRunPreviewModal() {
  document.getElementById('modal-run-preview').classList.add('hidden');
}

function _getRunPreviewData(layout) {
  const activeTab = document.querySelector('#modal-run-tabs .run-tab.active')?.dataset?.mtab || 'paste';
  let fieldValues = {};
  let detailRows = [];

  if (activeTab === 'paste') {
    if (!runPreviewModalParsedData) {
      const text = document.getElementById('modal-run-paste-area').value;
      if (text.trim()) {
        runPreviewModalParsedData = parsePasteData(text, layout.fields || []);
      }
    }
    if (runPreviewModalParsedData && runPreviewModalParsedData.length > 0) {
      fieldValues = { ...runPreviewModalParsedData[0] };
      detailRows = runPreviewModalParsedData;
    }
  } else {
    document.querySelectorAll('#modal-run-fields-form input[data-field]').forEach(input => {
      fieldValues[input.dataset.field] = input.value;
    });
    detailRows = [fieldValues];
  }

  return { fieldValues, detailRows };
}

function initRunPreviewModalEvents() {
  // Tab switching
  document.querySelectorAll('#modal-run-tabs .run-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('#modal-run-tabs .run-tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.modal-run-tab-content').forEach(c => c.classList.remove('active'));
      tab.classList.add('active');
      const el = document.getElementById('modal-run-tab-' + tab.dataset.mtab);
      if (el) el.classList.add('active');
    });
  });

  document.getElementById('btn-run-preview-close').addEventListener('click', closeRunPreviewModal);

  // Close on overlay click
  document.getElementById('modal-run-preview').addEventListener('click', (e) => {
    if (e.target === document.getElementById('modal-run-preview')) closeRunPreviewModal();
  });

  document.getElementById('modal-btn-parse-paste').addEventListener('click', () => {
    if (!designerInstance) return;
    const layout = designerInstance.getLayout();
    const text = document.getElementById('modal-run-paste-area').value;
    if (!text.trim()) {
      document.getElementById('modal-paste-status').textContent = 'Nothing to parse.';
      return;
    }
    runPreviewModalParsedData = parsePasteData(text, layout.fields || []);
    if (runPreviewModalParsedData && runPreviewModalParsedData.length > 0) {
      document.getElementById('modal-paste-status').textContent = '\u2713 ' + runPreviewModalParsedData.length + ' row(s) ready';
    } else {
      document.getElementById('modal-paste-status').textContent = 'No data found.';
    }
  });

  document.getElementById('modal-btn-preview').addEventListener('click', async () => {
    if (!designerInstance) return;
    const layout = designerInstance.getLayout();
    const { fieldValues, detailRows } = _getRunPreviewData(layout);
    const btn = document.getElementById('modal-btn-preview');
    btn.disabled = true;
    btn.textContent = 'Rendering\u2026';
    try {
      const container = document.getElementById('modal-run-preview-area');
      await renderLayoutPreview(layout, fieldValues, detailRows, container);
    } catch (err) {
      console.error('Preview error:', err);
    }
    btn.disabled = false;
    btn.innerHTML = '&#9654; Preview';
  });

  document.getElementById('modal-btn-gen-pdf').addEventListener('click', async () => {
    if (!designerInstance) return;
    const layout = designerInstance.getLayout();
    const { fieldValues, detailRows } = _getRunPreviewData(layout);
    const btn = document.getElementById('modal-btn-gen-pdf');
    btn.disabled = true;
    btn.textContent = 'Generating\u2026';
    try {
      await generatePDF(layout, fieldValues, detailRows);
    } catch (err) {
      console.error('PDF error:', err);
      alert('PDF generation failed: ' + err.message);
    }
    btn.disabled = false;
    btn.innerHTML = '&#128196; PDF';
  });
}

function initRunEvents() {
  // Tab switching
  document.querySelectorAll('.run-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.run-tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.run-tab-content').forEach(c => c.classList.remove('active'));
      tab.classList.add('active');
      const tabId = 'run-tab-' + tab.dataset.tab;
      const el = document.getElementById(tabId);
      if (el) el.classList.add('active');
      document.getElementById('btn-add-manual-row')?.classList.toggle('hidden', tab.dataset.tab !== 'manual');
    });
  });

  document.getElementById('btn-add-manual-row')?.addEventListener('click', () => {
    const layout = getLayoutById(currentLayoutId);
    if (!layout) return;
    manualRowCount += 1;
    renderManualRows(layout);
  });

  document.getElementById('btn-run-back').addEventListener('click', () => {
    renderHomeView();
  });

  document.getElementById('btn-parse-paste').addEventListener('click', () => {
    const layout = getLayoutById(currentLayoutId);
    if (!layout) return;
    const text = document.getElementById('run-paste-area').value;
    const fields = layout.fields || [];
    if (!text.trim()) {
      document.getElementById('paste-status').textContent = 'Nothing to parse.';
      return;
    }
    runParsedData = parsePasteData(text, fields);
    if (runParsedData && runParsedData.length > 0) {
      document.getElementById('paste-status').textContent = '\u2713 ' + runParsedData.length + ' row(s) ready';
      renderParsedPreview(runParsedData, fields);
    } else {
      document.getElementById('paste-status').textContent = 'No data found.';
    }
  });

  document.getElementById('btn-generate-pdf').addEventListener('click', async () => {
    const layoutId = currentLayoutId;
    const layout = getLayoutById(layoutId);
    if (!layout) { alert('Layout not found.'); return; }

    const activeTab = document.querySelector('.run-tab.active')?.dataset?.tab || 'paste';
    let fieldValues = {};
    let detailRows = [];

    if (activeTab === 'paste') {
      // Auto-parse if not done yet
      if (!runParsedData) {
        const text = document.getElementById('run-paste-area').value;
        if (text.trim()) {
          runParsedData = parsePasteData(text, layout.fields || []);
        }
      }
      if (runParsedData && runParsedData.length > 0) {
        fieldValues = { ...runParsedData[0] };
        detailRows = runParsedData;
      }
    } else {
      // Manual mode
      const rows = [];
      document.querySelectorAll('#run-fields-form input[data-field]').forEach(input => {
        const rowIndex = parseInt(input.dataset.row || '0', 10);
        rows[rowIndex] = rows[rowIndex] || {};
        rows[rowIndex][input.dataset.field] = input.value;
      });
      detailRows = rows.filter(row => row && Object.values(row).some(v => v !== ''));
      fieldValues = detailRows[0] || {};
    }

    const btn = document.getElementById('btn-generate-pdf');
    btn.disabled = true;
    btn.textContent = 'Generating\u2026';
    try {
      await generatePDF(layout, fieldValues, detailRows);
    } catch (err) {
      console.error('PDF generation error:', err);
      alert('PDF generation failed: ' + err.message);
    }
    btn.disabled = false;
    btn.innerHTML = '&#128196; Generate PDF';
  });
}

// ===== Toast notification =====
function showToast(message, duration = 2500) {
  let toast = document.getElementById('app-toast');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'app-toast';
    toast.style.cssText = `
      position: fixed; bottom: 24px; left: 50%; transform: translateX(-50%);
      background: #43e97b; color: #1e1e2e; padding: 10px 24px;
      border-radius: 24px; font-weight: 700; font-size: 13px;
      z-index: 99999; opacity: 0; transition: opacity 0.2s;
      pointer-events: none; white-space: nowrap;
    `;
    document.body.appendChild(toast);
  }
  toast.textContent = message;
  toast.style.opacity = '1';
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => { toast.style.opacity = '0'; }, duration);
}

window.showToast = showToast;
window.exportLayout = exportLayout;
window.importLayoutFromFile = importLayoutFromFile;
window.PAGE_SIZES = PAGE_SIZES;
window.loadLayouts = loadLayouts;
window.saveLayout = saveLayout;
window.deleteLayout = deleteLayout;
window.getLayoutById = getLayoutById;
window.parseFieldNames = parseFieldNames;
window.escapeHtml = escapeHtml;
window.closePageSettingsModal = closePageSettingsModal;
window.openRunView = openRunView;
window.parsePasteData = parsePasteData;
window.renderParsedPreview = renderParsedPreview;

// ===== Live Preview Mode =====
function initLivePreviewMode(layoutId) {
  // Hide all normal views and show the live preview view full-screen
  document.querySelectorAll('.view').forEach(v => { v.classList.add('hidden'); v.classList.remove('active'); });
  const view = document.getElementById('view-livepreview');
  view.classList.remove('hidden');
  view.classList.add('active');

  let lpLayout = getLayoutById(layoutId);
  let lpParsedData = null;
  let lpLastRendered = false; // whether a render has happened

  function setLayoutName(layout) {
    document.getElementById('livepreview-name').textContent = layout ? layout.name : 'Live Preview';
  }

  function populateFields(layout) {
    const form = document.getElementById('livepreview-fields-form');
    form.innerHTML = '';
    (layout.fields || []).forEach(f => {
      const row = document.createElement('div');
      row.className = 'field-input-row';
      row.innerHTML = `<label class="field-input-label">${escapeHtml(f)}</label><input type="text" class="field-input-val" data-field="${escapeHtml(f)}" placeholder="${escapeHtml(f)}" />`;
      form.appendChild(row);
    });
  }

  function getLpData() {
    const activeTab = document.querySelector('#livepreview-tabs .run-tab.active')?.dataset?.ltab || 'paste';
    let fieldValues = {}, detailRows = [];
    if (activeTab === 'paste') {
      if (!lpParsedData) {
        const text = document.getElementById('livepreview-paste-area').value;
        if (text.trim()) lpParsedData = parsePasteData(text, lpLayout.fields || []);
      }
      if (lpParsedData && lpParsedData.length > 0) {
        fieldValues = { ...lpParsedData[0] };
        detailRows = lpParsedData;
      }
    } else {
      document.querySelectorAll('#livepreview-fields-form input[data-field]').forEach(inp => {
        fieldValues[inp.dataset.field] = inp.value;
      });
      detailRows = [fieldValues];
    }
    return { fieldValues, detailRows };
  }

  async function doRender() {
    lpLayout = getLayoutById(layoutId);
    if (!lpLayout) return;
    const { fieldValues, detailRows } = getLpData();
    const area = document.getElementById('livepreview-preview-area');
    const btn = document.getElementById('livepreview-btn-render');
    btn.disabled = true;
    btn.textContent = 'Rendering\u2026';
    try {
      await renderLayoutPreview(lpLayout, fieldValues, detailRows, area);
      lpLastRendered = true;
      document.getElementById('livepreview-placeholder').style.display = 'none';
    } catch (err) {
      console.error('Live preview render error:', err);
    }
    btn.disabled = false;
    btn.innerHTML = '&#9654; Render';
  }

  // Initial setup
  if (!lpLayout) {
    document.getElementById('livepreview-name').textContent = 'Layout not found';
    return;
  }
  setLayoutName(lpLayout);
  populateFields(lpLayout);

  // Tab switching
  document.querySelectorAll('#livepreview-tabs .run-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('#livepreview-tabs .run-tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('#livepreview-tab-paste, #livepreview-tab-manual').forEach(c => c.classList.remove('active'));
      tab.classList.add('active');
      const el = document.getElementById('livepreview-tab-' + tab.dataset.ltab);
      if (el) el.classList.add('active');
    });
  });

  document.getElementById('livepreview-btn-parse').addEventListener('click', () => {
    const text = document.getElementById('livepreview-paste-area').value;
    if (!text.trim()) { document.getElementById('livepreview-paste-status').textContent = 'Nothing to parse.'; return; }
    lpParsedData = parsePasteData(text, lpLayout.fields || []);
    document.getElementById('livepreview-paste-status').textContent =
      lpParsedData && lpParsedData.length > 0 ? '\u2713 ' + lpParsedData.length + ' row(s) ready' : 'No data found.';
  });

  document.getElementById('livepreview-btn-render').addEventListener('click', () => doRender());

  document.getElementById('livepreview-btn-pdf').addEventListener('click', async () => {
    lpLayout = getLayoutById(layoutId);
    if (!lpLayout) return;
    const { fieldValues, detailRows } = getLpData();
    const btn = document.getElementById('livepreview-btn-pdf');
    btn.disabled = true;
    btn.textContent = 'Generating\u2026';
    try {
      await generatePDF(lpLayout, fieldValues, detailRows);
    } catch (err) {
      alert('PDF error: ' + err.message);
    }
    btn.disabled = false;
    btn.innerHTML = '&#128196; PDF';
  });

  function refreshLivePreviewFromSavedLayout() {
    if (window.LayoutStore) window.LayoutStore.syncFromLocal();
    const updated = getLayoutById(layoutId);
    if (!updated) return;
    lpLayout = updated;
    // Show the "updated" badge
    const badge = document.getElementById('livepreview-update-badge');
    badge.style.display = '';
    setTimeout(() => { badge.style.display = 'none'; }, 3000);
    // Auto-refresh if a previous render exists and checkbox is on
    if (lpLastRendered && document.getElementById('livepreview-auto-refresh').checked) {
      doRender();
    }
  }

  // Listen for layout changes saved by the designer in another tab.
  window.addEventListener('storage', (e) => {
    if (e.key !== STORAGE_KEY) return;
    refreshLivePreviewFromSavedLayout();
  });

  if (window.LayoutStore) {
    window.LayoutStore.onExternalChange((message) => {
      if (!message || message.id !== layoutId) return;
      refreshLivePreviewFromSavedLayout();
    });
  }
}

// ===== Init =====
document.addEventListener('DOMContentLoaded', async () => {
  initLoginEvents();
  initHomeEvents();
  initSetupEvents();
  initDesignerAppEvents();
  initRunEvents();
  initRunPreviewModalEvents();
  initAddUserEvents();
  initSmartUiEvents();
  initShareLayoutEvents();

  // Check if this tab is opened as a live preview
  const urlParams = new URLSearchParams(window.location.search);
  const livePreviewId = urlParams.get('livepreview');
  if (livePreviewId) {
    const ready = await loadCurrentUserLayouts();
    if (!ready) return;
    // Wait for scripts to load then init live preview
    initLivePreviewMode(livePreviewId);
    return;
  }

  const ready = await loadCurrentUserLayouts();
  if (!ready) return;
  renderHomeView();
});
