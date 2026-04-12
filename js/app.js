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

// ===== State =====
let currentView = 'home';
let currentLayoutId = null;
let designerInstance = null;

// Parsed data storage for Run view
let runParsedData = null;

function getCurrentUser() {
  return window.AuthStore ? window.AuthStore.currentUser() : null;
}

function updateUserChrome() {
  const user = getCurrentUser();
  const label = document.getElementById('current-user-label');
  const addUserBtn = document.getElementById('btn-add-user');

  if (label) label.textContent = user ? `User: ${user.username}` : '';
  if (addUserBtn) addUserBtn.classList.toggle('hidden', !user?.isSuperUser);
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
  const layouts = loadLayouts();

  grid.innerHTML = '';
  if (layouts.length === 0) {
    noMsg.classList.remove('hidden');
    return;
  }
  noMsg.classList.add('hidden');

  layouts.forEach(layout => {
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
        <span>Created ${formatDate(layout.createdAt)}</span>
      </div>
      <div class="layout-card-actions">
        <button class="btn btn-primary btn-sm" data-action="edit" data-id="${layout.id}">Edit</button>
        <button class="btn btn-secondary btn-sm" data-action="run" data-id="${layout.id}">Run</button>
        <button class="btn btn-ghost btn-sm" data-action="export" data-id="${layout.id}" title="Download as JSON">&#8595; Export</button>
        <button class="btn btn-danger btn-sm" data-action="delete" data-id="${layout.id}">Delete</button>
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
  document.getElementById('layout-name').value = prefill?.name || '';
  document.getElementById('page-size').value = prefill?.page?.size || 'A4';
  document.getElementById('page-orientation').value = prefill?.page?.orientation || 'portrait';
  document.getElementById('margin-top').value = prefill?.page?.marginTop ?? 15;
  document.getElementById('margin-bottom').value = prefill?.page?.marginBottom ?? 15;
  document.getElementById('margin-left').value = prefill?.page?.marginLeft ?? 15;
  document.getElementById('margin-right').value = prefill?.page?.marginRight ?? 15;
  document.getElementById('field-names-input').value = (prefill?.fields || []).join('\t');
  const ftList = document.getElementById('field-type-list');
  if (ftList) ftList.innerHTML = '';
  document.getElementById('parsed-fields-preview').classList.add('hidden');
}

function showSetupStep(n) {
  setupStep = n;
  document.getElementById('setup-step-1').classList.toggle('hidden', n !== 1);
  document.getElementById('setup-step-2').classList.toggle('hidden', n !== 2);
  document.getElementById('step-ind-1').classList.toggle('active', n === 1);
  document.getElementById('step-ind-1').classList.toggle('done', n > 1);
  document.getElementById('step-ind-2').classList.toggle('active', n === 2);
}

function createNewLayout() {
  const name = document.getElementById('layout-name').value.trim();
  const size = document.getElementById('page-size').value;
  const orientation = document.getElementById('page-orientation').value;
  const marginTop = parseFloat(document.getElementById('margin-top').value) || 15;
  const marginBottom = parseFloat(document.getElementById('margin-bottom').value) || 15;
  const marginLeft = parseFloat(document.getElementById('margin-left').value) || 15;
  const marginRight = parseFloat(document.getElementById('margin-right').value) || 15;
  const fieldText = document.getElementById('field-names-input').value;
  const fields = parseFieldNames(fieldText);

  // Collect field types from UI
  const fieldMeta = {};
  document.querySelectorAll('#field-type-list .field-type-row').forEach(row => {
    const fieldName = row.dataset.field;
    const activeBtn = row.querySelector('.ftype-btn.active');
    fieldMeta[fieldName] = { type: activeBtn?.dataset?.type || 'heading' };
  });

  const layout = {
    id: generateId(),
    name: name || 'Untitled Layout',
    createdAt: new Date().toISOString(),
    page: { size, orientation, marginTop, marginBottom, marginLeft, marginRight },
    fields,
    fieldMeta,
    elements: [],
  };

  saveLayout(layout);
  openDesigner(layout.id);
}

// ===== Designer View =====
function openDesigner(layoutId) {
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

// ===== Event wiring =====
function initHomeEvents() {
  document.getElementById('btn-logout').addEventListener('click', () => {
    if (designerInstance) {
      designerInstance.destroy();
      designerInstance = null;
    }
    if (window.AuthStore) window.AuthStore.logout();
    showLoginView();
  });

  document.getElementById('btn-add-user').addEventListener('click', () => {
    openAddUserModal();
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

  document.getElementById('layouts-grid').addEventListener('click', (e) => {
    const btn = e.target.closest('[data-action]');
    if (!btn) return;
    const action = btn.dataset.action;
    const id = btn.dataset.id;
    if (action === 'edit') {
      openDesigner(id);
    } else if (action === 'run') {
      openRunView(id);
    } else if (action === 'export') {
      exportLayout(id);
    } else if (action === 'delete') {
      if (confirm('Delete this layout? This cannot be undone.')) {
        deleteLayout(id);
        renderHomeView();
      }
    }
  });
}

function initLoginEvents() {
  document.getElementById('login-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const username = document.getElementById('login-user-id').value.trim();
    const password = document.getElementById('login-password').value;
    const error = document.getElementById('login-error');
    const btn = document.getElementById('btn-login');

    error.classList.add('hidden');
    btn.disabled = true;
    btn.textContent = 'Signing in...';

    try {
      await window.AuthStore.login(username, password);
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
  document.getElementById('new-user-id').value = '';
  document.getElementById('new-user-password').value = '';
  document.getElementById('admin-password-confirm').value = '';
  document.getElementById('reset-user-id').value = '';
  document.getElementById('reset-user-password').value = '';
  document.getElementById('reset-admin-password-confirm').value = '';
  document.getElementById('add-user-error').classList.add('hidden');
  document.getElementById('reset-user-error').classList.add('hidden');
  document.getElementById('modal-add-user').classList.remove('hidden');
}

function closeAddUserModal() {
  document.getElementById('modal-add-user').classList.add('hidden');
}

function initAddUserEvents() {
  document.getElementById('btn-add-user-close').addEventListener('click', closeAddUserModal);
  document.getElementById('btn-add-user-cancel').addEventListener('click', closeAddUserModal);
  document.getElementById('modal-add-user').addEventListener('click', (e) => {
    if (e.target === document.getElementById('modal-add-user')) closeAddUserModal();
  });

  document.getElementById('btn-add-user-save').addEventListener('click', async () => {
    const current = getCurrentUser();
    const username = document.getElementById('new-user-id').value.trim();
    const password = document.getElementById('new-user-password').value;
    const adminPassword = document.getElementById('admin-password-confirm').value;
    const error = document.getElementById('add-user-error');
    const btn = document.getElementById('btn-add-user-save');

    error.classList.add('hidden');

    if (!current?.isSuperUser) {
      error.textContent = 'Only the super user can add users.';
      error.classList.remove('hidden');
      return;
    }
    if (!username || !password || !adminPassword) {
      error.textContent = 'Enter user id, password, and your super user password.';
      error.classList.remove('hidden');
      return;
    }

    btn.disabled = true;
    btn.textContent = 'Creating...';
    try {
      await window.AuthStore.addUser(current.username, adminPassword, username, password);
      closeAddUserModal();
      showToast(`User "${username}" created.`);
    } catch (err) {
      error.textContent = err.message || 'Could not create user.';
      error.classList.remove('hidden');
    } finally {
      btn.disabled = false;
      btn.textContent = 'Create User';
    }
  });

  document.getElementById('btn-reset-user-password').addEventListener('click', async () => {
    const current = getCurrentUser();
    const username = document.getElementById('reset-user-id').value.trim();
    const newPassword = document.getElementById('reset-user-password').value;
    const adminPassword = document.getElementById('reset-admin-password-confirm').value;
    const error = document.getElementById('reset-user-error');
    const btn = document.getElementById('btn-reset-user-password');

    error.classList.add('hidden');

    if (!current?.isSuperUser) {
      error.textContent = 'Only the super user can reset passwords.';
      error.classList.remove('hidden');
      return;
    }
    if (!username || !newPassword || !adminPassword) {
      error.textContent = 'Enter user id, new password, and your super user password.';
      error.classList.remove('hidden');
      return;
    }

    btn.disabled = true;
    btn.textContent = 'Resetting...';
    try {
      await window.AuthStore.resetPassword(current.username, adminPassword, username, newPassword);
      document.getElementById('reset-user-id').value = '';
      document.getElementById('reset-user-password').value = '';
      document.getElementById('reset-admin-password-confirm').value = '';
      showToast(`Password reset for "${username}".`);
    } catch (err) {
      error.textContent = err.message || 'Could not reset password.';
      error.classList.remove('hidden');
    } finally {
      btn.disabled = false;
      btn.textContent = 'Reset Password';
    }
  });
}

function initSetupEvents() {
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

  document.getElementById('btn-create-layout').addEventListener('click', () => {
    createNewLayout();
  });

  document.getElementById('field-names-input').addEventListener('input', () => {
    const text = document.getElementById('field-names-input').value;
    const fields = parseFieldNames(text);
    const preview = document.getElementById('parsed-fields-preview');
    const list = document.getElementById('field-type-list');
    if (fields.length > 0) {
      preview.classList.remove('hidden');
      // Build field type rows
      list.innerHTML = '';
      fields.forEach(f => {
        const row = document.createElement('div');
        row.className = 'field-type-row';
        row.dataset.field = f;
        row.innerHTML = `
          <span class="field-name">${escapeHtml(f)}</span>
          <div class="field-type-toggle">
            <button class="ftype-btn heading active" data-type="heading" data-field="${escapeHtml(f)}" title="Single value">Heading</button>
            <button class="ftype-btn detail" data-type="detail" data-field="${escapeHtml(f)}" title="Repeating rows">Detail</button>
          </div>
        `;
        row.querySelectorAll('.ftype-btn').forEach(btn => {
          btn.addEventListener('click', () => {
            row.querySelectorAll('.ftype-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
          });
        });
        list.appendChild(row);
      });
    } else {
      preview.classList.add('hidden');
      list.innerHTML = '';
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
    });
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
      document.querySelectorAll('#run-fields-form input[data-field]').forEach(input => {
        fieldValues[input.dataset.field] = input.value;
      });
      detailRows = [fieldValues];
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
