/**
 * Supabase-backed layout storage with localStorage fallback.
 *
 * Only layout definitions are persisted. Runtime values pasted or typed while
 * generating PDFs stay in memory and are never written by this module.
 */

'use strict';

const LayoutStore = (() => {
  const LOCAL_KEY = 'printLayouts';
  const SMART_RULES_KEY = 'printmoreSmartRules';
  const settingsTableName = 'app_settings';
  const cfg = window.PRINT_LAYOUT_CONFIG || {};
  const tableName = cfg.SUPABASE_LAYOUTS_TABLE || 'layouts';
  const hasSupabaseConfig = Boolean(cfg.SUPABASE_URL && cfg.SUPABASE_ANON_KEY);
  const hasSupabaseClient = Boolean(window.supabase && window.supabase.createClient);
  const client = hasSupabaseConfig && hasSupabaseClient
    ? window.supabase.createClient(cfg.SUPABASE_URL, cfg.SUPABASE_ANON_KEY)
    : null;
  const channel = typeof BroadcastChannel !== 'undefined'
    ? new BroadcastChannel('printmore-layouts')
    : null;

  let cache = [];
  let smartRulesCache = null;
  let ready = false;
  let lastError = null;
  let currentUser = null;

  const defaultSmartRules = {
    header: ['customer', 'vendor', 'supplier', 'date', 'doc', 'document', 'invoice', 'order', 'vehicle', 'driver', 'warehouse', 'from', 'to', 'reference', 'ref', 'grn', 'delivery', 'carrier'],
    table: ['item', 'sku', 'material', 'product', 'description', 'desc', 'qty', 'quantity', 'uom', 'unit', 'weight', 'pallet', 'carton', 'box', 'batch', 'lot', 'serial', 'bin', 'code'],
    footer: ['prepared', 'checked', 'received', 'approved', 'signature', 'remarks', 'notes', 'total', 'printed', 'user'],
  };

  function readLocal() {
    try {
      const key = currentUser ? `${LOCAL_KEY}:${currentUser.username}` : LOCAL_KEY;
      const raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : [];
    } catch {
      return [];
    }
  }

  function normalizeSmartRules(rules) {
    const source = rules && typeof rules === 'object' ? rules : {};
    return ['header', 'table', 'footer'].reduce((acc, key) => {
      const values = Array.isArray(source[key]) ? source[key] : defaultSmartRules[key];
      acc[key] = [...new Set(values
        .map(value => String(value || '').trim().toLowerCase())
        .filter(Boolean))]
        .sort((a, b) => a.localeCompare(b));
      return acc;
    }, {});
  }

  function readLocalSmartRules() {
    try {
      const raw = localStorage.getItem(SMART_RULES_KEY);
      return normalizeSmartRules(raw ? JSON.parse(raw) : defaultSmartRules);
    } catch {
      return normalizeSmartRules(defaultSmartRules);
    }
  }

  function writeLocalSmartRules(rules) {
    localStorage.setItem(SMART_RULES_KEY, JSON.stringify(normalizeSmartRules(rules)));
  }

  function writeLocal(layouts) {
    const key = currentUser ? `${LOCAL_KEY}:${currentUser.username}` : LOCAL_KEY;
    localStorage.setItem(key, JSON.stringify(layouts));
  }

  function syncFromLocal() {
    cache = readLocal();
    return cache;
  }

  function notifyChange(type, id) {
    if (channel) channel.postMessage({ type, id });
  }

  function normalizeLayout(rowOrLayout) {
    if (!rowOrLayout) return null;
    const layout = rowOrLayout.layout || rowOrLayout;
    return layout && typeof layout === 'object' ? layout : null;
  }

  async function init(user) {
    currentUser = user || currentUser;
    cache = readLocal();
    if (!client || !currentUser) {
      ready = true;
      return cache;
    }

    try {
      const { data, error } = await client
        .from(tableName)
        .select('layout')
        .eq('user_id', currentUser.id)
        .order('updated_at', { ascending: false });

      if (error) throw error;

      cache = (data || []).map(normalizeLayout).filter(Boolean);
      writeLocal(cache);
      lastError = null;
    } catch (err) {
      console.error('Supabase layout load failed; using local layouts.', err);
      lastError = err;
    } finally {
      ready = true;
    }

    return cache;
  }

  async function loadSmartRules() {
    smartRulesCache = readLocalSmartRules();
    if (!client) return smartRulesCache;

    try {
      const { data, error } = await client
        .from(settingsTableName)
        .select('value')
        .eq('key', SMART_RULES_KEY)
        .maybeSingle();

      if (error) throw error;
      if (data?.value) {
        smartRulesCache = normalizeSmartRules(data.value);
        writeLocalSmartRules(smartRulesCache);
      }
      lastError = null;
    } catch (err) {
      console.error('Supabase smart rules load failed; using local rules.', err);
      lastError = err;
    }

    return smartRulesCache;
  }

  async function saveSmartRules(rules) {
    smartRulesCache = normalizeSmartRules(rules);
    writeLocalSmartRules(smartRulesCache);

    if (!client) return { ok: true, mode: 'local' };

    try {
      const { error } = await client
        .from(settingsTableName)
        .upsert({
          key: SMART_RULES_KEY,
          value: smartRulesCache,
          updated_at: new Date().toISOString(),
        });

      if (error) throw error;
      lastError = null;
      return { ok: true, mode: 'supabase' };
    } catch (err) {
      console.error('Supabase smart rules save failed; kept local copy.', err);
      lastError = err;
      return { ok: false, mode: 'local', error: err };
    }
  }

  function getSmartRules() {
    if (!smartRulesCache) smartRulesCache = readLocalSmartRules();
    return normalizeSmartRules(smartRulesCache);
  }

  function getAll() {
    return cache;
  }

  function getById(id) {
    return cache.find(layout => layout.id === id) || null;
  }

  function upsertLocal(layout) {
    const idx = cache.findIndex(item => item.id === layout.id);
    if (idx >= 0) {
      cache[idx] = layout;
    } else {
      cache.push(layout);
    }
    writeLocal(cache);
    notifyChange('save', layout.id);
  }

  async function save(layout) {
    const cleanLayout = JSON.parse(JSON.stringify(layout));
    if (currentUser) {
      cleanLayout.userId = currentUser.id;
      cleanLayout.username = currentUser.username;
    }
    upsertLocal(cleanLayout);

    if (!client || !currentUser) return { ok: true, mode: 'local' };

    try {
      const { error } = await client
        .from(tableName)
        .upsert({
          id: cleanLayout.id,
          user_id: currentUser.id,
          name: cleanLayout.name || 'Untitled Layout',
          layout: cleanLayout,
          updated_at: new Date().toISOString(),
        });

      if (error) throw error;
      lastError = null;
      return { ok: true, mode: 'supabase' };
    } catch (err) {
      console.error('Supabase layout save failed; kept local copy.', err);
      lastError = err;
      return { ok: false, mode: 'local', error: err };
    }
  }

  async function remove(id) {
    cache = cache.filter(layout => layout.id !== id);
    writeLocal(cache);
    notifyChange('delete', id);

    if (!client || !currentUser) return { ok: true, mode: 'local' };

    try {
      const { error } = await client
        .from(tableName)
        .delete()
        .eq('id', id)
        .eq('user_id', currentUser.id);
      if (error) throw error;
      lastError = null;
      return { ok: true, mode: 'supabase' };
    } catch (err) {
      console.error('Supabase layout delete failed; removed local copy only.', err);
      lastError = err;
      return { ok: false, mode: 'local', error: err };
    }
  }

  function status() {
    if (client && !lastError) return ready ? 'supabase' : 'loading';
    if (client && lastError) return 'offline';
    return 'local';
  }

  return {
    init,
    getAll,
    getById,
    syncFromLocal,
    save,
    remove,
    status,
    getSmartRules,
    loadSmartRules,
    saveSmartRules,
    onExternalChange: (handler) => {
      if (!channel) return;
      channel.addEventListener('message', event => handler(event.data));
    },
    isSupabaseEnabled: () => Boolean(client),
  };
})();

window.LayoutStore = LayoutStore;
