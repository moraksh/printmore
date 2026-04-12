/**
 * Supabase-backed layout storage with localStorage fallback.
 *
 * Only layout definitions are persisted. Runtime values pasted or typed while
 * generating PDFs stay in memory and are never written by this module.
 */

'use strict';

const LayoutStore = (() => {
  const LOCAL_KEY = 'printLayouts';
  const cfg = window.PRINT_LAYOUT_CONFIG || {};
  const tableName = cfg.SUPABASE_LAYOUTS_TABLE || 'layouts';
  const hasSupabaseConfig = Boolean(cfg.SUPABASE_URL && cfg.SUPABASE_ANON_KEY);
  const hasSupabaseClient = Boolean(window.supabase && window.supabase.createClient);
  const client = hasSupabaseConfig && hasSupabaseClient
    ? window.supabase.createClient(cfg.SUPABASE_URL, cfg.SUPABASE_ANON_KEY)
    : null;

  let cache = [];
  let ready = false;
  let lastError = null;

  function readLocal() {
    try {
      const raw = localStorage.getItem(LOCAL_KEY);
      return raw ? JSON.parse(raw) : [];
    } catch {
      return [];
    }
  }

  function writeLocal(layouts) {
    localStorage.setItem(LOCAL_KEY, JSON.stringify(layouts));
  }

  function normalizeLayout(rowOrLayout) {
    if (!rowOrLayout) return null;
    const layout = rowOrLayout.layout || rowOrLayout;
    return layout && typeof layout === 'object' ? layout : null;
  }

  async function init() {
    cache = readLocal();
    if (!client) {
      ready = true;
      return cache;
    }

    try {
      const { data, error } = await client
        .from(tableName)
        .select('layout')
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
  }

  async function save(layout) {
    const cleanLayout = JSON.parse(JSON.stringify(layout));
    upsertLocal(cleanLayout);

    if (!client) return { ok: true, mode: 'local' };

    try {
      const { error } = await client
        .from(tableName)
        .upsert({
          id: cleanLayout.id,
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

    if (!client) return { ok: true, mode: 'local' };

    try {
      const { error } = await client.from(tableName).delete().eq('id', id);
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
    save,
    remove,
    status,
    isSupabaseEnabled: () => Boolean(client),
  };
})();

window.LayoutStore = LayoutStore;
