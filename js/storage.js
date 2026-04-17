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
  const cfg = window.PRINT_LAYOUT_CONFIG || {};
  const dbCfg = cfg.DATABASE || {};
  const provider = dbCfg.PROVIDER || 'supabase';
  const supabaseCfg = dbCfg.SUPABASE || {};
  const rpc = cfg.RPC || {};
  const hasSupabaseConfig = provider === 'supabase' && Boolean(supabaseCfg.URL && supabaseCfg.ANON_KEY);
  const hasSupabaseClient = Boolean(window.supabase && window.supabase.createClient);
  const client = hasSupabaseConfig && hasSupabaseClient
    ? window.supabase.createClient(supabaseCfg.URL, supabaseCfg.ANON_KEY)
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

  function rpcName(key, fallback) {
    return rpc[key] || fallback;
  }

  function getSessionToken() {
    return window.AuthStore?.sessionToken?.() || '';
  }

  async function init(user) {
    currentUser = user || currentUser;
    cache = readLocal();
    if (provider !== 'supabase' || !client || !currentUser || !getSessionToken()) {
      ready = true;
      return cache;
    }

    try {
      const { data, error } = await client.rpc(rpcName('LIST_LAYOUTS', 'list_printmore_layouts'), {
        p_session_token: getSessionToken(),
      });

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
    if (!client || !getSessionToken()) return smartRulesCache;

    try {
      const { data, error } = await client.rpc(rpcName('GET_SMART_RULES', 'get_printmore_smart_rules'), {
        p_session_token: getSessionToken(),
      });

      if (error) throw error;
      if (data) {
        smartRulesCache = normalizeSmartRules(data);
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

    if (!client || !getSessionToken()) return { ok: true, mode: 'local' };

    try {
      const { error } = await client.rpc(rpcName('SET_SMART_RULES', 'set_printmore_smart_rules'), {
        p_session_token: getSessionToken(),
        p_rules: smartRulesCache,
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

    if (!client || !currentUser || !getSessionToken()) return { ok: true, mode: 'local' };

    try {
      const { error } = await client.rpc(rpcName('UPSERT_LAYOUT', 'upsert_printmore_layout'), {
        p_session_token: getSessionToken(),
        p_layout_id: cleanLayout.id,
        p_name: cleanLayout.name || 'Untitled Layout',
        p_layout: cleanLayout,
        p_target_username: null,
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

  async function saveForUser(layout, targetUser) {
    if (!client || !targetUser?.username || !getSessionToken()) return { ok: false, error: new Error('Supabase is not configured.') };
    const cleanLayout = JSON.parse(JSON.stringify(layout || {}));
    cleanLayout.username = String(targetUser.username || '').toUpperCase();

    try {
      const { error } = await client.rpc(rpcName('UPSERT_LAYOUT', 'upsert_printmore_layout'), {
        p_session_token: getSessionToken(),
        p_layout_id: cleanLayout.id,
        p_name: cleanLayout.name || 'Untitled Layout',
        p_layout: cleanLayout,
        p_target_username: cleanLayout.username,
      });
      if (error) throw error;
      return { ok: true, mode: 'supabase' };
    } catch (err) {
      console.error('Supabase cross-user layout save failed.', err);
      if (err?.code === '23505') {
        return { ok: false, error: new Error(`Layout name "${cleanLayout.name}" already exists for ${cleanLayout.username}.`) };
      }
      return { ok: false, error: err };
    }
  }

  async function remove(id) {
    cache = cache.filter(layout => layout.id !== id);
    writeLocal(cache);
    notifyChange('delete', id);

    if (!client || !currentUser || !getSessionToken()) return { ok: true, mode: 'local' };

    try {
      const { error } = await client.rpc(rpcName('DELETE_LAYOUT', 'delete_printmore_layout'), {
        p_session_token: getSessionToken(),
        p_layout_id: id,
      });
      if (error) throw error;
      lastError = null;
      return { ok: true, mode: 'supabase' };
    } catch (err) {
      console.error('Supabase layout delete failed; removed local copy only.', err);
      lastError = err;
      return { ok: false, mode: 'local', error: err };
    }
  }

  async function listForUser(userId) {
    if (!client) throw new Error('Supabase is not configured.');
    if (!userId) throw new Error('User id is required.');
    const token = getSessionToken();
    if (!token) throw new Error('Session expired. Please sign in again.');
    const { data, error } = await client.rpc(rpcName('LIST_LAYOUTS_FOR_USER', 'list_printmore_layouts_for_user'), {
      p_session_token: token,
      p_target_user_id: userId,
    });
    if (error) throw error;
    return (data || []).map(row => {
      const layout = normalizeLayout(row) || {};
      layout.id = layout.id || row.id;
      layout.name = layout.name || row.name || 'Untitled Layout';
      layout.createdAt = layout.createdAt || row.created_at || null;
      layout.updatedAt = layout.updatedAt || row.updated_at || null;
      layout.userId = layout.userId || row.user_id || userId;
      return layout;
    });
  }

  async function removeForUser(layoutId, userId) {
    if (!client) throw new Error('Supabase is not configured.');
    if (!layoutId || !userId) throw new Error('Layout id and user id are required.');
    const token = getSessionToken();
    if (!token) throw new Error('Session expired. Please sign in again.');
    const { error } = await client.rpc(rpcName('DELETE_LAYOUT_FOR_USER', 'delete_printmore_layout_for_user'), {
      p_session_token: token,
      p_layout_id: layoutId,
      p_target_user_id: userId,
    });
    if (error) throw error;
    return { ok: true };
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
    saveForUser,
    remove,
    listForUser,
    removeForUser,
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
