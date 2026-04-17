/**
 * PrintMore user login helpers.
 *
 * Passwords are verified by Supabase SQL functions. The browser stores only
 * the logged-in user's id, username, role, and super-user flag for the tab.
 */

'use strict';

const AuthStore = (() => {
  const SESSION_KEY = 'printmoreCurrentUser';
  const TOKEN_KEY = 'printmoreSessionToken';
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

  function rpcName(key, fallback) {
    return rpc[key] || fallback;
  }

  function normalizeUser(user) {
    if (!user) return null;
    return {
      id: user.id,
      username: String(user.username || '').toUpperCase(),
      fullName: String(user.full_name || user.fullName || '').trim(),
      role: user.role || (user.is_super_user || user.isSuperUser ? 'super' : 'user'),
      isSuperUser: Boolean(user.is_super_user ?? user.isSuperUser),
    };
  }

  function normalizeManagedUser(user) {
    if (!user) return null;
    return {
      id: user.id,
      username: String(user.username || '').toUpperCase(),
      fullName: String(user.full_name || user.fullName || '').trim(),
      role: user.role || 'user',
      isSuperUser: Boolean(user.is_super_user ?? user.isSuperUser),
      active: Boolean(user.active),
      lastLoginAt: user.last_login_at || user.lastLoginAt || null,
    };
  }

  function currentUser() {
    try {
      return normalizeUser(JSON.parse(sessionStorage.getItem(SESSION_KEY)));
    } catch {
      return null;
    }
  }

  function sessionToken() {
    try {
      return sessionStorage.getItem(TOKEN_KEY) || '';
    } catch {
      return '';
    }
  }

  function setCurrentUser(user, token) {
    const normalized = normalizeUser(user);
    if (!normalized) return;
    sessionStorage.setItem(SESSION_KEY, JSON.stringify(normalized));
    if (token) sessionStorage.setItem(TOKEN_KEY, token);
  }

  async function logout() {
    const token = sessionToken();
    if (client && token) {
      try {
        await client.rpc(rpcName('LOGOUT_USER', 'logout_printmore_user'), { p_session_token: token });
      } catch {
        // Best effort only; always clear local session.
      }
    }
    sessionStorage.removeItem(SESSION_KEY);
    sessionStorage.removeItem(TOKEN_KEY);
  }

  async function login(username, password) {
    if (provider !== 'supabase') throw new Error('Database provider is not configured for browser mode yet.');
    if (!client) throw new Error('Supabase is not configured.');

    const { data, error } = await client.rpc(rpcName('AUTHENTICATE_USER', 'authenticate_printmore_user'), {
      p_username: username,
      p_password: password,
    });

    if (error) throw error;
    const raw = Array.isArray(data) ? data[0] : data;
    if (!raw) throw new Error('Invalid user id or password.');
    const token = raw.session_token;
    const user = { ...raw };
    delete user.session_token;

    setCurrentUser(user, token);
    return normalizeUser(user);
  }

  async function addUser(username, fullName, password, role = 'user') {
    if (!client) throw new Error('Supabase is not configured.');
    const token = sessionToken();
    if (!token) throw new Error('Session expired. Please sign in again.');

    const { data, error } = await client.rpc(rpcName('ADD_USER', 'create_printmore_user'), {
      p_session_token: token,
      p_username: username,
      p_full_name: fullName,
      p_password: password,
      p_role: role,
    });

    if (error) throw error;
    return normalizeUser(Array.isArray(data) ? data[0] : data);
  }

  async function resetPassword(username, newPassword) {
    if (!client) throw new Error('Supabase is not configured.');
    const token = sessionToken();
    if (!token) throw new Error('Session expired. Please sign in again.');

    const { data, error } = await client.rpc(rpcName('RESET_PASSWORD', 'reset_printmore_user_password'), {
      p_session_token: token,
      p_username: username,
      p_new_password: newPassword,
    });

    if (error) throw error;
    return normalizeUser(Array.isArray(data) ? data[0] : data);
  }

  async function setUserActive(username, active) {
    if (!client) throw new Error('Supabase is not configured.');
    const token = sessionToken();
    if (!token) throw new Error('Session expired. Please sign in again.');

    const { data, error } = await client.rpc(rpcName('SET_USER_ACTIVE', 'set_printmore_user_active'), {
      p_session_token: token,
      p_username: username,
      p_active: active,
    });

    if (error) throw error;
    return normalizeUser(Array.isArray(data) ? data[0] : data);
  }

  async function updateUser(username, role, active) {
    if (!client) throw new Error('Supabase is not configured.');
    const token = sessionToken();
    if (!token) throw new Error('Session expired. Please sign in again.');

    const { data, error } = await client.rpc(rpcName('UPDATE_USER', 'update_printmore_user'), {
      p_session_token: token,
      p_username: username,
      p_role: role,
      p_active: active,
    });

    if (error) throw error;
    return normalizeUser(Array.isArray(data) ? data[0] : data);
  }

  async function findUser(username) {
    if (!client) throw new Error('Supabase is not configured.');
    const token = sessionToken();
    if (!token) throw new Error('Session expired. Please sign in again.');
    const { data, error } = await client.rpc(rpcName('FIND_USER', 'find_printmore_user'), {
      p_session_token: token,
      p_username: username,
    });
    if (error) throw error;
    const user = Array.isArray(data) ? data[0] : data;
    if (!user) return null;
    return normalizeUser(user);
  }

  async function listUsers() {
    if (!client) throw new Error('Supabase is not configured.');
    const token = sessionToken();
    if (!token) throw new Error('Session expired. Please sign in again.');
    const { data, error } = await client.rpc(rpcName('LIST_USERS', 'list_printmore_users'), {
      p_session_token: token,
    });
    if (error) throw error;
    return Array.isArray(data) ? data.map(normalizeManagedUser).filter(Boolean) : [];
  }

  async function deleteUser(username) {
    if (!client) throw new Error('Supabase is not configured.');
    const token = sessionToken();
    if (!token) throw new Error('Session expired. Please sign in again.');
    const { data, error } = await client.rpc(rpcName('DELETE_USER', 'delete_printmore_user'), {
      p_session_token: token,
      p_username: username,
    });
    if (error) throw error;
    return normalizeUser(Array.isArray(data) ? data[0] : data);
  }

  return {
    currentUser,
    sessionToken,
    login,
    logout,
    addUser,
    resetPassword,
    setUserActive,
    updateUser,
    findUser,
    listUsers,
    deleteUser,
  };
})();

window.AuthStore = AuthStore;
