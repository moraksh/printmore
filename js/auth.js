/**
 * PrintMore user login helpers.
 *
 * Passwords are verified by Supabase SQL functions. The browser stores only
 * the logged-in user's id, username, and super-user flag for the current tab.
 */

'use strict';

const AuthStore = (() => {
  const SESSION_KEY = 'printmoreCurrentUser';
  const cfg = window.PRINT_LAYOUT_CONFIG || {};
  const hasSupabaseConfig = Boolean(cfg.SUPABASE_URL && cfg.SUPABASE_ANON_KEY);
  const hasSupabaseClient = Boolean(window.supabase && window.supabase.createClient);
  const client = hasSupabaseConfig && hasSupabaseClient
    ? window.supabase.createClient(cfg.SUPABASE_URL, cfg.SUPABASE_ANON_KEY)
    : null;

  function normalizeUser(user) {
    if (!user) return null;
    return {
      id: user.id,
      username: String(user.username || '').toUpperCase(),
      role: user.role || (user.is_super_user || user.isSuperUser ? 'super' : 'user'),
      isSuperUser: Boolean(user.is_super_user ?? user.isSuperUser),
    };
  }

  function currentUser() {
    try {
      return normalizeUser(JSON.parse(sessionStorage.getItem(SESSION_KEY)));
    } catch {
      return null;
    }
  }

  function setCurrentUser(user) {
    const normalized = normalizeUser(user);
    if (!normalized) return;
    sessionStorage.setItem(SESSION_KEY, JSON.stringify(normalized));
  }

  function logout() {
    sessionStorage.removeItem(SESSION_KEY);
  }

  async function login(username, password) {
    if (!client) throw new Error('Supabase is not configured.');

    const { data, error } = await client.rpc('authenticate_printmore_user', {
      p_username: username,
      p_password: password,
    });

    if (error) throw error;
    const user = Array.isArray(data) ? data[0] : data;
    if (!user) throw new Error('Invalid user id or password.');

    setCurrentUser(user);
    return normalizeUser(user);
  }

  async function addUser(adminUsername, adminPassword, username, password, role = 'user') {
    if (!client) throw new Error('Supabase is not configured.');

    const { data, error } = await client.rpc('create_printmore_user', {
      p_admin_username: adminUsername,
      p_admin_password: adminPassword,
      p_username: username,
      p_password: password,
      p_role: role,
    });

    if (error) throw error;
    return normalizeUser(Array.isArray(data) ? data[0] : data);
  }

  async function resetPassword(adminUsername, adminPassword, username, newPassword) {
    if (!client) throw new Error('Supabase is not configured.');

    const { data, error } = await client.rpc('reset_printmore_user_password', {
      p_admin_username: adminUsername,
      p_admin_password: adminPassword,
      p_username: username,
      p_new_password: newPassword,
    });

    if (error) throw error;
    return normalizeUser(Array.isArray(data) ? data[0] : data);
  }

  return {
    currentUser,
    login,
    logout,
    addUser,
    resetPassword,
  };
})();

window.AuthStore = AuthStore;
