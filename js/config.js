/**
 * PrintMore server/database configuration.
 *
 * SINGLE PLACE TO CHANGE SERVER DETAILS:
 *   Edit this file only: js/config.js
 *
 * Current provider: "supabase"
 * Future-ready provider placeholder: "sqlserver"
 *
 * IMPORTANT:
 * - Never put service role keys or DB passwords in browser config.
 * - For SQL Server in future, expose only your backend API URL/token here.
 */
window.PRINT_LAYOUT_CONFIG = {
  DATABASE: {
    PROVIDER: 'supabase', // 'supabase' | 'sqlserver'

    SUPABASE: {
      URL: 'https://bvtabvtthuqpfhrgjklk.supabase.co',
      ANON_KEY: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJ2dGFidnR0aHVxcGZocmdqa2xrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYwMDk1MDQsImV4cCI6MjA5MTU4NTUwNH0.rtwh4a-EAX0CXj2txU1r9pBgD44l2Kbm8ibSKerT_0k',
      LAYOUTS_TABLE: 'layouts',
    },

    // Placeholder for future migration.
    // Keep blank until backend API is ready.
    SQLSERVER: {
      API_BASE_URL: '',
      API_KEY: '',
      API_TIMEOUT_MS: 15000,
    },
  },

  // Centralized RPC names (so backend function names are managed in one place)
  RPC: {
    AUTHENTICATE_USER: 'authenticate_printmore_user',
    LOGOUT_USER: 'logout_printmore_user',
    ADD_USER: 'create_printmore_user',
    RESET_PASSWORD: 'reset_printmore_user_password',
    SET_USER_ACTIVE: 'set_printmore_user_active',
    UPDATE_USER: 'update_printmore_user',
    FIND_USER: 'find_printmore_user',
    LIST_USERS: 'list_printmore_users',
    DELETE_USER: 'delete_printmore_user',

    LIST_LAYOUTS: 'list_printmore_layouts',
    LIST_LAYOUTS_FOR_USER: 'list_printmore_layouts_for_user',
    UPSERT_LAYOUT: 'upsert_printmore_layout',
    DELETE_LAYOUT: 'delete_printmore_layout',
    DELETE_LAYOUT_FOR_USER: 'delete_printmore_layout_for_user',

    GET_SMART_RULES: 'get_printmore_smart_rules',
    SET_SMART_RULES: 'set_printmore_smart_rules',
  },
};
