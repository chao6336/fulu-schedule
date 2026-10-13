// 雲端版後端：把 multiuser.js 發出的 /api/ 請求，改成呼叫 Supabase。
// 權限與資料檢查都在資料庫的函式裡（supabase-setup.sql），這裡只負責轉接。
(function () {
  const config = window.CLOUD_CONFIG || {};
  const configured = /^https:\/\/\S+$/.test(config.supabaseUrl || '') && Boolean(config.supabaseAnonKey);
  const client = configured && window.supabase ? window.supabase.createClient(config.supabaseUrl, config.supabaseAnonKey) : null;
  const STATUS = { forbidden: 403, conflict: 409, login_required: 401, last_admin: 409, cannot_delete_self: 409, user_missing: 404 };
  const KNOWN = new Set(['forbidden', 'conflict', 'no_ops', 'bad_ops', 'bad_record', 'duplicate_id', 'bad_role', 'bad_name', 'user_missing', 'last_admin', 'cannot_delete_self']);
  const backend = window.MU_BACKEND = { recovery: false };

  function fail(code, status) {
    const error = new Error(code);
    error.code = code;
    error.status = status || STATUS[code] || 400;
    return error;
  }

  // 連不上網路時不帶錯誤代碼，畫面會顯示「連線中斷」而不是把人登出。
  const offline = error => /failed to fetch|networkerror|load failed/i.test(error?.message || '');

  function authFail(error) {
    if (offline(error)) return fail('', 0);
    const text = (error?.message || '').toLowerCase(), code = error?.code || '';
    if (code === 'email_not_confirmed' || text.includes('not confirmed')) return fail('email_not_confirmed', 401);
    if (code === 'invalid_credentials' || text.includes('invalid login')) return fail('bad_login', 401);
    if (code === 'user_already_exists' || text.includes('already registered')) return fail('user_exists', 409);
    if (code === 'weak_password' || text.includes('password should')) return fail('weak_password', 400);
    if (code === 'same_password') return fail('weak_password', 400);
    if (text.includes('rate limit') || error?.status === 429) return fail('locked', 429);
    if (text.includes('email') && (text.includes('invalid') || text.includes('valid'))) return fail('bad_email', 400);
    return fail('server_error', 500);
  }

  async function rpc(name, args) {
    const { data, error } = await client.rpc(name, args || {});
    if (!error) return data;
    if (offline(error)) throw fail('', 0);
    const code = (error.message || '').trim();
    if (KNOWN.has(code)) throw fail(code);
    if (error.code === '42501' || error.code === 'PGRST301' || /jwt/i.test(error.message || '')) throw fail('login_required', 401);
    console.error('雲端回應錯誤', name, error);
    throw fail('server_error', 500);
  }

  async function hasSession() {
    const { data } = await client.auth.getSession();
    return Boolean(data.session);
  }

  async function currentUser() {
    if (!await hasSession()) return null;
    return await rpc('get_me');
  }

  const here = () => location.origin + location.pathname;

  async function signUp(body) {
    if (body.password.length < 6) throw fail('weak_password', 400);
    const { data, error } = await client.auth.signUp({
      email: body.username, password: body.password,
      options: { data: { name: body.name || '' }, emailRedirectTo: here() }
    });
    if (error) throw authFail(error);
    if (!data.session) return { notice: 'confirm_email' };
    return { user: await rpc('get_me') };
  }

  if (client) {
    client.auth.onAuthStateChange(event => {
      if (event === 'PASSWORD_RECOVERY') {
        backend.recovery = true;
        window.MU?.onRecovery?.();
      }
    });
  }

  backend.api = async function (method, path, body) {
    if (!client) throw fail('not_configured', 0);

    if (method === 'GET') {
      if (path === '/api/session') {
        // 等 Supabase 處理完網址裡的登入資訊（例如重設密碼連結）。
        const signedIn = await hasSession();
        if (backend.recovery) return { recovery: true };
        if (signedIn) return { setupNeeded: false, isLocal: true, user: await rpc('get_me') };
        return { setupNeeded: await rpc('needs_setup'), isLocal: true, user: null };
      }
      if (!await hasSession()) throw fail('login_required', 401);
      if (path === '/api/version') {
        const info = await rpc('get_version');
        if (!info.user) throw fail('login_required', 401);
        return info;
      }
      if (path === '/api/state') return await rpc('get_state');
      if (path === '/api/users') return { users: await rpc('admin_list_users') };
      throw fail('not_found', 404);
    }

    if (path === '/api/setup' || path === '/api/register') return await signUp(body);

    if (path === '/api/login') {
      const { error } = await client.auth.signInWithPassword({ email: body.username, password: body.password });
      if (error) throw authFail(error);
      return { user: await rpc('get_me') };
    }

    if (path === '/api/forgot') {
      const { error } = await client.auth.resetPasswordForEmail(body.username, { redirectTo: here() });
      if (error) throw authFail(error);
      return { notice: 'reset_sent' };
    }

    if (path === '/api/recover') {
      if (body.password.length < 6) throw fail('weak_password', 400);
      const { error } = await client.auth.updateUser({ password: body.password });
      if (error) throw authFail(error);
      backend.recovery = false;
      history.replaceState(null, '', here());
      return { user: await rpc('get_me') };
    }

    if (path === '/api/logout') {
      await client.auth.signOut({ scope: 'local' });
      return { ok: true };
    }

    const user = await currentUser();
    if (!user) throw fail('login_required', 401);

    if (path === '/api/password') {
      if ((body.newPassword || '').length < 6) throw fail('weak_password', 400);
      const check = await client.auth.signInWithPassword({ email: user.username, password: body.oldPassword });
      if (check.error) throw offline(check.error) ? fail('', 0) : fail('bad_old_password', 403);
      const { error } = await client.auth.updateUser({ password: body.newPassword });
      if (error) throw authFail(error);
      return { ok: true };
    }

    if (path === '/api/ops') return await rpc('apply_ops', { ops: body.ops });
    if (path === '/api/users/save') return await rpc('admin_save_user', { p_email: body.username, p_name: body.name || '', p_role: body.role, p_active: Boolean(body.active) });
    if (path === '/api/users/delete') return await rpc('admin_delete_user', { p_email: body.username });
    throw fail('not_found', 404);
  };
})();
