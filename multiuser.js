// 多人版：登入、權限、與共用主機同步。
// app.js 只管畫面與排班邏輯；每次 save() 會呼叫 MU.push()，由這裡把變動送到主機。
(function () {
  // 雲端版會先載入 cloud-backend.js，由它代替區網主機回應同一組 /api/ 請求。
  const CLOUD = Boolean(window.MU_BACKEND);
  const ROLE_LABELS = { viewer: '可閱覽', editor: '可編輯', manager: '可編輯＋刪除', admin: '管理者' };
  const ROLE_LEVELS = { viewer: 1, editor: 2, manager: 3, admin: 4 };
  const roleLabel = role => role === 'pending' ? '待開通' : ROLE_LABELS[role] || role;
  const ROLE_NOTES = {
    viewer: '只能看排班、查班表與匯出報表',
    editor: '可新增與修改工作、服務人員及排班',
    manager: '除了編輯，還可以刪除工作與服務人員',
    admin: '全部功能，另可管理使用者、還原或清除全部資料'
  };
  const ERRORS = {
    forbidden: '你的權限不能做這個操作',
    conflict: '別人剛剛改過同一筆資料，畫面已更新，請再操作一次',
    bad_login: '帳號或密碼不正確',
    locked: '密碼錯太多次，請 1 分鐘後再試',
    weak_password: '密碼至少要 6 個字',
    bad_username: '帳號需 2～32 個字，不能有空白或特殊符號',
    bad_name: '姓名太長或含有不能使用的符號',
    bad_old_password: '目前的密碼不正確',
    user_exists: '這個帳號已經存在',
    user_missing: '找不到這個使用者，請重新整理',
    last_admin: '至少要保留一位啟用中的管理者',
    cannot_delete_self: '不能刪除自己正在使用的帳號',
    setup_host_only: '第一次設定必須在主機那台電腦上完成',
    already_setup: '系統已經設定過管理者，請直接登入',
    duplicate_id: '資料裡有重複的編號，無法匯入，請先修正備份檔',
    bad_record: '資料格式不正確，沒有儲存',
    too_large: '資料太大，無法一次送出',
    login_required: '登入已過期，請重新登入',
    bad_email: '請輸入正確的 Email',
    email_not_confirmed: '這個 Email 還沒確認，請先到信箱點確認連結',
    confirm_email: '註冊成功！請到信箱點確認連結，再回來登入',
    reset_sent: '已寄出重設密碼的信，請到信箱點連結',
    not_configured: '雲端連線資料還沒填好（config.js）'
  };
  const EDIT_ACTIONS = new Set(['new-task', 'edit-task', 'new-person', 'edit-person', 'import', 'import-work', 'confirm-work-import', 'resync-leaders']);
  const DELETE_ACTIONS = new Set(['delete-task', 'delete-person']);
  const LEGACY_KEY = 'fuluyuanman_scheduler_v1';

  const MU = window.MU = { me: null, version: 0, forceReplace: false };
  let snapshot = { event: { rev: 0, json: '' }, people: new Map(), tasks: new Map() };
  let chain = Promise.resolve();
  let busy = 0, polling = false, stale = false, pollTimer = null, editingUser = null;

  const message = error => ERRORS[error?.code] || '連線或主機發生問題，請稍後再試';
  const level = () => ROLE_LEVELS[MU.me?.role] || 0;

  async function api(method, path, body) {
    if (CLOUD) return window.MU_BACKEND.api(method, path, body || {});
    const options ={ method, credentials: 'same-origin', cache: 'no-store' };
    if (method !== 'GET') {
      options.headers = { 'Content-Type': 'application/json', 'X-Requested-With': 'fuluyuanman' };
      options.body = JSON.stringify(body || {});
    }
    const response = await fetch(path, options);
    let data = null;
    try { data = await response.json(); } catch (error) { /* 非 JSON 回應 */ }
    if (!response.ok) {
      const failure = new Error(data?.error || 'http_' + response.status);
      failure.code = data?.error || '';
      failure.status = response.status;
      throw failure;
    }
    return data;
  }

  // ---------- 與主機同步 ----------
  function applyState(state) {
    db = {
      event: state.event && typeof state.event === 'object' ? state.event : seed().event,
      people: Array.isArray(state.people) ? state.people : [],
      tasks: Array.isArray(state.tasks) ? state.tasks : []
    };
    normalizePeopleReferences(db);
    snapshot = {
      event: { rev: state.eventRev || 0, json: JSON.stringify(db.event) },
      people: new Map(db.people.map((record, index) => [record.id, { rev: state.peopleRevs[index], json: JSON.stringify(record) }])),
      tasks: new Map(db.tasks.map((record, index) => [record.id, { rev: state.tasksRevs[index], json: JSON.stringify(record) }]))
    };
    MU.version = state.version;
    stale = false;
  }

  async function loadState() {
    applyState(await api('GET', '/api/state'));
  }

  function buildOps() {
    if (MU.forceReplace) {
      return [{
        op: 'replaceAll', id: '', rev: 0, record: JSON.stringify(db.event),
        people: db.people.map(record => JSON.stringify(record)),
        tasks: db.tasks.map(record => JSON.stringify(record))
      }];
    }
    const ops = [];
    for (const [kind, list, known] of [['Person', db.people, snapshot.people], ['Task', db.tasks, snapshot.tasks]]) {
      const present = new Set();
      for (const record of list) {
        const id = String(record.id), text = JSON.stringify(record), before = known.get(record.id);
        present.add(record.id);
        if (!before) ops.push({ op: 'put' + kind, id, rev: 0, record: text });
        else if (before.json !== text) ops.push({ op: 'put' + kind, id, rev: before.rev, record: text });
      }
      for (const [id, before] of known) {
        if (!present.has(id)) ops.push({ op: 'del' + kind, id: String(id), rev: before.rev });
      }
    }
    const eventText = JSON.stringify(db.event);
    if (eventText !== snapshot.event.json) ops.push({ op: 'setEvent', id: '', rev: snapshot.event.rev, record: eventText });
    return ops;
  }

  function rememberSent(ops) {
    for (const op of ops) {
      const known = op.op.endsWith('Person') ? snapshot.people : snapshot.tasks;
      if (op.op === 'setEvent') snapshot.event = { rev: op.rev + 1, json: op.record };
      else if (op.op.startsWith('put')) known.set(op.id, { rev: op.rev + 1, json: op.record });
      else if (op.op.startsWith('del')) known.delete(op.id);
    }
  }

  async function flush() {
    if (!MU.me) return;
    const replacing = MU.forceReplace, ops = buildOps();
    MU.forceReplace = false;
    if (!ops.length) return;
    busy++;
    try {
      const result = await api('POST', '/api/ops', { ops });
      if (replacing) {
        await loadState();
        softRefresh();
      } else {
        rememberSent(ops);
        if (result.version !== MU.version + 1) stale = true;
        MU.version = result.version;
      }
    } catch (error) {
      toast(message(error));
      if (error.status === 401) { showAuth('login'); return; }
      try { await loadState(); softRefresh(); } catch (reloadError) { setOnline(false); }
    } finally {
      busy--;
    }
  }

  MU.push = function () {
    chain = chain.then(flush, flush);
    return chain;
  };

  function canRefresh() {
    if (busy || document.querySelector('dialog[open]') || pendingWorkImport) return false;
    const focused = document.activeElement;
    return !(focused && content.contains(focused) && focused.tagName === 'INPUT');
  }

  function softRefresh() {
    const values = {};
    content.querySelectorAll('input[id],select[id]').forEach(control => { values[control.id] = control.value; });
    $('#eventLabel').textContent = `${db.event.name || '晁人'}・多人共用`;
    render(values.leaderSelect || '', values.leaderDate || '');
    for (const [id, value] of Object.entries(values)) {
      const control = document.getElementById(id);
      if (!control || id === 'peopleSort') continue;
      if (control.tagName === 'SELECT' && !Array.from(control.options).some(option => option.value === value)) continue;
      control.value = value;
    }
    if (page === 'tasks') applyTaskFilters();
    else if (page === 'people') applyPeopleFilters();
    else if (page === 'shifts' || (page === 'dashboard' && ['filterDate', 'filterTime', 'filterTask', 'filterLeader'].some(id => values[id]))) applyFilters();
    else if (page === 'personal' && values.personSelect) $('#personalResult').innerHTML = personalTable(values.personSelect, values.personalDate || '');
  }

  function setOnline(online) {
    document.body.classList.toggle('mu-offline', !online);
  }

  async function tick() {
    if (!MU.me || busy || polling || document.hidden) return;
    polling = true;
    try {
      const info = await api('GET', '/api/version');
      setOnline(true);
      setMe(info.user);
      if ((info.version !== MU.version || stale) && canRefresh()) {
        const state = await api('GET', '/api/state');
        if (canRefresh()) { applyState(state); softRefresh(); }
      }
    } catch (error) {
      if (error.status === 401) showAuth('login');
      else setOnline(false);
    } finally {
      polling = false;
    }
  }

  // ---------- 身分與權限 ----------
  function setMe(user) {
    const roleChanged = MU.me && user && MU.me.role !== user.role;
    MU.me = user;
    document.body.dataset.role = user?.role || '';
    $('#userEmail').textContent = user ? `${user.name}｜${roleLabel(user.role)}` : '';
    if (user && !ROLE_LEVELS[user.role]) { showAuth('pending'); return; }
    if (roleChanged) {
      toast('你的權限已調整為「' + roleLabel(user.role) + '」');
      if (page === 'users' && user.role !== 'admin') nav('dashboard');
      else if (canRefresh()) softRefresh();
    }
  }

  function blocked(event) {
    event.preventDefault();
    event.stopImmediatePropagation();
    toast(ERRORS.forbidden);
  }

  document.addEventListener('click', event => {
    if (!MU.me) return;
    const button = event.target.closest('[data-action],#settingsBtn,#restoreBtn,#resetDataBtn');
    if (!button) return;
    const action = button.dataset.action;
    const needed = button.id === 'restoreBtn' || button.id === 'resetDataBtn' ? 4
      : DELETE_ACTIONS.has(action) ? 3
        : EDIT_ACTIONS.has(action) || button.id === 'settingsBtn' ? 2 : 0;
    if (level() < needed) blocked(event);
  }, true);

  document.addEventListener('change', event => {
    if (event.target.matches?.('[data-status-id]') && level() < 2) { blocked(event); softRefresh(); }
  }, true);

  document.addEventListener('submit', event => {
    if (event.target.id === 'form' && level() < 2) blocked(event);
  }, true);

  // ---------- 登入畫面 ----------
  const authModes = {
    setup: {
      title: '第一次使用：建立管理者',
      note: '這個帳號可以新增其他使用者並設定權限。密碼請自己記好，系統查不到原本的密碼。',
      fields: [['username', '管理者帳號', 'text', 'username'], ['name', '顯示姓名', 'text', 'name'], ['password', '密碼（至少 6 個字）', 'password', 'new-password'], ['confirm', '再輸入一次密碼', 'password', 'new-password']],
      submit: '建立並登入'
    },
    login: {
      title: '登入排班系統',
      note: '請輸入管理者發給你的帳號與密碼。',
      fields: [['username', '帳號', 'text', 'username'], ['password', '密碼', 'password', 'current-password']],
      submit: '登入'
    },
    setupRemote: { title: '系統尚未完成設定', note: '請先在主機那台電腦上開啟排班系統，建立管理者帳號，再回到這裡重新整理。', fields: [], submit: '重新整理' },
    offline: { title: '連不上主機', note: '請確認主機電腦上的排班系統視窗還開著，而且你和主機連在同一個網路。', fields: [], submit: '重新整理' }
  };
  if (CLOUD) {
    const email = ['username', 'Email', 'email', 'username'], newPassword = ['password', '密碼（至少 6 個字）', 'password', 'new-password'], again = ['confirm', '再輸入一次密碼', 'password', 'new-password'];
    Object.assign(authModes, {
      setup: { title: '第一次使用：建立管理者', note: '第一個註冊的人會成為管理者，之後由你開通其他人的權限。', fields: [email, ['name', '顯示姓名', 'text', 'name'], newPassword, again], submit: '建立管理者', endpoint: '/api/setup' },
      login: { title: '登入排班系統', note: '請輸入你註冊的 Email 與密碼。', fields: [email, ['password', '密碼', 'password', 'current-password']], submit: '登入', links: [['register', '註冊新帳號'], ['forgot', '忘記密碼']] },
      register: { title: '註冊新帳號', note: '註冊後需要管理者開通權限，才能進入排班系統。', fields: [email, ['name', '顯示姓名', 'text', 'name'], newPassword, again], submit: '註冊', endpoint: '/api/register', links: [['login', '已有帳號，回到登入']] },
      forgot: { title: '忘記密碼', note: '輸入註冊的 Email，系統會寄一封重設密碼的信給你。', fields: [email], submit: '寄出重設信', endpoint: '/api/forgot', links: [['login', '回到登入']] },
      recovery: { title: '設定新密碼', note: '請輸入新的密碼。', fields: [newPassword, again], submit: '儲存新密碼', endpoint: '/api/recover' },
      pending: { title: '等待管理者開通', note: '你的帳號已註冊，請通知管理者到「使用者管理」幫你設定權限，再回來按重新整理。', fields: [], submit: '重新整理', links: [['logout', '改用其他帳號登入']] },
      offline: { title: '連不上雲端', note: '請確認這台裝置的網路正常，再按重新整理。', fields: [], submit: '重新整理' },
      notConfigured: { title: '尚未完成雲端設定', note: '請先依照 README 的步驟，把 Supabase 的網址與金鑰填進 config.js。', fields: [], submit: '重新整理' }
    });
  }
  let authMode = 'offline';

  function showAuth(mode) {
    authMode = mode;
    if (mode === 'login' || mode === 'offline') { MU.me = null; document.body.dataset.role = ''; }
    const config = authModes[mode];
    document.querySelectorAll('dialog[open]').forEach(dialog => dialog.close());
    $('#muAuthTitle').textContent = config.title;
    $('#muAuthNote').textContent = config.note;
    $('#muAuthFields').innerHTML = config.fields.map(([name, label, type, autocomplete]) =>
      `<label class="field"><span>${label}</span><input name="${name}" type="${type}" autocomplete="${autocomplete}" ${name === 'name' ? '' : 'required'}></label>`).join('');
    $('#muAuthSubmit').textContent = config.submit;
    $('#muAuthLinks').innerHTML = (config.links || []).map(([target, label]) => `<button type="button" class="subtle" data-mu-auth="${target}">${label}</button>`).join('');
    $('#muAuthError').textContent = '';
    document.body.classList.add('mu-locked');
    $('#muAuthFields input')?.focus();
  }

  async function enter(user) {
    setMe(user);
    if (!ROLE_LEVELS[user.role]) return;
    await loadState();
    document.body.classList.remove('mu-locked');
    setOnline(true);
    $('#eventLabel').textContent = `${db.event.name || '晁人'}・多人共用`;
    if (page === 'users' && user.role !== 'admin') nav('dashboard'); else render();
    if (!pollTimer) pollTimer = setInterval(tick, CLOUD ? 5000 : 3000);
    offerLegacyImport();
  }

  async function submitAuth(event) {
    event.preventDefault();
    const config = authModes[authMode];
    if (!config.fields.length) { location.reload(); return; }
    const data = Object.fromEntries(new FormData(event.currentTarget).entries());
    const errorBox = $('#muAuthError');
    errorBox.textContent = '';
    if ('confirm' in data && data.password !== data.confirm) { errorBox.textContent = '兩次輸入的密碼不一樣'; return; }
    $('#muAuthSubmit').disabled = true;
    try {
      const endpoint = config.endpoint || (authMode === 'setup' ? '/api/setup' : '/api/login');
      const result = await api('POST', endpoint, { username: (data.username || '').trim(), name: data.name || '', password: data.password || '' });
      if (result.notice) { showAuth('login'); errorBox.textContent = ERRORS[result.notice] || ''; }
      else await enter(result.user);
    } catch (error) {
      errorBox.textContent = message(error);
    } finally {
      $('#muAuthSubmit').disabled = false;
    }
  }

  // 主機第一次改用多人版時，把同一個瀏覽器裡的單機版資料帶過來。
  function offerLegacyImport() {
    if (MU.me?.role !== 'admin' || MU.version !== 0 || db.people.length || db.tasks.length) return;
    let legacy = null;
    try { legacy = JSON.parse(localStorage.getItem(LEGACY_KEY) || 'null'); } catch (error) { return; }
    if (!legacy || !Array.isArray(legacy.people) || !Array.isArray(legacy.tasks) || !(legacy.people.length || legacy.tasks.length)) return;
    if (!confirm(`這台電腦的瀏覽器裡有單機版的資料（${legacy.people.length} 位服務人員、${legacy.tasks.length} 項工作）。\n\n要把它帶進多人版，變成大家共用的資料嗎？\n單機版原本的資料不會被刪除。`)) return;
    db = { event: legacy.event && typeof legacy.event === 'object' ? legacy.event : seed().event, people: legacy.people, tasks: legacy.tasks };
    normalizePeopleReferences(db);
    MU.forceReplace = true;
    save();
    toast('單機版資料已帶入多人版');
  }

  // ---------- 使用者管理（管理者） ----------
  titles.users = ['系統管理', '使用者管理', '設定誰可以閱覽、編輯、刪除'];

  async function renderUsers() {
    $('#headActions').innerHTML = CLOUD ? '' : '<button class="btn primary" data-mu="new-user">＋ 新增使用者</button>';
    content.innerHTML = `${CLOUD ? '<div class="notice"><b>怎麼新增使用者</b><div>請對方打開這個網址，按「註冊新帳號」。註冊後會出現在下方，狀態是「待開通」；你按「編輯」選好權限，對方就能進來。</div></div>' : ''}<div class="notice"><b>四種權限</b>${Object.keys(ROLE_LABELS).map(role => `<div>・${ROLE_LABELS[role]}：${ROLE_NOTES[role]}</div>`).join('')}</div><div class="table-wrap"><table><thead><tr><th>帳號</th><th>姓名</th><th>權限</th><th>狀態</th><th></th></tr></thead><tbody id="muUserRows"><tr><td colspan="5" class="empty">讀取中…</td></tr></tbody></table></div>`;
    try {
      const { users } = await api('GET', '/api/users');
      MU.users = users;
      if (page !== 'users') return;
      $('#muUserRows').innerHTML = users.map(user => `<tr><td class="table-title">${esc(user.username)}${user.username === MU.me.username ? ' <span class="badge ok">你</span>' : ''}</td><td>${esc(user.name)}</td><td>${user.role === 'pending' ? '<span class="badge short">待開通</span>' : esc(roleLabel(user.role))}</td><td><span class="badge ${user.active ? 'ok' : 'short'}">${user.active ? '啟用' : '停用'}</span></td><td><div class="person-row-actions"><button class="subtle" data-mu="edit-user" data-username="${esc(user.username)}">編輯</button>${user.username === MU.me.username ? '' : ` <button class="subtle" data-mu="delete-user" data-username="${esc(user.username)}">刪除</button>`}</div></td></tr>`).join('');
    } catch (error) {
      if (page === 'users') $('#muUserRows').innerHTML = `<tr><td colspan="5" class="empty">${esc(message(error))}</td></tr>`;
    }
  }

  const baseRender = render;
  render = function (selectedLeader, selectedDate) {
    baseRender(selectedLeader, selectedDate);
    if (page === 'users') renderUsers();
  };

  function openUserDialog(user) {
    editingUser = user || null;
    $('#muUserTitle').textContent = user ? '編輯使用者' : '新增使用者';
    const form = $('#muUserForm');
    form.reset();
    form.username.value = user?.username || '';
    form.username.readOnly = Boolean(user);
    form.name.value = user?.name || '';
    form.role.value = ROLE_LEVELS[user?.role] ? user.role : 'viewer';
    form.active.value = user && !user.active ? 'false' : 'true';
    form.password.required = !user && !CLOUD;
    form.password.closest('label').hidden = CLOUD;
    $('#muUserPasswordLabel').textContent = user ? '重設密碼（不改就留空白）' : '初始密碼（至少 6 個字）';
    $('#muUserError').textContent = '';
    $('#muUserDialog').showModal();
  }

  async function submitUser(event) {
    event.preventDefault();
    const form = event.currentTarget;
    try {
      await api('POST', '/api/users/save', {
        isNew: !editingUser, username: form.username.value.trim(), name: form.name.value.trim(),
        role: form.role.value, active: form.active.value === 'true', password: form.password.value
      });
      $('#muUserDialog').close();
      toast(editingUser ? '使用者已更新' : '使用者已新增，請把帳號與密碼告訴對方');
      renderUsers();
    } catch (error) {
      $('#muUserError').textContent = message(error);
    }
  }

  async function submitPassword(event) {
    event.preventDefault();
    const form = event.currentTarget, errorBox = $('#muPasswordError');
    if (form.newPassword.value !== form.confirm.value) { errorBox.textContent = '兩次輸入的新密碼不一樣'; return; }
    try {
      await api('POST', '/api/password', { oldPassword: form.oldPassword.value, newPassword: form.newPassword.value });
      $('#muPasswordDialog').close();
      toast('密碼已更新');
    } catch (error) {
      errorBox.textContent = message(error);
    }
  }

  document.addEventListener('click', async event => {
    const link = event.target.closest('[data-mu-auth]');
    if (link) {
      if (link.dataset.muAuth === 'logout') { try { await api('POST', '/api/logout'); } catch (error) { /* 仍回到登入畫面 */ } showAuth('login'); }
      else showAuth(link.dataset.muAuth);
      return;
    }
    const button = event.target.closest('[data-mu]');
    if (!button) return;
    const action = button.dataset.mu, target = (MU.users || []).find(user => user.username === button.dataset.username);
    if (action === 'new-user') openUserDialog(null);
    if (action === 'edit-user' && target) openUserDialog(target);
    if (action === 'close-dialog') button.closest('dialog').close();
    if (action === 'change-password') { $('#muPasswordForm').reset(); $('#muPasswordError').textContent = ''; $('#muPasswordDialog').showModal(); }
    if (action === 'logout') {
      try { await api('POST', '/api/logout'); } catch (error) { /* 主機離線時仍回到登入畫面 */ }
      showAuth('login');
    }
    if (action === 'delete-user' && target) {
      if (!confirm(`確定要刪除使用者「${target.name}」（${target.username}）嗎？\n刪除後對方將無法登入；排班資料不受影響。`)) return;
      try { await api('POST', '/api/users/delete', { username: target.username }); toast('使用者已刪除'); }
      catch (error) { toast(message(error)); }
      renderUsers();
    }
  });

  // ---------- 畫面元件 ----------
  document.body.insertAdjacentHTML('beforeend', `
    <div id="muAuth"><form id="muAuthForm" class="mu-auth-card">
      <div class="brand"><span class="logo">晁</span><div><b>晁人</b><small>服務人員活動排班系統・多人版</small></div></div>
      <h2 id="muAuthTitle">連線中…</h2><p id="muAuthNote"></p>
      <div id="muAuthFields" class="mu-form-grid"></div>
      <div id="muAuthError" class="mu-error" role="alert"></div>
      <button type="submit" class="btn primary" id="muAuthSubmit">登入</button>
      <div id="muAuthLinks" class="mu-auth-links"></div>
    </form></div>
    <dialog id="muUserDialog"><form id="muUserForm">
      <div class="modal-head"><div><div class="eyebrow">使用者</div><h2 id="muUserTitle">新增使用者</h2></div><button type="button" class="icon" data-mu="close-dialog" aria-label="關閉">✕</button></div>
      <div class="mu-form-grid two">
        <label class="field"><span>帳號</span><input name="username" autocomplete="off" required></label>
        <label class="field"><span>姓名</span><input name="name" autocomplete="off"></label>
        <label class="field"><span>權限</span><select name="role">${Object.keys(ROLE_LABELS).map(role => `<option value="${role}">${ROLE_LABELS[role]}</option>`).join('')}</select></label>
        <label class="field"><span>狀態</span><select name="active"><option value="true">啟用</option><option value="false">停用（不能登入）</option></select></label>
        <label class="field full"><span id="muUserPasswordLabel">初始密碼</span><input name="password" type="password" autocomplete="new-password" minlength="6"></label>
      </div>
      <div id="muUserError" class="mu-error" role="alert"></div>
      <div class="form-actions"><button type="button" class="btn light" data-mu="close-dialog">取消</button><button type="submit" class="btn primary">儲存</button></div>
    </form></dialog>
    <dialog id="muPasswordDialog"><form id="muPasswordForm">
      <div class="modal-head"><div><div class="eyebrow">我的帳號</div><h2>修改密碼</h2></div><button type="button" class="icon" data-mu="close-dialog" aria-label="關閉">✕</button></div>
      <div class="mu-form-grid">
        <label class="field"><span>目前的密碼</span><input name="oldPassword" type="password" autocomplete="current-password" required></label>
        <label class="field"><span>新密碼（至少 6 個字）</span><input name="newPassword" type="password" autocomplete="new-password" minlength="6" required></label>
        <label class="field"><span>再輸入一次新密碼</span><input name="confirm" type="password" autocomplete="new-password" minlength="6" required></label>
      </div>
      <div id="muPasswordError" class="mu-error" role="alert"></div>
      <div class="form-actions"><button type="button" class="btn light" data-mu="close-dialog">取消</button><button type="submit" class="btn primary">更新密碼</button></div>
    </form></dialog>`);
  $('.account-area').insertAdjacentHTML('beforeend', '<button type="button" class="logout-btn" data-mu="change-password">改密碼</button><button type="button" class="logout-btn" id="logoutBtn" data-mu="logout">登出</button>');
  $('#muAuthForm').addEventListener('submit', submitAuth);
  $('#muUserForm').addEventListener('submit', submitUser);
  $('#muPasswordForm').addEventListener('submit', submitPassword);

  // ---------- 啟動 ----------
  (async function boot() {
    // 使用者點了「重設密碼」信裡的連結回來時，先請他設定新密碼。
    MU.onRecovery = () => showAuth('recovery');
    try {
      const session = await api('GET', '/api/session');
      if (session.recovery) showAuth('recovery');
      else if (session.user) await enter(session.user);
      else showAuth(session.setupNeeded ? (session.isLocal ? 'setup' : 'setupRemote') : 'login');
    } catch (error) {
      showAuth(error.code === 'not_configured' ? 'notConfigured' : 'offline');
    }
  })();
})();
