// Google Contacts Editor — App Logic

/* ── State ── */
let accessToken = null;
let tokenClient = null;
let allContacts = [];
let filteredContacts = [];
let contactGroups = {};  // id → name
let dirtySet = new Set(); // resourceNames with unsaved edits
let sortCol = 'name';
let sortAsc = true;
let pendingDeleteRN = null;

/* ── Enrichment state ── */
let enrichSelectedContacts = [];
let enrichResults = [];       // parsed AI results
let enrichCardStates = {};    // { resourceName: { status: 'pending'|'approved'|'skipped', fields: { fieldName: { checked, confidence, source, value } } } }
let currentTab = 'contacts';
let tokenRefreshPromise = null;

/* ── DOM refs ── */
const $ = id => document.getElementById(id);
const authBtn = $('auth-btn');
const searchInput = $('search-input');
const newContactBtn = $('new-contact-btn');
const tableContainer = $('table-container');
const signInPrompt = $('sign-in-prompt');
const contactCount = $('contact-count');
const syncStatus = $('sync-status');
const deleteModal = $('delete-modal');
const deleteModalMsg = $('delete-modal-msg');
const deleteCancelBtn = $('delete-cancel-btn');
const deleteConfirmBtn = $('delete-confirm-btn');
const toastContainer = $('toast-container');
const detailView = $('detail-view');
const tabNav = $('tab-nav');
const enrichView = $('enrich-view');

/* ── Column config ── */
const COLUMNS = [
  { key: 'name',     label: 'Name' },
  { key: 'email',    label: 'Email' },
  { key: 'phone',    label: 'Phone' },
  { key: 'company',  label: 'Company' },
  { key: 'title',    label: 'Title' },
  { key: 'address',  label: 'Address' },
  { key: 'birthday', label: 'Birthday' },
  { key: 'website',  label: 'Website' },
  { key: 'nickname', label: 'Nickname' },
  { key: 'relation', label: 'Relations' },
  { key: 'event',    label: 'Events' },
  { key: 'im',       label: 'IMs' },
  { key: 'groups',   label: 'Groups' },
  { key: 'notes',    label: 'Notes' },
];

/* ══════════════════════════════════════════
   AUTH
   ══════════════════════════════════════════ */

function initAuth() {
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: SCOPES,
    callback: onTokenResponse,
  });
  authBtn.onclick = () => {
    if (accessToken) signOut();
    else tokenClient.requestAccessToken();
  };
}

function onTokenResponse(resp) {
  if (resp.error) {
    if (tokenRefreshPromise) { tokenRefreshPromise.reject(new Error('Auth failed: ' + resp.error)); tokenRefreshPromise = null; }
    toast('Auth failed: ' + resp.error, 'error');
    return;
  }
  accessToken = resp.access_token;
  if (tokenRefreshPromise) { tokenRefreshPromise.resolve(accessToken); tokenRefreshPromise = null; return; }
  authBtn.textContent = 'Sign out';
  signInPrompt.style.display = 'none';
  tableContainer.style.display = '';
  newContactBtn.style.display = '';
  tabNav.classList.add('visible');
  loadAll();
}

function refreshToken() {
  if (tokenRefreshPromise) return tokenRefreshPromise.promise;
  let resolve, reject;
  const promise = new Promise((res, rej) => { resolve = res; reject = rej; });
  tokenRefreshPromise = { promise, resolve, reject };
  tokenClient.requestAccessToken();
  return promise;
}

function signOut() {
  if (accessToken) google.accounts.oauth2.revoke(accessToken);
  accessToken = null;
  allContacts = [];
  filteredContacts = [];
  dirtySet.clear();
  authBtn.textContent = 'Sign in';
  signInPrompt.style.display = '';
  tableContainer.style.display = 'none';
  newContactBtn.style.display = 'none';
  contactCount.textContent = '';
  syncStatus.textContent = '';
  tableContainer.innerHTML = '';
  tabNav.classList.remove('visible');
  switchTab('contacts');
}

/* ── API helper with transparent token refresh ── */
async function apiFetch(url, opts = {}, _retried = false) {
  const headers = { Authorization: 'Bearer ' + accessToken, ...opts.headers };
  const res = await fetch(url, { ...opts, headers });
  if (res.status === 401 && !_retried) {
    await refreshToken();
    return apiFetch(url, opts, true);
  }
  return res;
}

/* ══════════════════════════════════════════
   LOAD DATA
   ══════════════════════════════════════════ */

async function loadAll() {
  setSyncStatus('Loading...', 'saving');
  try {
    await Promise.all([fetchContactGroups(), loadContacts()]);
    applyFilterAndSort();
    setSyncStatus('Loaded ' + allContacts.length + ' contacts', 'success');
    setTimeout(() => { if (syncStatus.classList.contains('success')) syncStatus.textContent = ''; }, 3000);
  } catch (e) {
    setSyncStatus('Load failed', 'error');
    toast('Failed to load contacts: ' + e.message, 'error');
  }
}

async function fetchContactGroups() {
  const res = await apiFetch('https://people.googleapis.com/v1/contactGroups?pageSize=100');
  const data = await res.json();
  contactGroups = {};
  (data.contactGroups || []).forEach(g => {
    if (g.groupType === 'CONTACT_GROUP' || g.groupType === 'SYSTEM_CONTACT_GROUP') {
      contactGroups[g.resourceName] = g.name || g.formattedName || '';
    }
  });
}

async function loadContacts() {
  allContacts = [];
  let pageToken = '';
  do {
    const url = 'https://people.googleapis.com/v1/people/me/connections'
      + '?personFields=' + PERSON_FIELDS
      + '&pageSize=1000'
      + '&sortOrder=FIRST_NAME_ASCENDING'
      + (pageToken ? '&pageToken=' + pageToken : '');
    const res = await apiFetch(url);
    const data = await res.json();
    if (data.connections) allContacts.push(...data.connections);
    pageToken = data.nextPageToken || '';
  } while (pageToken);
}

/* ══════════════════════════════════════════
   DATA HELPERS
   ══════════════════════════════════════════ */

function getPrimaryName(p) {
  const n = (p.names && p.names[0]) || {};
  return { display: n.displayName || '', given: n.givenName || '', family: n.familyName || '' };
}

function getPrimaryEmail(p) {
  const e = (p.emailAddresses && p.emailAddresses[0]) || {};
  return e.value || '';
}

function getPrimaryPhone(p) {
  const ph = (p.phoneNumbers && p.phoneNumbers[0]) || {};
  return ph.value || '';
}

function getOrganization(p) {
  const o = (p.organizations && p.organizations[0]) || {};
  return { company: o.name || '', title: o.title || '' };
}

function getPrimaryAddress(p) {
  const a = (p.addresses && p.addresses[0]) || {};
  return a.formattedValue || '';
}

function getBirthday(p) {
  const b = (p.birthdays && p.birthdays[0]) || {};
  if (b.date) {
    const d = b.date;
    if (d.year) return `${d.year}-${String(d.month).padStart(2,'0')}-${String(d.day).padStart(2,'0')}`;
    return `${String(d.month).padStart(2,'0')}-${String(d.day).padStart(2,'0')}`;
  }
  return b.text || '';
}

function getNotes(p) {
  const bio = (p.biographies && p.biographies[0]) || {};
  return bio.value || '';
}

function getPrimaryWebsite(p) {
  const u = (p.urls && p.urls[0]) || {};
  return u.value || '';
}

function getNickname(p) {
  const n = (p.nicknames && p.nicknames[0]) || {};
  return n.value || '';
}

function getRelations(p) {
  if (!p.relations || !p.relations.length) return '';
  return p.relations.map(r => {
    const label = r.type ? ` (${r.type})` : '';
    return (r.person || '') + label;
  }).join(', ');
}

function getEvents(p) {
  if (!p.events || !p.events.length) return '';
  return p.events.map(ev => {
    const d = ev.date || {};
    const dateStr = d.year
      ? `${d.year}-${String(d.month).padStart(2,'0')}-${String(d.day).padStart(2,'0')}`
      : `${String(d.month||0).padStart(2,'0')}-${String(d.day||0).padStart(2,'0')}`;
    const label = ev.type ? ` (${ev.type})` : '';
    return dateStr + label;
  }).join(', ');
}

function getIMs(p) {
  if (!p.imClients || !p.imClients.length) return '';
  return p.imClients.map(im => {
    const label = im.protocol ? ` (${im.protocol})` : '';
    return (im.username || '') + label;
  }).join(', ');
}

function getGroups(p) {
  if (!p.memberships) return '';
  return p.memberships
    .filter(m => m.contactGroupMembership)
    .map(m => contactGroups[m.contactGroupMembership.contactGroupResourceName] || '')
    .filter(n => n && n !== 'myContacts')
    .join(', ');
}

function getFieldValue(p, key) {
  switch (key) {
    case 'name':     return getPrimaryName(p).display;
    case 'email':    return getPrimaryEmail(p);
    case 'phone':    return getPrimaryPhone(p);
    case 'company':  return getOrganization(p).company;
    case 'title':    return getOrganization(p).title;
    case 'address':  return getPrimaryAddress(p);
    case 'birthday': return getBirthday(p);
    case 'website':  return getPrimaryWebsite(p);
    case 'nickname': return getNickname(p);
    case 'relation': return getRelations(p);
    case 'event':    return getEvents(p);
    case 'im':       return getIMs(p);
    case 'groups':   return getGroups(p);
    case 'notes':    return getNotes(p);
    default: return '';
  }
}

function multiValueCount(p, key) {
  switch (key) {
    case 'email':    return (p.emailAddresses || []).length;
    case 'phone':    return (p.phoneNumbers || []).length;
    case 'address':  return (p.addresses || []).length;
    case 'website':  return (p.urls || []).length;
    case 'relation': return (p.relations || []).length;
    case 'event':    return (p.events || []).length;
    case 'im':       return (p.imClients || []).length;
    default: return 0;
  }
}

/* ══════════════════════════════════════════
   SEARCH + SORT
   ══════════════════════════════════════════ */

function applyFilterAndSort() {
  const q = searchInput.value.trim().toLowerCase();
  filteredContacts = allContacts.filter(p => {
    if (!q) return true;
    return COLUMNS.some(c => getFieldValue(p, c.key).toLowerCase().includes(q));
  });
  filteredContacts.sort((a, b) => {
    const va = getFieldValue(a, sortCol).toLowerCase();
    const vb = getFieldValue(b, sortCol).toLowerCase();
    if (va < vb) return sortAsc ? -1 : 1;
    if (va > vb) return sortAsc ? 1 : -1;
    return 0;
  });
  contactCount.textContent = filteredContacts.length + (filteredContacts.length !== allContacts.length ? ' / ' + allContacts.length : '') + ' contacts';
  renderTable();
}

searchInput.addEventListener('input', () => applyFilterAndSort());

/* ══════════════════════════════════════════
   RENDER TABLE
   ══════════════════════════════════════════ */

const RENDER_BATCH = 100;
let renderedCount = 0;

function renderTable() {
  if (filteredContacts.length === 0 && allContacts.length === 0) {
    tableContainer.innerHTML = '<div class="empty-state">No contacts found. Click "+ New Contact" to add one.</div>';
    return;
  }
  if (filteredContacts.length === 0) {
    tableContainer.innerHTML = '<div class="empty-state">No contacts match your search.</div>';
    return;
  }

  let html = '<table class="contacts-table"><colgroup>';
  COLUMNS.forEach(c => { html += `<col class="col-${c.key}">`; });
  html += '<col class="col-actions"></colgroup><thead><tr>';
  COLUMNS.forEach(c => {
    const arrow = sortCol === c.key ? (sortAsc ? ' ▲' : ' ▼') : '';
    const active = sortCol === c.key ? ' active' : '';
    html += `<th data-col="${c.key}">${esc(c.label)}<span class="sort-arrow${active}">${arrow}</span></th>`;
  });
  html += '<th>Actions</th></tr></thead><tbody>';

  renderedCount = Math.min(RENDER_BATCH, filteredContacts.length);
  for (let i = 0; i < renderedCount; i++) {
    html += buildRow(filteredContacts[i]);
  }

  html += '</tbody></table>';
  if (renderedCount < filteredContacts.length) {
    html += `<div class="load-more" id="load-more" style="text-align:center;padding:12px;color:#1a73e8;cursor:pointer;font-size:13px;">Show more (${filteredContacts.length - renderedCount} remaining)</div>`;
  }
  tableContainer.innerHTML = html;
  bindTableEvents();
  bindLoadMore();
}

function buildRow(p) {
  const rn = p.resourceName;
  const dirty = dirtySet.has(rn);
  let row = `<tr data-rn="${esc(rn)}" class="${dirty ? 'dirty' : ''}">`;
  COLUMNS.forEach(c => {
    const val = getFieldValue(p, c.key);
    const mvc = multiValueCount(p, c.key);
    const badge = mvc > 1 ? `<span class="multi-badge" data-col="${c.key}">+${mvc - 1}</span>` : '';
    const editable = c.key !== 'groups' ? ` data-col="${c.key}"` : '';
    row += `<td${editable}>${esc(val)}${badge}</td>`;
  });
  row += '<td>';
  if (dirty) row += `<button class="btn-icon save-btn" title="Save" data-action="save" data-rn="${esc(rn)}">✓</button>`;
  row += `<button class="btn-icon delete-btn" title="Delete" data-action="delete" data-rn="${esc(rn)}">🗑</button>`;
  row += '</td></tr>';
  return row;
}

function loadMoreRows() {
  const tbody = tableContainer.querySelector('tbody');
  if (!tbody) return;
  const next = Math.min(renderedCount + RENDER_BATCH, filteredContacts.length);
  let html = '';
  for (let i = renderedCount; i < next; i++) {
    html += buildRow(filteredContacts[i]);
  }
  tbody.insertAdjacentHTML('beforeend', html);
  renderedCount = next;
  bindTableEvents();
  const loadMoreEl = $('load-more');
  if (loadMoreEl) {
    if (renderedCount >= filteredContacts.length) loadMoreEl.remove();
    else loadMoreEl.textContent = `Show more (${filteredContacts.length - renderedCount} remaining)`;
  }
}

function bindLoadMore() {
  const el = $('load-more');
  if (el) el.addEventListener('click', loadMoreRows);
  // Also auto-load on scroll near bottom
  $('main-content').onscroll = () => {
    if (renderedCount >= filteredContacts.length) return;
    const mc = $('main-content');
    if (mc.scrollTop + mc.clientHeight >= mc.scrollHeight - 200) {
      loadMoreRows();
    }
  };
}

function esc(s) {
  const d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}

/* ══════════════════════════════════════════
   TABLE EVENTS
   ══════════════════════════════════════════ */

function bindTableEvents() {
  // Column header sort
  tableContainer.querySelectorAll('th[data-col]').forEach(th => {
    th.onclick = () => {
      const col = th.dataset.col;
      if (sortCol === col) sortAsc = !sortAsc;
      else { sortCol = col; sortAsc = true; }
      applyFilterAndSort();
    };
  });

  // Row click → open detail view
  tableContainer.querySelectorAll('tbody tr').forEach(tr => {
    tr.addEventListener('click', e => {
      // Don't open detail if clicking action buttons or editing
      if (e.target.closest('[data-action]') || e.target.classList.contains('cell-edit')) return;
      const rn = tr.dataset.rn;
      const person = allContacts.find(p => p.resourceName === rn);
      if (person) openDetailView(person);
    });
  });

  // Action buttons
  tableContainer.querySelectorAll('[data-action]').forEach(btn => {
    btn.onclick = e => {
      e.stopPropagation();
      const rn = btn.dataset.rn;
      if (btn.dataset.action === 'save') saveContact(rn);
      if (btn.dataset.action === 'delete') confirmDelete(rn);
    };
  });
}

/* ══════════════════════════════════════════
   CELL EDITING
   ══════════════════════════════════════════ */

function startEditing(td) {
  if (td.querySelector('.cell-edit')) return;
  const col = td.dataset.col;
  if (!col || col === 'groups') return;
  const rn = td.closest('tr').dataset.rn;
  const person = allContacts.find(p => p.resourceName === rn);
  if (!person) return;

  if (col === 'name') { openNameEditor(td, person); return; }
  const multiCols = ['email', 'phone', 'address', 'website', 'relation', 'event', 'im'];
  if (multiCols.includes(col) && multiValueCount(person, col) > 1) {
    openMultiValuePopup(td, person, col);
    return;
  }

  const val = getFieldValue(person, col);
  const isNotes = col === 'notes';

  if (isNotes) {
    const ta = document.createElement('textarea');
    ta.className = 'cell-edit';
    ta.value = val;
    td.textContent = '';
    td.appendChild(ta);
    ta.focus();
    ta.addEventListener('blur', () => commitEdit(td, ta.value, person, col));
    ta.addEventListener('keydown', e => {
      if (e.key === 'Escape') { ta.blur(); }
    });
  } else {
    const inp = document.createElement('input');
    inp.className = 'cell-edit';
    inp.type = col === 'birthday' ? 'date' : 'text';
    inp.value = val;
    td.textContent = '';
    td.appendChild(inp);
    inp.focus();
    inp.addEventListener('blur', () => commitEdit(td, inp.value, person, col));
    inp.addEventListener('keydown', e => {
      if (e.key === 'Enter') inp.blur();
      if (e.key === 'Escape') { inp.value = val; inp.blur(); }
    });
  }
}

function commitEdit(td, newVal, person, col) {
  const oldVal = getFieldValue(person, col);
  if (newVal !== oldVal) {
    setFieldValue(person, col, newVal);
    dirtySet.add(person.resourceName);
  }
  // Re-render the cell
  const mvc = multiValueCount(person, col);
  const badge = mvc > 1 ? `<span class="multi-badge" data-col="${col}">+${mvc - 1}</span>` : '';
  td.innerHTML = esc(getFieldValue(person, col)) + badge;

  // Update row dirty class and save button
  const tr = td.closest('tr');
  if (dirtySet.has(person.resourceName)) {
    tr.classList.add('dirty');
    if (!tr.querySelector('.save-btn')) {
      const saveBtn = document.createElement('button');
      saveBtn.className = 'btn-icon save-btn';
      saveBtn.title = 'Save';
      saveBtn.dataset.action = 'save';
      saveBtn.dataset.rn = person.resourceName;
      saveBtn.textContent = '✓';
      saveBtn.onclick = e => { e.stopPropagation(); saveContact(person.resourceName); };
      tr.querySelector('td:last-child').insertBefore(saveBtn, tr.querySelector('.delete-btn'));
    }
  }

  // Re-bind multi-badge click
  td.querySelectorAll('.multi-badge').forEach(b => {
    b.addEventListener('click', e => {
      e.stopPropagation();
      openMultiValuePopup(td, person, b.dataset.col);
    });
  });
}

function setFieldValue(person, col, val) {
  switch (col) {
    case 'email':
      if (!person.emailAddresses) person.emailAddresses = [];
      if (person.emailAddresses.length === 0) person.emailAddresses.push({});
      person.emailAddresses[0].value = val;
      break;
    case 'phone':
      if (!person.phoneNumbers) person.phoneNumbers = [];
      if (person.phoneNumbers.length === 0) person.phoneNumbers.push({});
      person.phoneNumbers[0].value = val;
      break;
    case 'company': {
      if (!person.organizations) person.organizations = [];
      if (person.organizations.length === 0) person.organizations.push({});
      person.organizations[0].name = val;
      break;
    }
    case 'title': {
      if (!person.organizations) person.organizations = [];
      if (person.organizations.length === 0) person.organizations.push({});
      person.organizations[0].title = val;
      break;
    }
    case 'address':
      if (!person.addresses) person.addresses = [];
      if (person.addresses.length === 0) person.addresses.push({});
      person.addresses[0].formattedValue = val;
      person.addresses[0].streetAddress = val;
      break;
    case 'birthday': {
      if (!person.birthdays) person.birthdays = [];
      if (person.birthdays.length === 0) person.birthdays.push({});
      const parts = val.split('-').map(Number);
      if (parts.length === 3) {
        person.birthdays[0].date = { year: parts[0], month: parts[1], day: parts[2] };
      } else if (parts.length === 2) {
        person.birthdays[0].date = { month: parts[0], day: parts[1] };
      }
      break;
    }
    case 'notes':
      if (!person.biographies) person.biographies = [];
      if (person.biographies.length === 0) person.biographies.push({ contentType: 'TEXT_PLAIN' });
      person.biographies[0].value = val;
      break;
    case 'website':
      if (!person.urls) person.urls = [];
      if (person.urls.length === 0) person.urls.push({});
      person.urls[0].value = val;
      break;
    case 'nickname':
      if (!person.nicknames) person.nicknames = [];
      if (person.nicknames.length === 0) person.nicknames.push({});
      person.nicknames[0].value = val;
      break;
    case 'relation':
      if (!person.relations) person.relations = [];
      if (person.relations.length === 0) person.relations.push({});
      person.relations[0].person = val;
      break;
    case 'event': {
      if (!person.events) person.events = [];
      if (person.events.length === 0) person.events.push({});
      const parts2 = val.split('-').map(Number);
      if (parts2.length === 3) {
        person.events[0].date = { year: parts2[0], month: parts2[1], day: parts2[2] };
      }
      break;
    }
    case 'im':
      if (!person.imClients) person.imClients = [];
      if (person.imClients.length === 0) person.imClients.push({});
      person.imClients[0].username = val;
      break;
  }
}

/* ══════════════════════════════════════════
   NAME EDITOR (popup with given + family)
   ══════════════════════════════════════════ */

function openNameEditor(td, person) {
  closeAllPopups();
  const name = getPrimaryName(person);
  const rect = td.getBoundingClientRect();

  const overlay = document.createElement('div');
  overlay.className = 'mv-popup-overlay';

  const popup = document.createElement('div');
  popup.className = 'mv-popup';
  popup.style.top = rect.bottom + 4 + 'px';
  popup.style.left = rect.left + 'px';

  popup.innerHTML = `
    <h3>Edit Name</h3>
    <div class="name-editor">
      <input type="text" placeholder="First name" value="${esc(name.given)}" data-field="given">
      <input type="text" placeholder="Last name" value="${esc(name.family)}" data-field="family">
    </div>
    <div class="mv-actions">
      <button class="btn btn-secondary mv-close-btn">Cancel</button>
      <button class="btn btn-primary mv-apply-btn">Apply</button>
    </div>
  `;

  overlay.appendChild(popup);
  document.body.appendChild(overlay);
  popup.querySelector('input').focus();

  overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); });
  popup.querySelector('.mv-close-btn').onclick = () => overlay.remove();
  popup.querySelector('.mv-apply-btn').onclick = () => {
    const given = popup.querySelector('[data-field="given"]').value.trim();
    const family = popup.querySelector('[data-field="family"]').value.trim();
    if (!person.names) person.names = [];
    if (person.names.length === 0) person.names.push({});
    person.names[0].givenName = given;
    person.names[0].familyName = family;
    person.names[0].unstructuredName = (given + ' ' + family).trim();
    person.names[0].displayName = person.names[0].unstructuredName;
    dirtySet.add(person.resourceName);
    overlay.remove();
    applyFilterAndSort();
  };

  // Enter to apply
  popup.querySelectorAll('input').forEach(inp => {
    inp.addEventListener('keydown', e => { if (e.key === 'Enter') popup.querySelector('.mv-apply-btn').click(); });
  });
}

/* ══════════════════════════════════════════
   MULTI-VALUE POPUP (email / phone / address)
   ══════════════════════════════════════════ */

function openMultiValuePopup(td, person, col) {
  closeAllPopups();
  const rect = td.getBoundingClientRect();
  const overlay = document.createElement('div');
  overlay.className = 'mv-popup-overlay';

  const popup = document.createElement('div');
  popup.className = 'mv-popup';
  popup.style.top = rect.bottom + 4 + 'px';
  popup.style.left = Math.min(rect.left, window.innerWidth - 440) + 'px';

  const fieldMap = {
    email: 'emailAddresses', phone: 'phoneNumbers', address: 'addresses',
    website: 'urls', relation: 'relations', event: 'events', im: 'imClients'
  };
  const arrKey = fieldMap[col];
  const items = person[arrKey] || [];
  const typeOptions = {
    email: ['home','work','other'],
    phone: ['mobile','home','work','other'],
    address: ['home','work','other'],
    website: ['homepage','blog','work','other'],
    relation: ['spouse','child','parent','sibling','friend','other'],
    event: ['anniversary','other'],
    im: ['aim','hangouts','icq','jabber','msn','qq','skype','yahoo','other'],
  }[col] || ['other'];

  const label = col.charAt(0).toUpperCase() + col.slice(1);

  function getItemValue(item) {
    if (col === 'address') return item.formattedValue || item.streetAddress || '';
    if (col === 'relation') return item.person || '';
    if (col === 'im') return item.username || '';
    if (col === 'event') {
      const d = item.date || {};
      if (d.year) return `${d.year}-${String(d.month).padStart(2,'0')}-${String(d.day).padStart(2,'0')}`;
      if (d.month) return `${String(d.month).padStart(2,'0')}-${String(d.day).padStart(2,'0')}`;
      return '';
    }
    return item.value || '';
  }

  function getItemType(item) {
    if (col === 'im') return (item.protocol || 'other').toLowerCase();
    if (col === 'relation') return (item.type || 'other').toLowerCase();
    if (col === 'event') return (item.type || 'other').toLowerCase();
    return (item.type || 'other').toLowerCase();
  }

  function buildRows() {
    let rows = '';
    items.forEach((item, i) => {
      const val = getItemValue(item);
      const type = getItemType(item);
      const opts = typeOptions.map(t => `<option value="${t}" ${t === type ? 'selected' : ''}>${t}</option>`).join('');
      const inputType = col === 'event' ? 'date' : 'text';
      rows += `<div class="mv-row" data-idx="${i}">
        <select>${opts}</select>
        <input type="${inputType}" value="${esc(val)}">
        <button class="mv-remove" title="Remove">×</button>
      </div>`;
    });
    return rows;
  }

  function render() {
    popup.innerHTML = `
      <h3>Edit ${label}s</h3>
      <div class="mv-rows">${buildRows()}</div>
      <div class="mv-actions">
        <button class="btn btn-secondary mv-add-btn">+ Add</button>
        <div>
          <button class="btn btn-secondary mv-close-btn">Cancel</button>
          <button class="btn btn-primary mv-apply-btn">Apply</button>
        </div>
      </div>
    `;
    popup.querySelector('.mv-add-btn').onclick = () => {
      if (col === 'address') items.push({ type: 'home', formattedValue: '' });
      else if (col === 'relation') items.push({ type: 'other', person: '' });
      else if (col === 'im') items.push({ protocol: 'other', username: '' });
      else if (col === 'event') items.push({ type: 'other', date: {} });
      else items.push({ type: 'other', value: '' });
      render();
      const inputs = popup.querySelectorAll('.mv-row input');
      if (inputs.length) inputs[inputs.length - 1].focus();
    };
    popup.querySelectorAll('.mv-remove').forEach(btn => {
      btn.onclick = () => {
        const idx = parseInt(btn.closest('.mv-row').dataset.idx);
        items.splice(idx, 1);
        render();
      };
    });
    popup.querySelector('.mv-close-btn').onclick = () => overlay.remove();
    popup.querySelector('.mv-apply-btn').onclick = () => {
      popup.querySelectorAll('.mv-row').forEach((row, i) => {
        const val = row.querySelector('input').value;
        const type = row.querySelector('select').value;
        if (col === 'im') {
          items[i].protocol = type;
          items[i].username = val;
        } else if (col === 'relation') {
          items[i].type = type;
          items[i].person = val;
        } else if (col === 'event') {
          items[i].type = type;
          const parts = val.split('-').map(Number);
          if (parts.length === 3) items[i].date = { year: parts[0], month: parts[1], day: parts[2] };
        } else if (col === 'address') {
          items[i].type = type;
          items[i].formattedValue = val;
          items[i].streetAddress = val;
        } else {
          items[i].type = type;
          items[i].value = val;
        }
      });
      person[arrKey] = items.filter(it => {
        return (getItemValue(it) || '').trim() !== '';
      });
      dirtySet.add(person.resourceName);
      overlay.remove();
      applyFilterAndSort();
    };
  }

  render();
  overlay.appendChild(popup);
  document.body.appendChild(overlay);
  overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); });
}

function closeAllPopups() {
  document.querySelectorAll('.mv-popup-overlay').forEach(el => el.remove());
}

/* ══════════════════════════════════════════
   SAVE (PATCH)
   ══════════════════════════════════════════ */

async function saveContact(resourceName) {
  const person = allContacts.find(p => p.resourceName === resourceName);
  if (!person) return;

  const tr = tableContainer.querySelector(`tr[data-rn="${CSS.escape(resourceName)}"]`);
  if (tr) tr.classList.add('saving');
  setSyncStatus('Saving...', 'saving');

  // Build update body — only mutable fields
  const body = {};
  if (person.names) body.names = person.names;
  if (person.nicknames) body.nicknames = person.nicknames;
  if (person.emailAddresses) body.emailAddresses = person.emailAddresses;
  if (person.phoneNumbers) body.phoneNumbers = person.phoneNumbers;
  if (person.organizations) body.organizations = person.organizations;
  if (person.addresses) body.addresses = person.addresses;
  if (person.birthdays) body.birthdays = person.birthdays;
  if (person.urls) body.urls = person.urls;
  if (person.imClients) body.imClients = person.imClients;
  if (person.relations) body.relations = person.relations;
  if (person.events) body.events = person.events;
  if (person.biographies) body.biographies = person.biographies;
  body.etag = person.etag;

  const updateFields = ['names','nicknames','emailAddresses','phoneNumbers','organizations','addresses','birthdays','urls','imClients','relations','events','biographies'];

  try {
    const res = await apiFetch(
      `https://people.googleapis.com/v1/${resourceName}:updateContact?updatePersonFields=${updateFields.join(',')}&personFields=${PERSON_FIELDS}`,
      {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      }
    );

    if (!res.ok) {
      const err = await res.json();
      if (res.status === 400 && err.error && err.error.status === 'FAILED_PRECONDITION') {
        toast('Contact was modified externally. Reloading...', 'info');
        dirtySet.delete(resourceName);
        await loadAll();
        return;
      }
      throw new Error(err.error?.message || 'Update failed');
    }

    const updated = await res.json();
    const idx = allContacts.findIndex(p => p.resourceName === resourceName);
    if (idx >= 0) allContacts[idx] = updated;
    dirtySet.delete(resourceName);
    setSyncStatus('Saved', 'success');
    toast('Contact saved', 'success');
    setTimeout(() => { if (syncStatus.classList.contains('success')) syncStatus.textContent = ''; }, 3000);
    // Only rebuild table if we're looking at it
    if (!detailPerson) applyFilterAndSort();
  } catch (e) {
    setSyncStatus('Save failed', 'error');
    toast('Save failed: ' + e.message, 'error');
    if (tr) tr.classList.remove('saving');
  }
}

/* ══════════════════════════════════════════
   CREATE
   ══════════════════════════════════════════ */

async function createContact() {
  setSyncStatus('Creating...', 'saving');
  try {
    const body = {
      names: [{ givenName: '', familyName: '' }],
    };
    const res = await apiFetch(
      'https://people.googleapis.com/v1/people:createContact?personFields=' + PERSON_FIELDS,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      }
    );
    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.error?.message || 'Create failed');
    }
    const person = await res.json();
    allContacts.unshift(person);
    searchInput.value = '';
    sortCol = 'name';
    sortAsc = true;
    applyFilterAndSort();
    setSyncStatus('Created', 'success');
    toast('New contact created — edit the name to get started', 'info');
    setTimeout(() => { if (syncStatus.classList.contains('success')) syncStatus.textContent = ''; }, 3000);

    // Auto-open name editor on new contact
    requestAnimationFrame(() => {
      const tr = tableContainer.querySelector(`tr[data-rn="${CSS.escape(person.resourceName)}"]`);
      if (tr) {
        const td = tr.querySelector('td[data-col="name"]');
        if (td) openNameEditor(td, person);
      }
    });
  } catch (e) {
    setSyncStatus('Create failed', 'error');
    toast('Failed to create contact: ' + e.message, 'error');
  }
}

newContactBtn.onclick = createContact;

/* ══════════════════════════════════════════
   DELETE
   ══════════════════════════════════════════ */

function confirmDelete(resourceName) {
  const person = allContacts.find(p => p.resourceName === resourceName);
  const name = person ? getPrimaryName(person).display || 'this contact' : 'this contact';
  deleteModalMsg.textContent = `Are you sure you want to delete "${name}"? This cannot be undone.`;
  pendingDeleteRN = resourceName;
  deleteModal.classList.remove('hidden');
}

deleteCancelBtn.onclick = () => {
  deleteModal.classList.add('hidden');
  pendingDeleteRN = null;
};

deleteConfirmBtn.onclick = async () => {
  deleteModal.classList.add('hidden');
  if (!pendingDeleteRN) return;
  const rn = pendingDeleteRN;
  pendingDeleteRN = null;
  await deleteContact(rn);
};

async function deleteContact(resourceName) {
  setSyncStatus('Deleting...', 'saving');
  try {
    const res = await apiFetch(
      `https://people.googleapis.com/v1/${resourceName}:deleteContact`,
      { method: 'DELETE' }
    );
    if (!res.ok && res.status !== 204) {
      const err = await res.json();
      throw new Error(err.error?.message || 'Delete failed');
    }
    allContacts = allContacts.filter(p => p.resourceName !== resourceName);
    dirtySet.delete(resourceName);
    setSyncStatus('Deleted', 'success');
    toast('Contact deleted', 'success');
    setTimeout(() => { if (syncStatus.classList.contains('success')) syncStatus.textContent = ''; }, 3000);
    applyFilterAndSort();
  } catch (e) {
    setSyncStatus('Delete failed', 'error');
    toast('Delete failed: ' + e.message, 'error');
  }
}

/* ══════════════════════════════════════════
   UI HELPERS
   ══════════════════════════════════════════ */

function setSyncStatus(text, cls) {
  syncStatus.textContent = text;
  syncStatus.className = 'sync-status' + (cls ? ' ' + cls : '');
}

function toast(msg, type = 'info') {
  const el = document.createElement('div');
  el.className = 'toast ' + type;
  el.textContent = msg;
  toastContainer.appendChild(el);
  setTimeout(() => { el.remove(); }, 4000);
}

/* ══════════════════════════════════════════
   DETAIL VIEW
   ══════════════════════════════════════════ */

let detailPerson = null;

function openDetailView(person) {
  detailPerson = person;
  tableContainer.classList.add('hidden-keep-layout');
  detailView.classList.add('active');
  renderDetailView();
}

let detailWasSaved = false;

function closeDetailView() {
  const needsRefresh = detailWasSaved || (detailPerson && dirtySet.has(detailPerson.resourceName));
  detailView.classList.remove('active');
  detailView.innerHTML = '';
  detailPerson = null;
  detailWasSaved = false;
  tableContainer.classList.remove('hidden-keep-layout');
  if (needsRefresh) applyFilterAndSort();
}

function renderDetailView() {
  const p = detailPerson;
  if (!p) return;
  const name = getPrimaryName(p);
  const org = getOrganization(p);
  const dirty = dirtySet.has(p.resourceName);

  let html = '';
  html += `<button class="detail-back" onclick="closeDetailView()">&#8592; Back to list</button>`;
  html += `<div class="detail-header">`;
  html += `<h2>${esc(name.display || 'No Name')}</h2>`;
  html += `<div class="detail-actions">`;
  html += `<button class="btn btn-primary" onclick="detailSave()" ${dirty ? '' : 'disabled style="opacity:0.5"'}>Save Changes</button>`;
  html += `<button class="btn btn-danger" onclick="confirmDelete('${esc(p.resourceName)}')">Delete</button>`;
  html += `</div></div>`;

  // Name section
  html += detailSection('Name', [
    detailEditable('First Name', name.given, 'name-given'),
    detailEditable('Last Name', name.family, 'name-family'),
    detailEditable('Nickname', getNickname(p), 'nickname'),
  ]);

  // Contact info
  html += detailMultiSection('Email', p.emailAddresses || [], 'email',
    it => it.value || '', it => it.type || '');
  html += detailMultiSection('Phone', p.phoneNumbers || [], 'phone',
    it => it.value || '', it => it.type || '');
  html += detailMultiSection('Website', p.urls || [], 'website',
    it => it.value || '', it => it.type || '');
  html += detailMultiSection('IM', p.imClients || [], 'im',
    it => it.username || '', it => it.protocol || '');

  // Work
  html += detailSection('Work', [
    detailEditable('Company', org.company, 'company'),
    detailEditable('Title', org.title, 'jobtitle'),
  ]);

  // Address
  html += detailMultiSection('Address', p.addresses || [], 'address',
    it => it.formattedValue || it.streetAddress || '', it => it.type || '');

  // Dates
  html += detailSection('Dates', [
    detailEditable('Birthday', getBirthday(p), 'birthday', 'date'),
    ...((p.events || []).map((ev, i) => {
      const d = ev.date || {};
      const ds = d.year ? `${d.year}-${String(d.month).padStart(2,'0')}-${String(d.day).padStart(2,'0')}` : '';
      return detailEditable((ev.type || 'Event') + ' ' + (i+1), ds, 'event-' + i, 'date');
    })),
  ]);

  // Relations
  html += detailMultiSection('Relations', p.relations || [], 'relation',
    it => it.person || '', it => it.type || '');

  // Groups (read-only)
  const groups = getGroups(p);
  html += detailSection('Groups', [
    `<div class="detail-field"><span class="detail-field-label">Labels</span><div class="detail-field-value">${esc(groups) || '<span style="color:#ccc">—</span>'}</div></div>`
  ]);

  // Notes
  html += detailSection('Notes', [
    detailEditable('Notes', getNotes(p), 'notes', 'textarea'),
  ]);

  detailView.innerHTML = html;
  bindDetailEvents();
}

function detailSection(title, fieldHtmls) {
  return `<div class="detail-section"><h3>${esc(title)}</h3>${fieldHtmls.join('')}</div>`;
}

function detailEditable(label, value, fieldKey, inputType) {
  return `<div class="detail-field">
    <span class="detail-field-label">${esc(label)}</span>
    <div class="detail-field-value" data-field="${fieldKey}" data-input-type="${inputType || 'text'}">${esc(value)}</div>
  </div>`;
}

function detailMultiSection(title, items, colType, valFn, typeFn) {
  let inner = '';
  if (items.length === 0) {
    inner = `<div class="detail-field"><span class="detail-field-label">${esc(title)}</span><div class="detail-field-value"><span style="color:#ccc">—</span></div></div>`;
  } else {
    items.forEach((it, i) => {
      const type = typeFn(it);
      inner += `<div class="detail-field">
        <span class="detail-field-label">${esc(type || title)}</span>
        <div class="detail-field-value" data-multi="${colType}" data-idx="${i}">${esc(valFn(it))}</div>
      </div>`;
    });
  }
  inner += `<div style="padding:4px 16px 8px;"><button class="detail-add-btn" data-add="${colType}">+ Add ${title.toLowerCase()}</button></div>`;
  return `<div class="detail-section"><h3>${esc(title)}</h3>${inner}</div>`;
}

function bindDetailEvents() {
  const p = detailPerson;
  if (!p) return;

  // Editable single fields — click to edit inline
  detailView.querySelectorAll('.detail-field-value[data-field]').forEach(el => {
    el.addEventListener('click', () => {
      if (el.querySelector('input, textarea')) return;
      const fieldKey = el.dataset.field;
      const inputType = el.dataset.inputType || 'text';
      const currentVal = getDetailFieldValue(p, fieldKey);

      if (inputType === 'textarea') {
        const ta = document.createElement('textarea');
        ta.value = currentVal;
        el.textContent = '';
        el.appendChild(ta);
        ta.focus();
        ta.addEventListener('blur', () => {
          setDetailFieldValue(p, fieldKey, ta.value);
          renderDetailView();
        });
      } else {
        const inp = document.createElement('input');
        inp.type = inputType;
        inp.value = currentVal;
        el.textContent = '';
        el.appendChild(inp);
        inp.focus();
        inp.addEventListener('blur', () => {
          setDetailFieldValue(p, fieldKey, inp.value);
          renderDetailView();
        });
        inp.addEventListener('keydown', e => { if (e.key === 'Enter') inp.blur(); });
      }
    });
  });

  // Multi-value fields — click to edit inline
  detailView.querySelectorAll('.detail-field-value[data-multi]').forEach(el => {
    el.addEventListener('click', () => {
      if (el.querySelector('input')) return;
      const colType = el.dataset.multi;
      const idx = parseInt(el.dataset.idx);
      const currentVal = el.textContent.trim();
      const inp = document.createElement('input');
      inp.type = 'text';
      inp.value = currentVal;
      el.textContent = '';
      el.appendChild(inp);
      inp.focus();
      inp.addEventListener('blur', () => {
        setDetailMultiValue(p, colType, idx, inp.value);
        renderDetailView();
      });
      inp.addEventListener('keydown', e => { if (e.key === 'Enter') inp.blur(); });
    });
  });

  // Add buttons
  detailView.querySelectorAll('[data-add]').forEach(btn => {
    btn.addEventListener('click', () => {
      const colType = btn.dataset.add;
      addDetailMultiValue(p, colType);
      dirtySet.add(p.resourceName);
      renderDetailView();
    });
  });
}

function getDetailFieldValue(p, key) {
  switch (key) {
    case 'name-given': return getPrimaryName(p).given;
    case 'name-family': return getPrimaryName(p).family;
    case 'nickname': return getNickname(p);
    case 'company': return getOrganization(p).company;
    case 'jobtitle': return getOrganization(p).title;
    case 'birthday': return getBirthday(p);
    case 'notes': return getNotes(p);
    default:
      if (key.startsWith('event-')) {
        const idx = parseInt(key.split('-')[1]);
        const ev = (p.events || [])[idx];
        if (!ev || !ev.date) return '';
        const d = ev.date;
        return d.year ? `${d.year}-${String(d.month).padStart(2,'0')}-${String(d.day).padStart(2,'0')}` : '';
      }
      return '';
  }
}

function setDetailFieldValue(p, key, val) {
  dirtySet.add(p.resourceName);
  switch (key) {
    case 'name-given':
      if (!p.names) p.names = [{}];
      if (!p.names[0]) p.names[0] = {};
      p.names[0].givenName = val;
      p.names[0].unstructuredName = (val + ' ' + (p.names[0].familyName || '')).trim();
      p.names[0].displayName = p.names[0].unstructuredName;
      break;
    case 'name-family':
      if (!p.names) p.names = [{}];
      if (!p.names[0]) p.names[0] = {};
      p.names[0].familyName = val;
      p.names[0].unstructuredName = ((p.names[0].givenName || '') + ' ' + val).trim();
      p.names[0].displayName = p.names[0].unstructuredName;
      break;
    case 'nickname':
      setFieldValue(p, 'nickname', val);
      break;
    case 'company':
      setFieldValue(p, 'company', val);
      break;
    case 'jobtitle':
      setFieldValue(p, 'title', val);
      break;
    case 'birthday':
      setFieldValue(p, 'birthday', val);
      break;
    case 'notes':
      setFieldValue(p, 'notes', val);
      break;
    default:
      if (key.startsWith('event-')) {
        const idx = parseInt(key.split('-')[1]);
        if (!p.events) p.events = [];
        while (p.events.length <= idx) p.events.push({ type: 'other', date: {} });
        const parts = val.split('-').map(Number);
        if (parts.length === 3) p.events[idx].date = { year: parts[0], month: parts[1], day: parts[2] };
      }
  }
}

function setDetailMultiValue(p, colType, idx, val) {
  dirtySet.add(p.resourceName);
  const map = { email: 'emailAddresses', phone: 'phoneNumbers', address: 'addresses', website: 'urls', relation: 'relations', im: 'imClients' };
  const arrKey = map[colType];
  if (!arrKey || !p[arrKey] || !p[arrKey][idx]) return;

  if (colType === 'address') { p[arrKey][idx].formattedValue = val; p[arrKey][idx].streetAddress = val; }
  else if (colType === 'relation') p[arrKey][idx].person = val;
  else if (colType === 'im') p[arrKey][idx].username = val;
  else p[arrKey][idx].value = val;
}

function addDetailMultiValue(p, colType) {
  const map = { email: 'emailAddresses', phone: 'phoneNumbers', address: 'addresses', website: 'urls', relation: 'relations', im: 'imClients' };
  const arrKey = map[colType];
  if (!arrKey) return;
  if (!p[arrKey]) p[arrKey] = [];
  if (colType === 'address') p[arrKey].push({ type: 'home', formattedValue: '' });
  else if (colType === 'relation') p[arrKey].push({ type: 'other', person: '' });
  else if (colType === 'im') p[arrKey].push({ protocol: 'other', username: '' });
  else p[arrKey].push({ type: 'other', value: '' });
}

async function detailSave() {
  if (!detailPerson) return;
  // Immediate visual feedback
  const btn = detailView.querySelector('.btn-primary');
  if (btn) { btn.textContent = 'Saving...'; btn.disabled = true; }
  await saveContact(detailPerson.resourceName);
  detailWasSaved = true;
  detailPerson = allContacts.find(p => p.resourceName === detailPerson.resourceName);
  renderDetailView();
}

/* ══════════════════════════════════════════
   TAB NAVIGATION
   ══════════════════════════════════════════ */

function switchTab(tab) {
  currentTab = tab;
  tabNav.querySelectorAll('.tab-btn').forEach(b => {
    b.classList.toggle('active', b.dataset.tab === tab);
  });
  // Contacts tab
  const showContacts = tab === 'contacts';
  tableContainer.style.display = showContacts && accessToken ? '' : 'none';
  searchInput.style.display = showContacts ? '' : 'none';
  newContactBtn.style.display = showContacts && accessToken ? '' : 'none';
  contactCount.style.display = showContacts ? '' : 'none';
  if (showContacts && detailPerson) {
    detailView.classList.add('active');
    tableContainer.classList.add('hidden-keep-layout');
  } else if (!showContacts) {
    detailView.classList.remove('active');
  }
  // Enrich tab
  enrichView.classList.toggle('active', tab === 'enrich');
  if (tab === 'enrich') {
    if (enrichSelectedContacts.length === 0) {
      enrichRefreshSelection();
    }
    renderEnrichProgress();
  }
}

tabNav.addEventListener('click', e => {
  const btn = e.target.closest('.tab-btn');
  if (btn) switchTab(btn.dataset.tab);
});

/* ══════════════════════════════════════════
   ENRICHMENT — CONTACT SELECTION
   ══════════════════════════════════════════ */

function getEnrichmentHistory() {
  try { return JSON.parse(localStorage.getItem('enrichment_history') || '{}'); }
  catch { return {}; }
}

function saveEnrichmentHistory(hist) {
  localStorage.setItem('enrichment_history', JSON.stringify(hist));
}

function pickEnrichmentBatch(size) {
  const hist = getEnrichmentHistory();
  const contacts = allContacts.filter(p => {
    const name = getPrimaryName(p).display;
    return name && name.trim() !== '';
  });
  contacts.sort((a, b) => {
    const ha = hist[a.resourceName];
    const hb = hist[b.resourceName];
    if (!ha && hb) return -1;
    if (ha && !hb) return 1;
    if (!ha && !hb) return 0;
    return (ha.lastEnrichedAt || 0) - (hb.lastEnrichedAt || 0);
  });
  return contacts.slice(0, size);
}

function enrichRefreshSelection() {
  const size = parseInt($('batch-size').value) || 15;
  enrichSelectedContacts = pickEnrichmentBatch(size);
  renderEnrichChecklist();
}

function renderEnrichChecklist() {
  const hist = getEnrichmentHistory();
  const container = $('enrich-checklist');
  if (enrichSelectedContacts.length === 0) {
    container.innerHTML = '<div style="padding:20px;text-align:center;color:#999;">No contacts available to enrich.</div>';
    return;
  }
  container.innerHTML = enrichSelectedContacts.map((p, i) => {
    const name = getPrimaryName(p).display;
    const email = getPrimaryEmail(p);
    const org = getOrganization(p);
    const h = hist[p.resourceName];
    const meta = h ? `Last enriched: ${new Date(h.lastEnrichedAt).toLocaleDateString()}` : 'Never enriched';
    return `<label class="checklist-item">
      <input type="checkbox" checked data-idx="${i}">
      <span class="checklist-name">${esc(name)}</span>
      <span class="checklist-meta">${esc(email)}${org.company ? ' · ' + esc(org.company) : ''} · ${meta}</span>
    </label>`;
  }).join('');
  // All start checked, so show Deselect All
  $('enrich-select-all-btn').textContent = 'Deselect All';
}

$('enrich-select-all-btn').addEventListener('click', () => {
  const checkboxes = $('enrich-checklist').querySelectorAll('input[type="checkbox"]');
  const allChecked = Array.from(checkboxes).every(cb => cb.checked);
  checkboxes.forEach(cb => cb.checked = !allChecked);
  $('enrich-select-all-btn').textContent = allChecked ? 'Select All' : 'Deselect All';
});

$('enrich-refresh-btn').addEventListener('click', enrichRefreshSelection);

/* ══════════════════════════════════════════
   ENRICHMENT — PROMPT GENERATION
   ══════════════════════════════════════════ */

function getCheckedContacts() {
  const checkboxes = $('enrich-checklist').querySelectorAll('input[type="checkbox"]');
  const selected = [];
  checkboxes.forEach(cb => {
    if (cb.checked) {
      const idx = parseInt(cb.dataset.idx);
      if (enrichSelectedContacts[idx]) selected.push(enrichSelectedContacts[idx]);
    }
  });
  return selected;
}

function generateEnrichmentPrompt() {
  const contacts = getCheckedContacts();
  if (contacts.length === 0) { toast('Select at least one contact', 'error'); return; }

  let prompt = `I need you to research the following ${contacts.length} people and return structured JSON data for each.\n\n`;
  prompt += `For each person, use any publicly available information (LinkedIn, company websites, social media, etc.) to fill in missing details.\n\n`;
  prompt += `Here are the contacts:\n\n`;

  contacts.forEach((p, i) => {
    const name = getPrimaryName(p).display;
    const email = getPrimaryEmail(p);
    const org = getOrganization(p);
    const addr = getPrimaryAddress(p);
    prompt += `${i + 1}. ${name}\n`;
    prompt += `   resourceName: "${p.resourceName}"\n`;
    if (email) prompt += `   email: ${email}\n`;
    if (org.company) prompt += `   company: ${org.company}\n`;
    if (org.title) prompt += `   title: ${org.title}\n`;
    if (addr) prompt += `   location: ${addr}\n`;
    prompt += `\n`;
  });

  prompt += `Return a JSON array with one object per person. Each field is an object with value, confidence, and source.\n\n`;
  prompt += `[\n  {\n`;
  prompt += `    "resourceName": "people/cXXX",\n`;
  prompt += `    "company": { "value": "Acme Corp", "confidence": 90, "source": "https://linkedin.com/in/johndoe" },\n`;
  prompt += `    "title": { "value": "VP Engineering", "confidence": 85, "source": "https://linkedin.com/in/johndoe" },\n`;
  prompt += `    "email": { "value": "john@acme.com", "confidence": 70, "source": "https://acme.com/team" },\n`;
  prompt += `    "phone": { "value": "", "confidence": 0, "source": "" },\n`;
  prompt += `    "location": { "value": "San Francisco, CA", "confidence": 80, "source": "https://linkedin.com/in/johndoe" },\n`;
  prompt += `    "linkedin": { "value": "https://linkedin.com/in/johndoe", "confidence": 95, "source": "https://linkedin.com/in/johndoe" },\n`;
  prompt += `    "twitter": { "value": "", "confidence": 0, "source": "" },\n`;
  prompt += `    "github": { "value": "", "confidence": 0, "source": "" },\n`;
  prompt += `    "facebook": { "value": "", "confidence": 0, "source": "" },\n`;
  prompt += `    "instagram": { "value": "", "confidence": 0, "source": "" },\n`;
  prompt += `    "website": { "value": "", "confidence": 0, "source": "" }\n`;
  prompt += `  }\n]\n\n`;
  prompt += `Rules:\n`;
  prompt += `- "resourceName" must match exactly what I gave you above\n`;
  prompt += `- Each field is { "value": "...", "confidence": 0-100, "source": "URL" }\n`;
  prompt += `- "confidence" is per-field: how sure you are about THIS specific piece of data\n`;
  prompt += `- "source" is the URL where you found the info, or "" if general knowledge\n`;
  prompt += `- Empty string value means no data found for that field\n`;
  prompt += `- Don't repeat information that was already provided above — only add NEW info\n`;
  prompt += `- Return ONLY the JSON array, no other text\n`;

  // Show prompt step
  showEnrichStep('enrich-prompt');
  $('enrich-prompt-text').value = prompt;
  $('enrich-response-text').value = '';
  $('enrich-parse-error').textContent = '';
}

$('enrich-generate-btn').addEventListener('click', generateEnrichmentPrompt);

$('enrich-copy-btn').addEventListener('click', () => {
  const text = $('enrich-prompt-text').value;
  navigator.clipboard.writeText(text).then(() => {
    toast('Prompt copied to clipboard', 'success');
    $('enrich-copy-btn').textContent = 'Copied!';
    setTimeout(() => { $('enrich-copy-btn').textContent = 'Copy to Clipboard'; }, 2000);
  }).catch(() => {
    $('enrich-prompt-text').select();
    document.execCommand('copy');
    toast('Prompt copied', 'success');
  });
});

$('enrich-back-select-btn').addEventListener('click', () => showEnrichStep('enrich-select'));

/* ══════════════════════════════════════════
   ENRICHMENT — PARSE RESPONSE
   ══════════════════════════════════════════ */

const ENRICH_FIELDS = ['company','title','email','phone','location','linkedin','twitter','github','facebook','instagram','website'];

function parseEnrichmentResponse() {
  const raw = $('enrich-response-text').value.trim();
  if (!raw) { $('enrich-parse-error').textContent = 'Paste the AI response first.'; return; }

  // Strip markdown code fences
  let cleaned = raw;
  cleaned = cleaned.replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/, '');

  let parsed;
  try {
    parsed = JSON.parse(cleaned);
  } catch (e) {
    $('enrich-parse-error').textContent = `JSON parse error: ${e.message}`;
    return;
  }

  if (!Array.isArray(parsed)) {
    $('enrich-parse-error').textContent = 'Expected a JSON array of results.';
    return;
  }

  // Validate structure
  const checkedContacts = getCheckedContacts();
  const validRNs = new Set(checkedContacts.map(p => p.resourceName));
  const errors = [];
  parsed.forEach((item, i) => {
    if (!item.resourceName) errors.push(`Item ${i + 1}: missing resourceName`);
    else if (!validRNs.has(item.resourceName)) errors.push(`Item ${i + 1}: resourceName "${item.resourceName}" not in selected batch`);
  });

  if (errors.length > 0) {
    $('enrich-parse-error').textContent = 'Validation errors:\n' + errors.join('\n');
    return;
  }

  $('enrich-parse-error').textContent = '';

  // Normalize: support both new per-field {value, confidence, source} and old plain-string format
  parsed.forEach(result => {
    ENRICH_FIELDS.forEach(f => {
      const raw = result[f];
      if (raw && typeof raw === 'object' && 'value' in raw) {
        // New format — keep as-is
      } else if (typeof raw === 'string') {
        // Old plain-string format — wrap with defaults
        result[f] = { value: raw, confidence: 50, source: '' };
      } else if (raw == null) {
        result[f] = { value: '', confidence: 0, source: '' };
      }
    });
  });

  enrichResults = parsed;

  // Initialize card states with per-field confidence and source
  enrichCardStates = {};
  parsed.forEach(result => {
    const fields = {};
    ENRICH_FIELDS.forEach(f => {
      const fld = result[f] || { value: '', confidence: 0, source: '' };
      if (fld.value && fld.value.trim()) {
        fields[f] = {
          checked: (fld.confidence || 0) >= 85,
          confidence: fld.confidence || 0,
          source: fld.source || '',
          value: fld.value.trim(),
          action: null, // null = auto-detect via getFieldAction
        };
      }
    });
    enrichCardStates[result.resourceName] = { status: 'pending', fields };
  });

  showEnrichStep('enrich-review');
  renderEnrichmentReview();
}

$('enrich-parse-btn').addEventListener('click', parseEnrichmentResponse);

/* ══════════════════════════════════════════
   ENRICHMENT — REVIEW CARDS
   ══════════════════════════════════════════ */

// Fields that append (add new entry) vs replace (overwrite existing)
const APPEND_FIELDS = new Set(['email', 'phone', 'linkedin', 'twitter', 'github', 'facebook', 'instagram', 'website']);

function getFieldAction(field, currentVal, newVal) {
  if (currentVal.toLowerCase() === newVal.toLowerCase()) return 'same';
  return APPEND_FIELDS.has(field) ? 'add' : 'replace';
}

function confidencePill(conf) {
  const cls = conf >= 85 ? 'field-conf-high' : conf >= 60 ? 'field-conf-mid' : 'field-conf-low';
  return `<span class="field-confidence ${cls}">${conf}%</span>`;
}

function sourceLink(src) {
  if (!src) return '<span class="field-source"><span class="no-source">--</span></span>';
  const short = src.replace(/^https?:\/\/(www\.)?/, '').split('/')[0];
  return `<span class="field-source"><a href="${esc(src)}" target="_blank" rel="noopener" title="${esc(src)}">${esc(short)}</a></span>`;
}

function renderEnrichmentReview() {
  // Bulk actions bar
  const pending = enrichResults.filter(r => enrichCardStates[r.resourceName]?.status === 'pending').length;
  const approved = enrichResults.filter(r => enrichCardStates[r.resourceName]?.status === 'approved').length;
  const skipped = enrichResults.filter(r => enrichCardStates[r.resourceName]?.status === 'skipped').length;
  const bulkEl = $('enrich-review-bulk');
  bulkEl.innerHTML = `
    <span class="review-summary">${approved} approved, ${skipped} skipped, ${pending} pending</span>
    ${pending > 0 ? `<button class="btn btn-primary btn-sm" id="enrich-approve-all">Approve All Remaining</button>
    <button class="btn btn-secondary btn-sm" id="enrich-skip-all">Skip All Remaining</button>` : ''}
  `;

  if ($('enrich-approve-all')) {
    $('enrich-approve-all').addEventListener('click', () => {
      enrichResults.forEach(r => {
        if (enrichCardStates[r.resourceName]?.status === 'pending') {
          enrichCardStates[r.resourceName].status = 'approved';
        }
      });
      renderEnrichmentReview();
    });
  }
  if ($('enrich-skip-all')) {
    $('enrich-skip-all').addEventListener('click', () => {
      enrichResults.forEach(r => {
        if (enrichCardStates[r.resourceName]?.status === 'pending') {
          enrichCardStates[r.resourceName].status = 'skipped';
        }
      });
      renderEnrichmentReview();
    });
  }

  // Render cards
  const cardsEl = $('enrich-review-cards');
  cardsEl.innerHTML = enrichResults.map(result => {
    const rn = result.resourceName;
    const person = allContacts.find(p => p.resourceName === rn);
    if (!person) return '';
    const state = enrichCardStates[rn];
    const name = getPrimaryName(person).display;
    const cardClass = state.status !== 'pending' ? state.status : '';

    let diffRows = '';
    ENRICH_FIELDS.forEach(field => {
      const fld = state.fields[field];
      if (!fld) return; // no data for this field
      const newVal = fld.value;
      if (!newVal) return;
      const currentVal = getEnrichCurrentValue(person, field);
      const autoAction = getFieldAction(field, currentVal, newVal);
      const action = (autoAction === 'same') ? 'same' : (fld.action || autoAction);
      const isPending = state.status === 'pending';
      const clickable = isPending && action !== 'same' ? ' clickable' : '';
      const actionBadge = action === 'add'
        ? `<span class="action-badge action-add${clickable}" data-rn="${esc(rn)}" data-field="${field}">ADD</span>`
        : action === 'replace'
        ? `<span class="action-badge action-replace${clickable}" data-rn="${esc(rn)}" data-field="${field}">REPLACE</span>`
        : '<span class="action-badge action-same">SAME</span>';
      const checked = fld.checked && isPending ? 'checked' : '';
      const disabled = !isPending ? 'disabled' : '';
      diffRows += `<tr>
        <td><input type="checkbox" ${checked} ${disabled} data-rn="${esc(rn)}" data-field="${field}"></td>
        <td>${actionBadge}</td>
        <td>${esc(field)}</td>
        <td class="diff-current">${esc(currentVal) || '<span style="color:#ccc">empty</span>'}</td>
        <td class="diff-arrow">&rarr;</td>
        <td class="${action === 'same' ? 'diff-same' : 'diff-new'}">${esc(newVal)}</td>
        <td>${confidencePill(fld.confidence)}</td>
        <td>${sourceLink(fld.source)}</td>
      </tr>`;
    });

    if (!diffRows) {
      diffRows = '<tr><td colspan="8" style="text-align:center;color:#999;padding:12px;">No new data found</td></tr>';
    }

    const actions = state.status === 'pending'
      ? `<button class="btn btn-primary btn-sm" data-approve="${esc(rn)}">Approve Selected</button>
         <button class="btn btn-secondary btn-sm" data-skip="${esc(rn)}">Skip</button>`
      : `<span style="font-size:12px;color:#888;text-transform:uppercase;">${state.status}</span>`;

    return `<div class="review-card ${cardClass}" data-card-rn="${esc(rn)}">
      <div class="review-card-header">
        <h3>${esc(name)}</h3>
      </div>
      <table class="review-diff-table">
        <thead><tr><th></th><th>Action</th><th>Field</th><th>Current</th><th></th><th>Found</th><th>Conf</th><th>Source</th></tr></thead>
        <tbody>${diffRows}</tbody>
      </table>
      <div class="review-card-actions">${actions}</div>
    </div>`;
  }).join('');

  // Bind card events
  cardsEl.querySelectorAll('[data-approve]').forEach(btn => {
    btn.addEventListener('click', () => {
      const rn = btn.dataset.approve;
      // Read checkbox states before approving
      cardsEl.querySelectorAll(`input[data-rn="${CSS.escape(rn)}"]`).forEach(cb => {
        const field = cb.dataset.field;
        if (enrichCardStates[rn].fields[field]) {
          enrichCardStates[rn].fields[field].checked = cb.checked;
        }
      });
      enrichCardStates[rn].status = 'approved';
      renderEnrichmentReview();
    });
  });
  cardsEl.querySelectorAll('[data-skip]').forEach(btn => {
    btn.addEventListener('click', () => {
      enrichCardStates[btn.dataset.skip].status = 'skipped';
      renderEnrichmentReview();
    });
  });

  // Checkbox change updates state
  cardsEl.querySelectorAll('input[type="checkbox"][data-rn]').forEach(cb => {
    cb.addEventListener('change', () => {
      const rn = cb.dataset.rn;
      const field = cb.dataset.field;
      if (enrichCardStates[rn]?.fields[field]) {
        enrichCardStates[rn].fields[field].checked = cb.checked;
      }
    });
  });

  // Action badge toggle (ADD ↔ REPLACE)
  cardsEl.querySelectorAll('.action-badge.clickable').forEach(badge => {
    badge.addEventListener('click', () => {
      const rn = badge.dataset.rn;
      const field = badge.dataset.field;
      const fld = enrichCardStates[rn]?.fields[field];
      if (!fld) return;
      const person = allContacts.find(p => p.resourceName === rn);
      if (!person) return;
      const currentVal = getEnrichCurrentValue(person, field);
      const autoAction = getFieldAction(field, currentVal, fld.value);
      const current = fld.action || autoAction;
      fld.action = current === 'add' ? 'replace' : 'add';
      renderEnrichmentReview();
    });
  });
}

function getEnrichCurrentValue(person, field) {
  switch (field) {
    case 'company': return getOrganization(person).company;
    case 'title': return getOrganization(person).title;
    case 'email': return getPrimaryEmail(person);
    case 'phone': return getPrimaryPhone(person);
    case 'location': return getPrimaryAddress(person);
    case 'linkedin': return (person.urls || []).find(u => /linkedin/i.test(u.value || ''))?.value || '';
    case 'twitter': return (person.urls || []).find(u => /twitter|x\.com/i.test(u.value || ''))?.value || '';
    case 'github': return (person.urls || []).find(u => /github/i.test(u.value || ''))?.value || '';
    case 'facebook': return (person.urls || []).find(u => /facebook/i.test(u.value || ''))?.value || '';
    case 'instagram': return (person.urls || []).find(u => /instagram/i.test(u.value || ''))?.value || '';
    case 'website': return getPrimaryWebsite(person);
    default: return '';
  }
}

$('enrich-back-prompt-btn').addEventListener('click', () => showEnrichStep('enrich-prompt'));

/* ══════════════════════════════════════════
   ENRICHMENT — SYNC APPROVED
   ══════════════════════════════════════════ */

$('enrich-sync-btn').addEventListener('click', syncEnrichmentApproved);

function getEnrichmentLog() {
  try { return JSON.parse(localStorage.getItem('enrichment_log') || '[]'); }
  catch { return []; }
}

function saveEnrichmentLog(log) {
  if (log.length > 500) log.length = 500;
  localStorage.setItem('enrichment_log', JSON.stringify(log));
}

async function syncEnrichmentApproved() {
  const approvedResults = enrichResults.filter(r => enrichCardStates[r.resourceName]?.status === 'approved');
  if (approvedResults.length === 0) {
    toast('No approved contacts to sync', 'info');
    return;
  }

  const syncBtn = $('enrich-sync-btn');
  syncBtn.disabled = true;
  syncBtn.textContent = 'Syncing...';
  setSyncStatus('Syncing enrichment...', 'saving');

  let updated = 0;
  let failed = 0;
  const hist = getEnrichmentHistory();
  const log = getEnrichmentLog();

  for (const result of approvedResults) {
    const rn = result.resourceName;
    const person = allContacts.find(p => p.resourceName === rn);
    if (!person) { failed++; continue; }

    const state = enrichCardStates[rn];
    const checkedFields = Object.entries(state.fields)
      .filter(([, fld]) => fld.checked)
      .map(([k]) => k);
    if (checkedFields.length === 0) continue;

    // Build per-field change log before applying
    const contactName = getPrimaryName(person).display;
    const changes = [];
    checkedFields.forEach(field => {
      const fld = state.fields[field];
      if (!fld || !fld.value) return;
      const currentVal = getEnrichCurrentValue(person, field);
      const autoAction = getFieldAction(field, currentVal, fld.value);
      const action = (autoAction === 'same') ? 'same' : (fld.action || autoAction);
      if (action === 'same') return;
      changes.push({
        field,
        action,
        oldValue: currentVal || '',
        newValue: fld.value,
        source: fld.source || '',
        confidence: fld.confidence || 0,
      });
    });

    // Apply enrichment fields to person object
    checkedFields.forEach(field => {
      const fld = state.fields[field];
      if (!fld || !fld.value) return;
      const currentVal = getEnrichCurrentValue(person, field);
      const autoAction = getFieldAction(field, currentVal, fld.value);
      const action = (autoAction === 'same') ? 'same' : (fld.action || autoAction);
      if (action === 'same') return;
      applyEnrichField(person, field, fld.value, action);
    });

    // Prepend notes audit trail
    if (changes.length > 0) {
      const now = new Date();
      const dateStr = now.getFullYear().toString()
        + String(now.getMonth() + 1).padStart(2, '0')
        + String(now.getDate()).padStart(2, '0');
      const enrichedParts = changes.map(c =>
        c.action === 'replace' && c.oldValue ? `${c.field}(was: "${c.oldValue}")` : c.field
      ).join(', ');
      const deleted = changes.filter(c => c.action === 'replace' && c.oldValue);
      let auditLine = `${dateStr} Enriched: ${enrichedParts}`;
      if (deleted.length > 0) {
        auditLine += ` DELETED: ${deleted.map(c => c.field).join(', ')}`;
      }
      const existingNotes = getNotes(person);
      const newNotes = existingNotes ? auditLine + '\n' + existingNotes : auditLine;
      if (!person.biographies) person.biographies = [];
      if (person.biographies.length === 0) person.biographies.push({ contentType: 'TEXT_PLAIN' });
      person.biographies[0].value = newNotes;
      dirtySet.add(person.resourceName);
    }

    // Save to Google
    let status = 'synced';
    try {
      await saveContact(rn);
      hist[rn] = { lastEnrichedAt: Date.now(), status: 'enriched' };
      updated++;
    } catch (e) {
      failed++;
      status = 'failed';
    }

    // Log per-field changes
    if (changes.length > 0) {
      log.unshift({
        date: Date.now(),
        contactName,
        resourceName: rn,
        changes,
        status,
      });
    }
  }

  saveEnrichmentLog(log);

  // Record batch summary
  const batches = getBatchHistory();
  batches.unshift({
    date: Date.now(),
    total: enrichResults.length,
    approved: approvedResults.length,
    skipped: enrichResults.length - approvedResults.length,
    updated,
    failed,
  });
  if (batches.length > 50) batches.length = 50;
  localStorage.setItem('enrichment_batches', JSON.stringify(batches));

  saveEnrichmentHistory(hist);
  syncBtn.disabled = false;
  syncBtn.textContent = 'Sync Approved';
  toast(`Updated ${updated} contacts${failed ? `, ${failed} failed` : ''}`, updated > 0 ? 'success' : 'error');
  setSyncStatus('', '');

  // Reset to select step for next batch
  enrichResults = [];
  enrichCardStates = {};
  enrichSelectedContacts = [];
  showEnrichStep('enrich-select');
  enrichRefreshSelection();
  renderEnrichProgress();
}

function applyEnrichField(person, field, val, action) {
  // action: 'add' = append, 'replace' = overwrite first entry
  switch (field) {
    case 'company':
      if (!person.organizations) person.organizations = [];
      if (person.organizations.length === 0) person.organizations.push({});
      person.organizations[0].name = val;
      break;
    case 'title':
      if (!person.organizations) person.organizations = [];
      if (person.organizations.length === 0) person.organizations.push({});
      person.organizations[0].title = val;
      break;
    case 'email': {
      if (!person.emailAddresses) person.emailAddresses = [];
      if (action === 'replace' && person.emailAddresses.length > 0) {
        person.emailAddresses[0].value = val;
      } else {
        const exists = person.emailAddresses.some(e => e.value?.toLowerCase() === val.toLowerCase());
        if (!exists) person.emailAddresses.push({ value: val, type: 'other' });
      }
      break;
    }
    case 'phone': {
      if (!person.phoneNumbers) person.phoneNumbers = [];
      if (action === 'replace' && person.phoneNumbers.length > 0) {
        person.phoneNumbers[0].value = val;
      } else {
        const normalized = val.replace(/[\s\-()]/g, '');
        const exists = person.phoneNumbers.some(p => p.value?.replace(/[\s\-()]/g, '') === normalized);
        if (!exists) person.phoneNumbers.push({ value: val, type: 'other' });
      }
      break;
    }
    case 'location':
      if (!person.addresses) person.addresses = [];
      if (person.addresses.length === 0) person.addresses.push({});
      if (action === 'add' && person.addresses[0].formattedValue) {
        person.addresses.push({ formattedValue: val });
      } else {
        person.addresses[0].formattedValue = val;
      }
      break;
    case 'linkedin':
    case 'twitter':
    case 'github':
    case 'facebook':
    case 'instagram': {
      if (!person.urls) person.urls = [];
      if (action === 'replace') {
        const pattern = field === 'twitter' ? /twitter|x\.com/i : new RegExp(field, 'i');
        const idx = person.urls.findIndex(u => pattern.test(u.value || ''));
        if (idx >= 0) { person.urls[idx].value = val; }
        else person.urls.push({ value: val, type: field });
      } else {
        const exists = person.urls.some(u => u.value?.toLowerCase() === val.toLowerCase());
        if (!exists) person.urls.push({ value: val, type: field });
      }
      break;
    }
    case 'website': {
      if (!person.urls) person.urls = [];
      if (action === 'replace' && person.urls.length > 0) {
        person.urls[0].value = val;
      } else {
        const exists = person.urls.some(u => u.value?.toLowerCase() === val.toLowerCase());
        if (!exists) person.urls.push({ value: val, type: 'homepage' });
      }
      break;
    }
  }
  dirtySet.add(person.resourceName);
}

/* ══════════════════════════════════════════
   ENRICHMENT — PROGRESS & HISTORY
   ══════════════════════════════════════════ */

function getBatchHistory() {
  try { return JSON.parse(localStorage.getItem('enrichment_batches') || '[]'); }
  catch { return []; }
}

function renderEnrichProgress() {
  const hist = getEnrichmentHistory();
  const total = allContacts.length;
  const enriched = Object.keys(hist).length;
  const pct = total > 0 ? Math.round((enriched / total) * 100) : 0;

  const log = getEnrichmentLog();
  let logHtml = '';
  if (log.length > 0) {
    logHtml = '<div style="margin-top:16px;"><strong style="font-size:12px;color:#555;">Change History:</strong>';
    logHtml += log.slice(0, 50).map((entry, idx) => {
      const time = new Date(entry.date).toLocaleString();
      const statusCls = entry.status === 'failed' ? 'failed' : 'synced';
      const entryFailed = entry.status === 'failed' ? ' log-failed' : '';
      let changesHtml = '';
      (entry.changes || []).forEach(ch => {
        const actionBadge = ch.action === 'add'
          ? '<span class="action-badge action-add">ADD</span>'
          : '<span class="action-badge action-replace">REPLACE</span>';
        const src = ch.source
          ? `<span class="field-source"><a href="${esc(ch.source)}" target="_blank" rel="noopener">${esc(ch.source.replace(/^https?:\/\/(www\.)?/, '').split('/')[0])}</a></span>`
          : '';
        changesHtml += `<div class="log-change-row">
          ${actionBadge}
          <span class="log-field">${esc(ch.field)}</span>
          <span class="log-old">${esc(ch.oldValue) || '<em>empty</em>'}</span>
          <span class="log-arrow">&rarr;</span>
          <span class="log-new">${esc(ch.newValue)}</span>
          ${confidencePill(ch.confidence)}
          ${src}
        </div>`;
      });
      return `<div class="log-entry${entryFailed}">
        <div class="log-entry-header" data-log-idx="${idx}">
          <span><span class="log-time">${esc(time)}</span> &nbsp; <span class="log-name">${esc(entry.contactName)}</span></span>
          <span class="log-status ${statusCls}">${entry.status} (${(entry.changes || []).length} fields)</span>
        </div>
        <div class="log-changes" id="log-changes-${idx}">${changesHtml}</div>
      </div>`;
    }).join('');
    logHtml += '<button class="clear-history-btn" id="clear-enrich-log">Clear History</button>';
    logHtml += '</div>';
  }

  $('enrich-progress').innerHTML = `
    <h3>Enrichment Progress</h3>
    <div class="progress-bar-outer"><div class="progress-bar-inner" style="width:${pct}%"></div></div>
    <div class="progress-text">${enriched} of ${total} contacts enriched (${pct}%)</div>
    ${logHtml}
  `;

  // Bind expand/collapse
  $('enrich-progress').querySelectorAll('.log-entry-header').forEach(hdr => {
    hdr.addEventListener('click', () => {
      const idx = hdr.dataset.logIdx;
      const changes = $('log-changes-' + idx);
      if (changes) changes.classList.toggle('expanded');
    });
  });

  // Bind clear history
  const clearBtn = $('clear-enrich-log');
  if (clearBtn) {
    clearBtn.addEventListener('click', () => {
      localStorage.removeItem('enrichment_log');
      localStorage.removeItem('enrichment_batches');
      renderEnrichProgress();
      toast('History cleared', 'info');
    });
  }
}

/* ══════════════════════════════════════════
   ENRICHMENT — UI HELPERS
   ══════════════════════════════════════════ */

function showEnrichStep(stepId) {
  document.querySelectorAll('.enrich-step').forEach(el => el.classList.remove('active'));
  $(stepId).classList.add('active');
}

/* ── Init on GIS load ── */
window.addEventListener('load', () => {
  const wait = setInterval(() => {
    if (window.google && google.accounts && google.accounts.oauth2) {
      clearInterval(wait);
      initAuth();
    }
  }, 100);
});
