/* Bookkeeper — Double-Entry Personal Finance PWA
   Dependencies: idb@7 (IndexedDB wrapper), SheetJS/xlsx (Excel import/export)
   Storage: IndexedDB only (offline-first, no backend)
*/
const DB_NAME = 'bookkeeper-db';
const DB_VERSION = 1;
const TX_STORE = 'transactions';
const SETTINGS_STORE = 'settings';

const state = {
  transactions: [],
  settings: { baseCurrency: 'RUB' },
  filters: { year: 'all', month: 'all', search: '' },
  report: { periodType: 'all', year: 'all', month: 'all', selectedAccount: '' },
  deferredPrompt: null,
  db: null,
};

const els = {};

document.addEventListener('DOMContentLoaded', init);

async function init() {
  cacheElements();
  setupTheme();
  wireEvents();
  await initDb();
  await loadAllData();
  ensureDemoIfEmpty();
  renderAll();
  registerServiceWorker();
  setupInstallPrompt();
}

function cacheElements() {
  const ids = [
    'txCount','accountCount','periodLabel','journalList','filterYear','filterMonth','searchInput',
    'reportPeriodType','reportYear','reportMonth','balanceList','accountLedger','ledgerTitle',
    'transactionDialog','transactionForm','dialogTitle','txId','dateInput','debitAccountInput',
    'debitTypeInput','creditAccountInput','creditTypeInput','amountRubInput','currencyInput',
    'foreignAmountInput','rateInput','descriptionInput','deleteTxBtn','newTransactionBtn',
    'closeDialogBtn','cancelDialogBtn','baseCurrencyInput','settingsForm','exportBtn',
    'importFile','clearDataBtn','installBtn'
  ];
  ids.forEach(id => { els[id] = document.getElementById(id); });
  els.tabButtons = [...document.querySelectorAll('.tab-btn')];
  els.views = [...document.querySelectorAll('.view')];
  els.accountsDebitList = document.getElementById('accountsDebitList');
  els.accountsCreditList = document.getElementById('accountsCreditList');
}

function setupTheme() {
  const root = document.documentElement;
  const toggle = document.querySelector('[data-theme-toggle]');
  let theme = window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  root.setAttribute('data-theme', theme);
  if (toggle) {
    toggle.addEventListener('click', () => {
      theme = theme === 'dark' ? 'light' : 'dark';
      root.setAttribute('data-theme', theme);
    });
  }
}

function wireEvents() {
  els.tabButtons.forEach(btn => btn.addEventListener('click', () => switchView(btn.dataset.view)));
  els.newTransactionBtn.addEventListener('click', () => openTransactionDialog());
  els.closeDialogBtn.addEventListener('click', closeDialog);
  els.cancelDialogBtn.addEventListener('click', closeDialog);
  els.transactionDialog.addEventListener('click', e => { if (e.target === els.transactionDialog) closeDialog(); });
  els.transactionForm.addEventListener('submit', saveTransactionFromForm);
  els.deleteTxBtn.addEventListener('click', deleteCurrentTransaction);
  els.filterYear.addEventListener('change', e => { state.filters.year = e.target.value; renderJournal(); updateSummary(); });
  els.filterMonth.addEventListener('change', e => { state.filters.month = e.target.value; renderJournal(); updateSummary(); });
  els.searchInput.addEventListener('input', e => { state.filters.search = e.target.value.trim().toLowerCase(); renderJournal(); });
  els.reportPeriodType.addEventListener('change', e => { state.report.periodType = e.target.value; renderReport(); updateSummary(); });
  els.reportYear.addEventListener('change', e => { state.report.year = e.target.value; renderReport(); updateSummary(); });
  els.reportMonth.addEventListener('change', e => { state.report.month = e.target.value; renderReport(); updateSummary(); });
  els.baseCurrencyInput.addEventListener('change', async e => {
    state.settings.baseCurrency = (e.target.value || 'RUB').trim().toUpperCase();
    await saveSetting('baseCurrency', state.settings.baseCurrency);
    renderAll();
  });
  els.exportBtn.addEventListener('click', exportToXlsx);
  els.importFile.addEventListener('change', importFromXlsx);
  els.clearDataBtn.addEventListener('click', clearAllData);
  els.currencyInput.addEventListener('input', autoComputeRateOrForeign);
  els.foreignAmountInput.addEventListener('input', autoComputeRateOrForeign);
  els.rateInput.addEventListener('input', autoComputeRateOrForeign);
  els.amountRubInput.addEventListener('input', autoComputeRateOrForeign);
}

async function initDb() {
  state.db = await idb.openDB(DB_NAME, DB_VERSION, {
    upgrade(db) {
      if (!db.objectStoreNames.contains(TX_STORE)) {
        const store = db.createObjectStore(TX_STORE, { keyPath: 'id' });
        store.createIndex('date', 'date');
      }
      if (!db.objectStoreNames.contains(SETTINGS_STORE)) {
        db.createObjectStore(SETTINGS_STORE);
      }
    }
  });
}

async function loadAllData() {
  state.transactions = await state.db.getAll(TX_STORE);
  state.transactions.sort((a, b) =>
    b.date.localeCompare(a.date) || (b.createdAt || '').localeCompare(a.createdAt || '')
  );
  const baseCurrency = await state.db.get(SETTINGS_STORE, 'baseCurrency');
  if (baseCurrency) state.settings.baseCurrency = baseCurrency;
  els.baseCurrencyInput.value = state.settings.baseCurrency;
}

async function ensureDemoIfEmpty() {
  if (state.transactions.length) return;
  const today = todayStr();
  const sample = [
    mkTx({ date: today, debitAccount: 'Наличные RUB', creditAccount: 'Доход: Зарплата', amountRub: 150000, description: 'Зарплата за месяц', debitType: 'active', creditType: 'passive' }),
    mkTx({ date: today, debitAccount: 'Брокерский счет', creditAccount: 'Наличные RUB', amountRub: 20000, description: 'Пополнение брокерского счета', debitType: 'active', creditType: 'active' }),
    mkTx({ date: today, debitAccount: 'Наличный USD', creditAccount: 'Наличные RUB', amountRub: 8750, currency: 'USD', foreignAmount: 100, rate: 87.50, description: 'Обмен валюты', debitType: 'active', creditType: 'active' }),
  ];
  for (const tx of sample) await upsertTransaction(tx, false);
}

function mkTx(data) {
  const now = new Date().toISOString();
  return {
    id: crypto.randomUUID(),
    date: data.date,
    debitAccount: (data.debitAccount || '').trim(),
    creditAccount: (data.creditAccount || '').trim(),
    debitType: data.debitType || 'active',
    creditType: data.creditType || 'passive',
    amountRub: Number(data.amountRub || 0),
    currency: ((data.currency || '').trim()).toUpperCase(),
    foreignAmount: (data.foreignAmount === '' || data.foreignAmount == null) ? null : Number(data.foreignAmount),
    rate: (data.rate === '' || data.rate == null) ? null : Number(data.rate),
    description: (data.description || '').trim(),
    createdAt: now,
    updatedAt: now,
  };
}

async function upsertTransaction(tx, rerender = true) {
  tx.updatedAt = new Date().toISOString();
  const idx = state.transactions.findIndex(item => item.id === tx.id);
  if (idx >= 0) state.transactions[idx] = tx;
  else state.transactions.push(tx);
  state.transactions.sort((a, b) =>
    b.date.localeCompare(a.date) || (b.createdAt || '').localeCompare(a.createdAt || '')
  );
  await state.db.put(TX_STORE, tx);
  if (rerender) renderAll();
}

async function saveSetting(key, value) {
  await state.db.put(SETTINGS_STORE, value, key);
}

function switchView(viewName) {
  els.tabButtons.forEach(btn => btn.classList.toggle('active', btn.dataset.view === viewName));
  els.views.forEach(v => v.classList.toggle('active', v.id === `view-${viewName}`));
}

function renderAll() {
  populatePeriodSelectors();
  populateAccountDatalists();
  renderJournal();
  renderReport();
  updateSummary();
}

function populatePeriodSelectors() {
  const years = [...new Set(state.transactions.map(t => Number(t.date.slice(0,4))))].sort((a,b) => b-a);
  const months = Array.from({ length: 12 }, (_, i) => String(i + 1).padStart(2, '0'));
  fillSelect(els.filterYear, [{ value: 'all', label: 'Все' }, ...years.map(y => ({ value: String(y), label: String(y) }))] , state.filters.year);
  fillSelect(els.reportYear, [{ value: 'all', label: 'Все' }, ...years.map(y => ({ value: String(y), label: String(y) }))], state.report.year);
  fillSelect(els.filterMonth, [{ value: 'all', label: 'Все' }, ...months.map(m => ({ value: m, label: monthLabel(m) }))], state.filters.month);
  fillSelect(els.reportMonth, [{ value: 'all', label: 'Все' }, ...months.map(m => ({ value: m, label: monthLabel(m) }))], state.report.month);
}

function fillSelect(select, options, currentValue) {
  select.innerHTML = options.map(opt => `<option value="${escapeHtml(opt.value)}">${escapeHtml(opt.label)}</option>`).join('');
  select.value = options.some(o => o.value === currentValue) ? currentValue : (options[0]?.value ?? 'all');
}

function populateAccountDatalists() {
  const accounts = getAccounts();
  const html = accounts.map(a => `<option value="${escapeHtml(a.name)}">${escapeHtml(a.name)}</option>`).join('');
  els.accountsDebitList.innerHTML = html;
  els.accountsCreditList.innerHTML = html;
}

function getAccounts() {
  const map = new Map();
  state.transactions.forEach(t => {
    if (t.debitAccount) {
      if (!map.has(t.debitAccount)) map.set(t.debitAccount, t.debitType || 'active');
    }
    if (t.creditAccount) {
      if (!map.has(t.creditAccount)) map.set(t.creditAccount, t.creditType || 'passive');
    }
  });
  return [...map.entries()].map(([name, type]) => ({ name, type })).sort((a, b) => a.name.localeCompare(b.name, 'ru'));
}

function getFilteredTransactions() {
  return state.transactions.filter(tx => {
    const year = tx.date.slice(0,4);
    const month = tx.date.slice(5,7);
    const matchesYear = state.filters.year === 'all' || year === state.filters.year;
    const matchesMonth = state.filters.month === 'all' || month === state.filters.month;
    const q = (state.filters.search || '').trim();
    const hay = `${tx.debitAccount} ${tx.creditAccount} ${tx.description}`.toLowerCase();
    const matchesSearch = !q || hay.includes(q);
    return matchesYear && matchesMonth && matchesSearch;
  });
}

function getReportTransactions() {
  return state.transactions.filter(tx => {
    const year = tx.date.slice(0,4);
    const month = tx.date.slice(5,7);
    if (state.report.periodType === 'year') return year === state.report.year;
    if (state.report.periodType === 'month') return year === state.report.year && month === state.report.month;
    return true;
  });
}

function renderJournal() {
  const txs = getFilteredTransactions();
  if (!txs.length) {
    els.journalList.innerHTML = '<div class="empty-state">Нет проводок по выбранному фильтру.</div>';
    return;
  }
  els.journalList.innerHTML = txs.map(tx => `
    <article class="journal-item" data-id="${tx.id}">
      <div class="item-topline">
        <strong>${formatDate(tx.date)}</strong>
        <button class="ghost-btn" type="button" data-edit-id="${tx.id}" style="min-height:36px;padding:0 12px;font-size:var(--text-xs);">Редактировать</button>
      </div>
      <div class="item-meta">
        <span><span class="faint">ДТ</span> <strong>${escapeHtml(tx.debitAccount)}</strong></span>
        <span><span class="faint">КТ</span> <strong>${escapeHtml(tx.creditAccount)}</strong></span>
      </div>
      <div class="item-meta">
        <strong class="money">${formatMoney(tx.amountRub)}&thinsp;${escapeHtml(state.settings.baseCurrency)}</strong>
        ${tx.currency
          ? `<span class="muted">${tx.foreignAmount != null ? formatNumber(tx.foreignAmount) : '—'}&thinsp;${escapeHtml(tx.currency)}${tx.rate ? ` · курс&thinsp;${formatNumber(tx.rate)}` : ''}</span>`
          : '<span class="faint">—</span>'}
      </div>
      ${tx.description ? `<p class="muted">${escapeHtml(tx.description)}</p>` : ''}
    </article>
  `).join('');
  els.journalList.querySelectorAll('[data-edit-id]').forEach(btn =>
    btn.addEventListener('click', () => openTransactionDialog(btn.dataset.editId))
  );
}

function renderReport() {
  const txs = getReportTransactions();
  const accountMap = new Map();
  txs.forEach(tx => {
    applyAccountMovement(accountMap, tx.debitAccount, tx.debitType || 'active', tx.amountRub, 'debit');
    applyAccountMovement(accountMap, tx.creditAccount, tx.creditType || 'passive', tx.amountRub, 'credit');
  });
  const rows = [...accountMap.values()].sort((a,b) => Math.abs(b.balance) - Math.abs(a.balance) || a.name.localeCompare(b.name,'ru'));

  if (!rows.length) {
    els.balanceList.innerHTML = '<div class="empty-state">Нет данных для отчета.</div>';
    els.accountLedger.innerHTML = '<div class="empty-state">Нет карточки счета.</div>';
    return;
  }
  els.balanceList.innerHTML = rows.map(row => `
    <article class="balance-item">
      <button type="button" data-account="${escapeHtmlAttr(row.name)}">
        <div class="balance-row">
          <strong>${escapeHtml(row.name)}</strong>
          <span class="pill">${row.type === 'active' ? 'Активный' : 'Пассивный'}</span>
        </div>
        <div class="balance-row">
          <span class="muted">Оборот ДТ&thinsp;${formatMoney(row.debitTotal)}</span>
          <span class="muted">Оборот КТ&thinsp;${formatMoney(row.creditTotal)}</span>
        </div>
        <div class="balance-row">
          <strong class="money ${row.balance >= 0 ? 'positive' : 'negative'}">Остаток&thinsp;${formatMoney(row.balance)}&thinsp;${escapeHtml(state.settings.baseCurrency)}</strong>
        </div>
      </button>
    </article>
  `).join('');
  els.balanceList.querySelectorAll('[data-account]').forEach(btn =>
    btn.addEventListener('click', () => renderLedger(btn.dataset.account))
  );
  renderLedger(state.report.selectedAccount || rows[0].name);
}

function applyAccountMovement(map, name, type, amount, side) {
  if (!name) return;
  if (!map.has(name)) map.set(name, { name, type, debitTotal: 0, creditTotal: 0, balance: 0 });
  const row = map.get(name);
  if (side === 'debit') row.debitTotal += Number(amount || 0);
  if (side === 'credit') row.creditTotal += Number(amount || 0);
  row.balance = row.type === 'passive'
    ? row.creditTotal - row.debitTotal
    : row.debitTotal - row.creditTotal;
}

function renderLedger(accountName) {
  if (!accountName) return;
  state.report.selectedAccount = accountName;
  els.ledgerTitle.textContent = accountName;
  const txs = getReportTransactions().filter(tx =>
    tx.debitAccount === accountName || tx.creditAccount === accountName
  );
  if (!txs.length) {
    els.accountLedger.innerHTML = '<div class="empty-state">По счету нет движений за выбранный период.</div>';
    return;
  }
  els.accountLedger.innerHTML = txs.map(tx => {
    const isDt = tx.debitAccount === accountName;
    const side = isDt ? 'ДТ' : 'КТ';
    const counterpart = isDt ? tx.creditAccount : tx.debitAccount;
    return `
      <article class="ledger-item">
        <div class="item-topline">
          <strong>${formatDate(tx.date)}</strong>
          <span class="pill">${side}</span>
        </div>
        <div class="item-meta">
          <span class="muted">${isDt ? 'Корр. КТ:' : 'Корр. ДТ:'} <strong>${escapeHtml(counterpart)}</strong></span>
          <strong class="money">${formatMoney(tx.amountRub)}&thinsp;${escapeHtml(state.settings.baseCurrency)}</strong>
        </div>
        ${tx.description ? `<p class="muted">${escapeHtml(tx.description)}</p>` : ''}
      </article>
    `;
  }).join('');
}

function updateSummary() {
  const filtered = getFilteredTransactions();
  els.txCount.textContent = String(filtered.length);
  els.accountCount.textContent = String(getAccounts().length);
  els.periodLabel.textContent = buildPeriodLabel();
}

function buildPeriodLabel() {
  if (state.report.periodType === 'year' && state.report.year !== 'all') return state.report.year;
  if (state.report.periodType === 'month' && state.report.year !== 'all' && state.report.month !== 'all')
    return `${monthLabel(state.report.month)} ${state.report.year}`;
  if (state.filters.year !== 'all' || state.filters.month !== 'all') {
    const parts = [];
    if (state.filters.month !== 'all') parts.push(monthLabel(state.filters.month));
    if (state.filters.year !== 'all') parts.push(state.filters.year);
    return parts.join(' ');
  }
  return 'Весь период';
}

function openTransactionDialog(id = '') {
  els.transactionForm.reset();
  els.txId.value = '';
  els.dateInput.value = todayStr();
  els.debitTypeInput.value = 'active';
  els.creditTypeInput.value = 'passive';
  els.dialogTitle.textContent = id ? 'Редактировать проводку' : 'Новая проводка';
  els.deleteTxBtn.hidden = !id;

  if (id) {
    const tx = state.transactions.find(item => item.id === id);
    if (!tx) return;
    els.txId.value = tx.id;
    els.dateInput.value = tx.date;
    els.debitAccountInput.value = tx.debitAccount;
    els.creditAccountInput.value = tx.creditAccount;
    els.debitTypeInput.value = tx.debitType || 'active';
    els.creditTypeInput.value = tx.creditType || 'passive';
    els.amountRubInput.value = tx.amountRub;
    els.currencyInput.value = tx.currency || '';
    els.foreignAmountInput.value = tx.foreignAmount != null ? tx.foreignAmount : '';
    els.rateInput.value = tx.rate != null ? tx.rate : '';
    els.descriptionInput.value = tx.description || '';
  }

  if (typeof els.transactionDialog.showModal === 'function') {
    els.transactionDialog.showModal();
  } else {
    els.transactionDialog.setAttribute('open', 'open');
  }
}

function closeDialog() {
  if (typeof els.transactionDialog.close === 'function') {
    els.transactionDialog.close();
  } else {
    els.transactionDialog.removeAttribute('open');
  }
}

async function saveTransactionFromForm(e) {
  e.preventDefault();
  const debit = els.debitAccountInput.value.trim();
  const credit = els.creditAccountInput.value.trim();
  const amountRub = parseFloat(els.amountRubInput.value || '0');

  if (!debit || !credit || !amountRub) {
    alert('Заполните счета ДТ, КТ и сумму.');
    return;
  }

  const raw = {
    date: els.dateInput.value || todayStr(),
    debitAccount: debit,
    creditAccount: credit,
    debitType: els.debitTypeInput.value,
    creditType: els.creditTypeInput.value,
    amountRub,
    currency: els.currencyInput.value,
    foreignAmount: els.foreignAmountInput.value !== '' ? els.foreignAmountInput.value : null,
    rate: els.rateInput.value !== '' ? els.rateInput.value : null,
    description: els.descriptionInput.value,
  };

  let tx;
  if (els.txId.value) {
    const original = state.transactions.find(item => item.id === els.txId.value);
    if (!original) return;
    tx = { ...mkTx(raw), id: original.id, createdAt: original.createdAt };
  } else {
    tx = mkTx(raw);
  }
  await upsertTransaction(tx, true);
  closeDialog();
}

async function deleteCurrentTransaction() {
  const id = els.txId.value;
  if (!id) return;
  if (!confirm('Удалить эту проводку?')) return;
  state.transactions = state.transactions.filter(tx => tx.id !== id);
  await state.db.delete(TX_STORE, id);
  closeDialog();
  renderAll();
}

function autoComputeRateOrForeign() {
  const currency = els.currencyInput.value.trim();
  if (!currency) return;
  const amountRub = parseFloat(els.amountRubInput.value || '0');
  const foreignAmount = parseFloat(els.foreignAmountInput.value || '0');
  const rate = parseFloat(els.rateInput.value || '0');
  const active = document.activeElement;
  if (active === els.foreignAmountInput && amountRub > 0 && foreignAmount > 0) {
    els.rateInput.value = (amountRub / foreignAmount).toFixed(4);
  } else if (active === els.rateInput && amountRub > 0 && rate > 0) {
    els.foreignAmountInput.value = (amountRub / rate).toFixed(4);
  } else if (active === els.amountRubInput && rate > 0) {
    els.foreignAmountInput.value = (amountRub / rate).toFixed(4);
  }
}

async function exportToXlsx() {
  if (typeof XLSX === 'undefined') { alert('SheetJS не загружен. Проверьте интернет-соединение.'); return; }
  const rows = state.transactions.map(tx => ({
    id: tx.id,
    date: tx.date,
    debitAccount: tx.debitAccount,
    creditAccount: tx.creditAccount,
    amountRub: tx.amountRub,
    currency: tx.currency || '',
    foreignAmount: tx.foreignAmount ?? '',
    rate: tx.rate ?? '',
    description: tx.description || '',
    accountTypeDebit: tx.debitType || 'active',
    accountTypeCredit: tx.creditType || 'passive',
    createdAt: tx.createdAt || '',
    updatedAt: tx.updatedAt || '',
  }));
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'Transactions');
  XLSX.writeFile(wb, `bookkeeper-${todayStr()}.xlsx`);
}

async function importFromXlsx(e) {
  if (typeof XLSX === 'undefined') { alert('SheetJS не загружен.'); return; }
  const file = e.target.files?.[0];
  if (!file) return;
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array', cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  const normalized = rows.map(row => ({
    id: row.id || crypto.randomUUID(),
    date: normalizeDate(row.date),
    debitAccount: String(row.debitAccount || '').trim(),
    creditAccount: String(row.creditAccount || '').trim(),
    amountRub: Number(row.amountRub || 0),
    currency: String(row.currency || '').trim().toUpperCase(),
    foreignAmount: (row.foreignAmount === '' || row.foreignAmount == null) ? null : Number(row.foreignAmount),
    rate: (row.rate === '' || row.rate == null) ? null : Number(row.rate),
    description: String(row.description || '').trim(),
    debitType: row.accountTypeDebit === 'passive' ? 'passive' : 'active',
    creditType: row.accountTypeCredit === 'active' ? 'active' : 'passive',
    createdAt: row.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  })).filter(row => row.date && row.debitAccount && row.creditAccount && row.amountRub);

  if (!normalized.length) { alert('Не удалось найти валидные строки. Проверьте формат файла.'); return; }
  for (const tx of normalized) await upsertTransaction(tx, false);
  await loadAllData();
  renderAll();
  e.target.value = '';
  alert(`Импортировано ${normalized.length} строк.`);
}

async function clearAllData() {
  if (!confirm('Очистить все проводки и настройки? Это действие нельзя отменить.')) return;
  const tx = state.db.transaction([TX_STORE, SETTINGS_STORE], 'readwrite');
  await tx.objectStore(TX_STORE).clear();
  await tx.objectStore(SETTINGS_STORE).clear();
  await tx.done;
  state.transactions = [];
  state.settings = { baseCurrency: 'RUB' };
  els.baseCurrencyInput.value = 'RUB';
  renderAll();
}

function registerServiceWorker() {
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () =>
      navigator.serviceWorker.register('sw.js').catch(err => console.warn('SW registration failed:', err))
    );
  }
}

function setupInstallPrompt() {
  window.addEventListener('beforeinstallprompt', e => {
    e.preventDefault();
    state.deferredPrompt = e;
    els.installBtn.hidden = false;
  });
  els.installBtn.addEventListener('click', async () => {
    if (!state.deferredPrompt) return;
    await state.deferredPrompt.prompt();
    state.deferredPrompt = null;
    els.installBtn.hidden = true;
  });
}

function todayStr() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

function monthLabel(m) {
  const names = ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
  const idx = Number(m) - 1;
  return (idx >= 0 && idx < 12) ? names[idx] : 'Все';
}

function formatDate(v) {
  const [y,m,d] = v.split('-');
  return `${d}.${m}.${y}`;
}

function formatMoney(value) {
  return new Intl.NumberFormat('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(Number(value || 0));
}

function formatNumber(value) {
  return new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 4 }).format(Number(value || 0));
}

function normalizeDate(v) {
  if (!v) return todayStr();
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth()+1).padStart(2,'0');
    const d = String(v.getDate()).padStart(2,'0');
    return `${y}-${m}-${d}`;
  }
  if (typeof v === 'number') {
    try {
      const parsed = XLSX.SSF.parse_date_code(v);
      return `${parsed.y}-${String(parsed.m).padStart(2,'0')}-${String(parsed.d).padStart(2,'0')}`;
    } catch { return todayStr(); }
  }
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (/^\d{2}\.\d{2}\.\d{4}$/.test(s)) {
    const [d,m,y] = s.split('.');
    return `${y}-${m}-${d}`;
  }
  return todayStr();
}

function escapeHtml(str) {
  return String(str ?? '').replace(/[&<>"']/g, s => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  }[s]));
}
function escapeHtmlAttr(str) { return escapeHtml(str); }
