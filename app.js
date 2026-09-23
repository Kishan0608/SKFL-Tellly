// Tally Pro Executive Financial Suite Logic - Strict Rakesh Bhai Account Filtering Edition

// Load Saved State from localStorage (or fallback to empty null defaults)
let transactions = JSON.parse(localStorage.getItem('tally_entries')) || [];

let accountOpening = JSON.parse(localStorage.getItem('tally_rakesh_opening')) || {
  cash: null,
  idbi: null
};

let userTargetClosing = JSON.parse(localStorage.getItem('tally_rakesh_target_closing')) || {
  cashTarget: null,
  idbiTarget: null
};

// DOM Elements
const dateFilter = document.getElementById('dateFilter');

// Cash Reconciliation Elements
const cashOpeningInput = document.getElementById('cashOpeningInput'); // Field 1: User Input
const expectedCashInput = document.getElementById('expectedCashInput'); // Field 2: User Input
const tallyCashOpeningDisplay = document.getElementById('tallyCashOpeningDisplay'); // Field 3: Auto Default
const tallyCashClosingDisplay = document.getElementById('tallyCashClosingDisplay'); // Field 4: Auto Calculated

// IDBI Reconciliation Elements
const idbiOpeningInput = document.getElementById('idbiOpeningInput'); // Field 1: User Input
const expectedIdbiInput = document.getElementById('expectedIdbiInput'); // Field 2: User Input
const tallyIdbiOpeningDisplay = document.getElementById('tallyIdbiOpeningDisplay'); // Field 3: Auto Default
const tallyIdbiClosingDisplay = document.getElementById('tallyIdbiClosingDisplay'); // Field 4: Auto Calculated

const tableBody = document.getElementById('tableBody');

// Summary Matrix Elements
const summaryCashOpening = document.getElementById('summaryCashOpening');
const summaryCashIncome = document.getElementById('summaryCashIncome');
const summaryCashExpense = document.getElementById('summaryCashExpense');
const summaryCashClosing = document.getElementById('summaryCashClosing');

const summaryIdbiOpening = document.getElementById('summaryIdbiOpening');
const summaryIdbiIncome = document.getElementById('summaryIdbiIncome');
const summaryIdbiExpense = document.getElementById('summaryIdbiExpense');
const summaryIdbiClosing = document.getElementById('summaryIdbiClosing');

const summaryTotalOpening = document.getElementById('summaryTotalOpening');
const summaryTotalIncome = document.getElementById('summaryTotalIncome');
const summaryTotalExpense = document.getElementById('summaryTotalExpense');
const summaryTotalClosing = document.getElementById('summaryTotalClosing');

const cashDiffBadge = document.getElementById('cashDiffBadge');
const idbiDiffBadge = document.getElementById('idbiDiffBadge');

// Modal Elements
const modalOverlay = document.getElementById('modalOverlay');
const pasteModalOverlay = document.getElementById('pasteModalOverlay');
const entryForm = document.getElementById('entryForm');
const excelPasteArea = document.getElementById('excelPasteArea');
const pastePreviewContainer = document.getElementById('pastePreviewContainer');

let pendingParsedEntries = [];
let isPasteModalOpen = false;

// Initialize App
document.addEventListener('DOMContentLoaded', () => {
  populateDateDropdown();
  renderTable();
  setupEventListeners();
  setupGlobalExcelPaste();
});

function populateDateDropdown() {
  const dates = new Set();
  transactions.forEach(t => { if (t.date) dates.add(t.date); });

  dateFilter.innerHTML = '<option value="ALL">All Dates Combined</option>';
  Array.from(dates).sort().reverse().forEach(d => {
    const opt = document.createElement('option');
    opt.value = d;
    opt.textContent = `📅 Date: ${d}`;
    dateFilter.appendChild(opt);
  });
}

// Strict Account Verification Helpers (Only RAKESH BHAI accounts are included)
function isRakeshCash(accountStr) {
  if (!accountStr) return false;
  const str = accountStr.toUpperCase().trim();
  // Reject other people / third-party accounts
  if (str.includes('TEJAS') || str.includes('NARESH') || str.includes('RAMESH') || str.includes('RAMALA') || str.includes('NIKANJ') || str.includes('ABRAJ') || str.includes('RAJESH') || str.includes('ICICI')) {
    return false;
  }
  return (str.includes('RAKESH') && str.includes('CASH')) || str === 'RAKESHBHAI CASH' || str === 'RAKESH CASH';
}

function isRakeshIdbi(accountStr) {
  if (!accountStr) return false;
  const str = accountStr.toUpperCase().trim();
  // Reject other people / third-party accounts
  if (str.includes('TEJAS') || str.includes('NARESH') || str.includes('RAMESH') || str.includes('ICICI') || str.includes('HDFC') || str.includes('SBI')) {
    return false;
  }
  return (str.includes('RAKESH') && str.includes('IDBI')) || str.includes('RAKESH IDBI');
}

function renderTable() {
  // Field 1 User Inputs (Blank if null/0)
  cashOpeningInput.value = (accountOpening.cash !== null && accountOpening.cash !== undefined && accountOpening.cash !== 0) ? accountOpening.cash : '';
  idbiOpeningInput.value = (accountOpening.idbi !== null && accountOpening.idbi !== undefined && accountOpening.idbi !== 0) ? accountOpening.idbi : '';

  // Field 2 User Inputs (NEVER AUTO POPULATED — ONLY SHOWS IF USER TYPED IT)
  expectedCashInput.value = (userTargetClosing.cashTarget !== null && userTargetClosing.cashTarget !== undefined) ? userTargetClosing.cashTarget : '';
  expectedIdbiInput.value = (userTargetClosing.idbiTarget !== null && userTargetClosing.idbiTarget !== undefined) ? userTargetClosing.idbiTarget : '';

  const selectedDate = dateFilter.value;
  let filtered = transactions;
  if (selectedDate !== 'ALL') {
    filtered = transactions.filter(t => t.date === selectedDate);
  }

  tableBody.innerHTML = '';

  let totalExpense = 0;
  let totalIncome = 0;

  let cashExpense = 0;
  let cashIncome = 0;
  let idbiExpense = 0;
  let idbiIncome = 0;

  const currentCashOp = (accountOpening.cash || 0);
  const currentIdbiOp = (accountOpening.idbi || 0);
  const totalOpening = currentCashOp + currentIdbiOp;
  let runningBal = totalOpening;

  if (filtered.length === 0) {
    tableBody.innerHTML = `
      <tr>
        <td colspan="13" class="empty-state">
          No entries found. Press <strong>Ctrl + V</strong> anywhere to paste your Excel rows or click "+ Add Entry".
        </td>
      </tr>
    `;
  } else {
    filtered.forEach((t, index) => {
      let exp = parseFloat(t.expenseAmount) || 0;
      let inc = parseFloat(t.incomeAmount) || 0;

      const typeStr = (t.type || '').toUpperCase();
      const particularStr = (t.particular || '').toUpperCase();

      const isIncomeType = typeStr.includes('INCOME') || typeStr.includes('RECEIPT') || particularStr.includes('INCOME');
      if (isIncomeType && exp > 0 && inc === 0) {
        inc = exp;
        exp = 0;
      }

      totalExpense += exp;
      totalIncome += inc;

      const isCashPaid = isRakeshCash(t.paidBy);
      const isCashRecv = isRakeshCash(t.receivedBy);
      const isIdbiPaid = isRakeshIdbi(t.paidBy);
      const isIdbiRecv = isRakeshIdbi(t.receivedBy);

      // Aggregate ONLY RAKESH BHAI Accounts
      if (exp > 0) {
        if (isCashPaid) cashExpense += exp;
        if (isIdbiPaid) idbiExpense += exp;
      }

      if (inc > 0) {
        if (isCashRecv) cashIncome += inc;
        if (isIdbiRecv) idbiIncome += inc;
      }

      runningBal = runningBal + inc - exp;

      let paidBadge = t.paidBy || '-';
      if (isCashPaid) paidBadge = '<span class="account-tag cash">💵 CASH</span> ' + t.paidBy;
      else if (isIdbiPaid) paidBadge = '<span class="account-tag idbi">🏦 BANK</span> ' + t.paidBy;

      let recvBadge = t.receivedBy || '-';
      if (isCashRecv) recvBadge = '<span class="account-tag cash">💵 CASH</span> ' + t.receivedBy;
      else if (isIdbiRecv) recvBadge = '<span class="account-tag idbi">🏦 BANK</span> ' + t.receivedBy;

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${index + 1}</td>
        <td><strong>${t.date}</strong></td>
        <td>${t.firmName || '-'}</td>
        <td><span style="color:var(--blue-text); font-weight:700;">${t.partyLedger || '-'}</span></td>
        <td>${t.particular || '-'}</td>
        <td><span class="type-badge">${t.type || 'GENERAL'}</span></td>
        <td class="col-expense">${paidBadge}</td>
        <td class="col-expense">${exp > 0 ? '₹ ' + exp.toLocaleString('en-IN') : '-'}</td>
        <td class="col-income">${recvBadge}</td>
        <td class="col-income">${inc > 0 ? '₹ ' + inc.toLocaleString('en-IN') : '-'}</td>
        <td>${t.site || '-'}</td>
        <td class="running-balance">₹ ${runningBal.toLocaleString('en-IN')}</td>
        <td>
          <button class="btn btn-danger btn-sm" onclick="deleteEntry(${t.id})">🗑️</button>
        </td>
      `;
      tableBody.appendChild(tr);
    });
  }

  // FIELD 3: Auto default same as Field 1 Opening
  const tallyCashOpening = currentCashOp;
  const tallyIdbiOpening = currentIdbiOp;

  // FIELD 4: Auto calculated based on provided entries (Field 3 Opening + Income - Expense)
  const tallyCalculatedCashClosing = tallyCashOpening + cashIncome - cashExpense;
  const tallyCalculatedIdbiClosing = tallyIdbiOpening + idbiIncome - idbiExpense;

  const calcTotalIncome = cashIncome + idbiIncome;
  const calcTotalExpense = cashExpense + idbiExpense;

  // Render Accounts Matrix Table
  summaryCashOpening.textContent = currentCashOp.toLocaleString('en-IN');
  summaryCashIncome.textContent = cashIncome.toLocaleString('en-IN');
  summaryCashExpense.textContent = cashExpense.toLocaleString('en-IN');
  summaryCashClosing.textContent = tallyCalculatedCashClosing.toLocaleString('en-IN');

  summaryIdbiOpening.textContent = currentIdbiOp.toLocaleString('en-IN');
  summaryIdbiIncome.textContent = idbiIncome.toLocaleString('en-IN');
  summaryIdbiExpense.textContent = idbiExpense.toLocaleString('en-IN');
  summaryIdbiClosing.textContent = tallyCalculatedIdbiClosing.toLocaleString('en-IN');

  summaryTotalOpening.textContent = totalOpening.toLocaleString('en-IN');
  summaryTotalIncome.textContent = calcTotalIncome.toLocaleString('en-IN');
  summaryTotalExpense.textContent = calcTotalExpense.toLocaleString('en-IN');
  summaryTotalClosing.textContent = (tallyCalculatedCashClosing + tallyCalculatedIdbiClosing).toLocaleString('en-IN');

  // Render Field 3 (Auto Default Opening) & Field 4 (Auto Calculated Closing)
  tallyCashOpeningDisplay.textContent = '₹ ' + tallyCashOpening.toLocaleString('en-IN');
  tallyCashClosingDisplay.textContent = '₹ ' + tallyCalculatedCashClosing.toLocaleString('en-IN');

  tallyIdbiOpeningDisplay.textContent = '₹ ' + tallyIdbiOpening.toLocaleString('en-IN');
  tallyIdbiClosingDisplay.textContent = '₹ ' + tallyCalculatedIdbiClosing.toLocaleString('en-IN');

  // Reconciliation Spotter Logic: Compare Field 2 (User's Typed Target) vs Field 4 (Entries Auto-Calculated Closing)
  const userCashTyped = (userTargetClosing.cashTarget !== null && userTargetClosing.cashTarget !== undefined);
  const userIdbiTyped = (userTargetClosing.idbiTarget !== null && userTargetClosing.idbiTarget !== undefined);

  if (userCashTyped) {
    const cashDiff = tallyCalculatedCashClosing - userTargetClosing.cashTarget;
    if (cashDiff === 0) {
      cashDiffBadge.className = 'diff-pill matched';
      cashDiffBadge.innerHTML = '✅ OK';
    } else {
      cashDiffBadge.className = 'diff-pill mismatch';
      cashDiffBadge.innerHTML = `⚠️ MISMATCH: ₹${cashDiff > 0 ? '+' : ''}${cashDiff.toLocaleString('en-IN')}`;
    }
  } else {
    cashDiffBadge.className = 'diff-pill matched';
    cashDiffBadge.innerHTML = 'Ready for Input';
  }

  if (userIdbiTyped) {
    const idbiDiff = tallyCalculatedIdbiClosing - userTargetClosing.idbiTarget;
    if (idbiDiff === 0) {
      idbiDiffBadge.className = 'diff-pill matched';
      idbiDiffBadge.innerHTML = '✅ OK';
    } else {
      idbiDiffBadge.className = 'diff-pill mismatch';
      idbiDiffBadge.innerHTML = `⚠️ MISMATCH: ₹${idbiDiff > 0 ? '+' : ''}${idbiDiff.toLocaleString('en-IN')}`;
    }
  } else {
    idbiDiffBadge.className = 'diff-pill matched';
    idbiDiffBadge.innerHTML = 'Ready for Input';
  }
}

// Save State to LocalStorage
function saveData() {
  localStorage.setItem('tally_entries', JSON.stringify(transactions));
  localStorage.setItem('tally_rakesh_opening', JSON.stringify(accountOpening));
  localStorage.setItem('tally_rakesh_target_closing', JSON.stringify(userTargetClosing));
}

// User Input Event Listeners for Fields 1 & 2
cashOpeningInput.addEventListener('input', (e) => {
  const val = e.target.value.trim();
  accountOpening.cash = val === '' ? null : parseFloat(val);
  saveData();
  renderTable();
});

expectedCashInput.addEventListener('input', (e) => {
  const val = e.target.value.trim();
  userTargetClosing.cashTarget = val === '' ? null : parseFloat(val);
  saveData();
  renderTable();
});

idbiOpeningInput.addEventListener('input', (e) => {
  const val = e.target.value.trim();
  accountOpening.idbi = val === '' ? null : parseFloat(val);
  saveData();
  renderTable();
});

expectedIdbiInput.addEventListener('input', (e) => {
  const val = e.target.value.trim();
  userTargetClosing.idbiTarget = val === '' ? null : parseFloat(val);
  saveData();
  renderTable();
});

// App Event Listeners
function setupEventListeners() {
  dateFilter.addEventListener('change', renderTable);

  document.getElementById('addEntryBtn').addEventListener('click', () => {
    document.getElementById('entryDate').value = new Date().toISOString().split('T')[0];
    document.getElementById('partyLedger').value = 'Rakesh Bhai';
    modalOverlay.classList.add('active');
  });

  document.getElementById('cancelModalBtn').addEventListener('click', closeModal);
  document.getElementById('closeModalBtn').addEventListener('click', closeModal);

  document.getElementById('pasteExcelBtn').addEventListener('click', openPasteModal);
  document.getElementById('closePasteModalBtn').addEventListener('click', closePasteModal);
  document.getElementById('cancelPasteModalBtn').addEventListener('click', closePasteModal);

  excelPasteArea.addEventListener('input', () => {
    parseAndPreviewExcelText(excelPasteArea.value);
  });

  document.getElementById('importPasteBtn').addEventListener('click', () => {
    if (pendingParsedEntries.length === 0) {
      alert('Please paste valid Excel rows first.');
      return;
    }

    transactions = [...pendingParsedEntries];
    saveData();
    populateDateDropdown();
    renderTable();
    closePasteModal();
    alert(`Successfully imported ${pendingParsedEntries.length} entries!`);
    pendingParsedEntries = [];
    excelPasteArea.value = '';
    pastePreviewContainer.innerHTML = '';
  });

  entryForm.addEventListener('submit', (e) => {
    e.preventDefault();

    const expAmt = parseFloat(document.getElementById('expenseAmount').value) || 0;
    const incAmt = parseFloat(document.getElementById('incomeAmount').value) || 0;

    const newEntry = {
      id: Date.now(),
      date: formatDateString(document.getElementById('entryDate').value),
      firmName: document.getElementById('firmName').value,
      partyLedger: document.getElementById('partyLedger').value,
      particular: document.getElementById('particular').value,
      type: document.getElementById('entryType').value,
      paidBy: document.getElementById('paidBy').value,
      expenseAmount: expAmt,
      receivedBy: document.getElementById('receivedBy').value,
      incomeAmount: incAmt,
      site: document.getElementById('site').value
    };

    transactions.push(newEntry);
    saveData();
    populateDateDropdown();
    renderTable();
    closeModal();
    entryForm.reset();
  });

  // COMPLETE RESET DATA BUTTON
  document.getElementById('resetDataBtn').addEventListener('click', () => {
    localStorage.clear();
    transactions = [];
    accountOpening = { cash: null, idbi: null };
    userTargetClosing = { cashTarget: null, idbiTarget: null };
    
    cashOpeningInput.value = '';
    expectedCashInput.value = '';
    idbiOpeningInput.value = '';
    expectedIdbiInput.value = '';

    populateDateDropdown();
    renderTable();
  });

  document.getElementById('exportCsvBtn').addEventListener('click', exportToCSV);
}

function setupGlobalExcelPaste() {
  window.addEventListener('paste', (e) => {
    if (e.target.tagName === 'INPUT' || (e.target.tagName === 'TEXTAREA' && e.target.id !== 'excelPasteArea')) {
      return;
    }
    if (isPasteModalOpen) return;

    const pastedText = e.clipboardData ? e.clipboardData.getData('text') : '';
    if (pastedText && pastedText.trim().length > 0) {
      e.preventDefault();
      openPasteModal();
      excelPasteArea.value = pastedText;
      parseAndPreviewExcelText(pastedText);
    }
  });
}

function openPasteModal() { isPasteModalOpen = true; pasteModalOverlay.classList.add('active'); excelPasteArea.focus(); }
function closePasteModal() { isPasteModalOpen = false; pasteModalOverlay.classList.remove('active'); }
function closeModal() { modalOverlay.classList.remove('active'); }

function deleteEntry(id) {
  transactions = transactions.filter(t => t.id !== id);
  saveData();
  populateDateDropdown();
  renderTable();
}

function formatDateString(dateStr) {
  if (!dateStr) return '';
  const parts = dateStr.split('-');
  return parts.length === 3 ? `${parts[2]}-${parts[1]}-${parts[0]}` : dateStr;
}

function cleanNumber(val) {
  if (val === null || val === undefined) return 0;
  const str = String(val).trim();
  if (!str) return 0;
  const numStr = str.replace(/[^0-9.-]/g, '');
  const parsed = parseFloat(numStr);
  return isNaN(parsed) ? 0 : parsed;
}

function isPureNumber(val) {
  if (!val) return false;
  return /^[0-9.,\s₹$-]+$/.test(String(val).trim());
}

function parseAndPreviewExcelText(text) {
  if (!text || !text.trim()) {
    pendingParsedEntries = [];
    pastePreviewContainer.innerHTML = '';
    return;
  }

  const lines = text.trim().split(/\r?\n/);
  const rawList = [];

  lines.forEach((line, idx) => {
    if (idx === 0 && (line.toLowerCase().includes('date') && line.toLowerCase().includes('party'))) {
      return;
    }

    const rawCols = line.includes('\t') ? line.split('\t') : line.split(',');
    const cols = rawCols.map(c => c ? c.trim().replace(/^"(.*)"$/, '$1') : '');

    if (cols.every(c => !c)) return;

    const date = cols[0] || new Date().toISOString().split('T')[0];
    const firmName = cols[1] || '';
    const partyLedger = cols[2] || 'Rakesh Bhai';
    const particular = cols[3] || '';
    const type = cols[4] || 'GENERAL';

    let paidBy = cols[5] || '';
    if (isPureNumber(paidBy)) paidBy = '';

    const expenseAmount = cleanNumber(cols[6]);

    let receivedBy = cols[7] || '';
    if (isPureNumber(receivedBy)) receivedBy = '';

    const incomeAmount = cleanNumber(cols[8]);

    let site = '';
    for (let sIdx = 9; sIdx < cols.length; sIdx++) {
      const candidate = cols[sIdx] ? cols[sIdx].trim() : '';
      if (candidate && !isPureNumber(candidate) && candidate.toLowerCase() !== partyLedger.toLowerCase()) {
        site = candidate;
        break;
      }
    }

    const entry = {
      id: Date.now() + Math.random(),
      date,
      firmName,
      partyLedger,
      particular: particular || 'Excel Entry',
      type: type || 'GENERAL',
      paidBy,
      expenseAmount,
      receivedBy,
      incomeAmount,
      site
    };

    rawList.push(entry);
  });

  // Unique Deduplication
  const uniqueEntries = [];
  const seenKeys = new Set();

  rawList.forEach(entry => {
    const key = `${entry.date}_${entry.partyLedger}_${entry.particular}_${entry.expenseAmount}_${entry.incomeAmount}_${entry.paidBy}_${entry.receivedBy}`;
    if (!seenKeys.has(key)) {
      seenKeys.add(key);
      uniqueEntries.push(entry);
    }
  });

  pendingParsedEntries = uniqueEntries;

  let html = `
    <table class="data-table">
      <thead>
        <tr>
          <th>#</th><th>DATE</th><th>FIRM NAME</th><th>PARTY LEDGER</th><th>PARTICULAR</th><th>TYPE</th><th>PAID BY</th><th>EXPENSE (₹)</th><th>RECEIVED BY</th><th>INCOME (₹)</th><th>SITE</th>
        </tr>
      </thead>
      <tbody>
  `;

  pendingParsedEntries.forEach((entry, idx) => {
    html += `
      <tr>
        <td>${idx + 1}</td>
        <td>${entry.date}</td>
        <td>${entry.firmName || '-'}</td>
        <td><strong>${entry.partyLedger}</strong></td>
        <td>${entry.particular}</td>
        <td>${entry.type}</td>
        <td>${entry.paidBy || '-'}</td>
        <td class="col-expense">${entry.expenseAmount > 0 ? '₹ ' + entry.expenseAmount.toLocaleString('en-IN') : '-'}</td>
        <td>${entry.receivedBy || '-'}</td>
        <td class="col-income">${entry.incomeAmount > 0 ? '₹ ' + entry.incomeAmount.toLocaleString('en-IN') : '-'}</td>
        <td>${entry.site || '-'}</td>
      </tr>
    `;
  });

  html += '</tbody></table>';

  if (pendingParsedEntries.length === 0) {
    pastePreviewContainer.innerHTML = '<div style="padding: 15px; color: var(--text-muted);">No valid rows detected. Make sure text is copied from Excel.</div>';
  } else {
    pastePreviewContainer.innerHTML = html;
  }
}

function exportToCSV() {
  const selectedDate = dateFilter.value;
  let filtered = transactions;
  if (selectedDate !== 'ALL') filtered = transactions.filter(t => t.date === selectedDate);

  const headers = ['SR NO', 'DATE', 'FIRM NAME', 'PARTY LEDGER', 'PARTICULAR', 'TYPE', 'PAID BY (EXPENSE)', 'EXPENSE AMOUNT', 'RECEIVED BY (INCOME)', 'INCOME AMOUNT', 'SITE'];
  const rows = filtered.map((t, index) => [
    index + 1, `"${t.date}"`, `"${t.firmName}"`, `"${t.partyLedger}"`, `"${t.particular}"`, `"${t.type}"`, `"${t.paidBy}"`, t.expenseAmount, `"${t.receivedBy}"`, t.incomeAmount, `"${t.site}"`
  ]);

  let csvContent = "data:text/csv;charset=utf-8," + [headers.join(','), ...rows.map(e => e.join(','))].join('\n');
  const encodedUri = encodeURI(csvContent);
  const link = document.createElement('a');
  link.setAttribute('href', encodedUri);
  link.setAttribute('download', `Rakesh_Bhai_Tally_Entries_${selectedDate}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}
