document.addEventListener('DOMContentLoaded', () => {
  initSidebarToggle();
  initGlobalMonthSelector();
  initNavigation();
  initPeopleManager();
  initMonthlySheets();
  initWorksTracker();
  initClientsTracker();
  initExpensesTracker();
  initSettlement();
  initInvoices();
  initMobileInputVisibilityHandling();
  initSettings();
  
  // Initial render
  renderAll();

  // Re-render on state changes
  window.addEventListener('appStateChanged', () => {
    applySettings();
    renderAll();
  });

  // Initial apply
  applySettings();
  if (typeof firebase !== 'undefined' && firebase.auth) {
    initFirebaseAuth();
  }
});

function initSidebarToggle() {
  const btn = document.getElementById('btn-toggle-sidebar');
  const sidebar = document.getElementById('sidebar');
  const mainContent = document.getElementById('main-content');

  if (!btn || !sidebar || !mainContent) return;

  const isCollapsed = localStorage.getItem('sidebar-collapsed') === 'true';
  if (isCollapsed) {
    sidebar.classList.add('collapsed');
    mainContent.classList.add('expanded');
    btn.innerHTML = '<i data-lucide="panel-left-open" id="sidebar-toggle-icon"></i>';
  }

  btn.addEventListener('click', () => {
    sidebar.classList.toggle('collapsed');
    mainContent.classList.toggle('expanded');
    
    const currentlyCollapsed = sidebar.classList.contains('collapsed');
    localStorage.setItem('sidebar-collapsed', currentlyCollapsed);
    
    btn.innerHTML = `<i data-lucide="${currentlyCollapsed ? 'panel-left-open' : 'panel-left-close'}" id="sidebar-toggle-icon"></i>`;
    lucide.createIcons();
  });
}

function initNavigation() {
  const navLinks = document.querySelectorAll('.nav-item');

  navLinks.forEach(link => {
    link.addEventListener('click', (e) => {
      e.preventDefault();

      activateNavigationTarget(link.getAttribute('data-target'));
    });
  });
}

function activateNavigationTarget(targetId) {
  const navLinks = document.querySelectorAll('.nav-item');
  const views = document.querySelectorAll('.view');
  const mainContent = document.getElementById('main-content');
  if (!targetId) return;

  navLinks.forEach(link => {
    link.classList.toggle('active', link.getAttribute('data-target') === targetId);
  });
  views.forEach(view => {
    view.classList.toggle('active', view.id === targetId);
  });

  if (mainContent) {
    mainContent.scrollTop = 0;
  }

  if (targetId === 'hours-view' && window.currentSheetId) {
    renderSheetDetail(window.currentSheetId);
  } else if (targetId === 'works-view' && window.currentWorksSheetId) {
    renderWorksSheetDetail(window.currentWorksSheetId);
  } else if (targetId === 'invoices-view') {
    renderInvoices();
  }
}

function renderAll() {
  renderGlobalMonthSelector();
  renderSummary();
  renderPeople();
  renderMonthlySheets();
  renderWorksSheets();
  renderClients();
  renderExpenses();
  renderSettlement();
  renderInvoices();
  renderInvoiceEqualization();
  lucide.createIcons();
}

function getSelectedMonthKey() {
  return Store.getSelectedMonth ? Store.getSelectedMonth() : Store.getState().selectedMonth;
}

function formatMonthLabel(monthKey) {
  if (!monthKey || !monthKey.includes('-')) return monthKey || '';
  const [year, month] = monthKey.split('-').map(Number);
  const date = new Date(year, (month || 1) - 1, 1);
  return date.toLocaleDateString('pl-PL', { month: 'long', year: 'numeric' });
}

function renderGlobalMonthSelector() {
  const monthInput = document.getElementById('global-month-select');
  if (monthInput) {
    monthInput.value = getSelectedMonthKey();
    monthInput.dataset.lastValidMonth = monthInput.value;
  }

  const summarySubtitle = document.getElementById('summary-subtitle');
  if (summarySubtitle) {
    summarySubtitle.textContent = `Podsumowanie ${formatMonthLabel(getSelectedMonthKey())}`;
  }

  const settlementSubtitle = document.getElementById('settlement-subtitle');
  if (settlementSubtitle) {
    settlementSubtitle.textContent = `Kalkulacja zarobku, kosztów i wypłat ${formatMonthLabel(getSelectedMonthKey())}`;
  }
}

function initGlobalMonthSelector() {
  const monthInput = document.getElementById('global-month-select');
  if (!monthInput) return;

  const syncMonthInputValue = (value) => {
    const safeValue = value || monthInput.dataset.lastValidMonth || getSelectedMonthKey();
    monthInput.value = safeValue;
    monthInput.dataset.lastValidMonth = safeValue;
    return safeValue;
  };

  const openMonthPicker = () => {
    if (typeof monthInput.showPicker !== 'function') return;
    try {
      monthInput.showPicker();
    } catch {
      // Ignore browsers that restrict programmatic picker opening.
    }
  };

  monthInput.required = true;
  syncMonthInputValue(getSelectedMonthKey());

  monthInput.addEventListener('click', openMonthPicker);
  monthInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ' || e.key === 'ArrowDown') {
      e.preventDefault();
      openMonthPicker();
    }
  });

  monthInput.addEventListener('input', () => {
    if (!monthInput.value) {
      syncMonthInputValue();
    }
  });

  monthInput.addEventListener('blur', () => {
    if (!monthInput.value) {
      syncMonthInputValue();
    }
  });

  monthInput.addEventListener('change', (e) => {
    const nextMonth = syncMonthInputValue(e.target.value);
    window.currentSheetId = null;
    Store.setSelectedMonth(nextMonth);
  });
}

function getDefaultDateForSelectedMonth() {
  const selectedMonth = getSelectedMonthKey();
  const today = new Date().toISOString().split('T')[0];
  return today.startsWith(`${selectedMonth}-`) ? today : `${selectedMonth}-01`;
}

function getMonthlySheetPersonActivityOverride(sheet, dayKey, personId) {
  return sheet?.days?.[dayKey]?.activityOverrides?.[personId] || null;
}

function hasPreviousContinuousMonthlySheetPersonActivityOverride(sheet, dayKey, personId, override) {
  if (!sheet?.days || !personId || !override) return false;

  const startDay = parseInt(dayKey, 10);
  if (!Number.isFinite(startDay) || startDay <= 1) return false;

  const oppositeOverride = override === 'inactive' ? 'active' : 'inactive';

  for (let day = startDay - 1; day >= 1; day--) {
    const previousOverride = getMonthlySheetPersonActivityOverride(sheet, String(day), personId);
    if (!previousOverride) continue;
    if (previousOverride === oppositeOverride) return false;
    if (previousOverride === override) return true;
  }

  return false;
}

function pruneFutureMonthlySheetPersonActivityOverrides(sheet, dayKey, personId, override) {
  if (!sheet?.days || !personId || !override) return;

  const startDay = parseInt(dayKey, 10);
  if (!Number.isFinite(startDay)) return;

  Object.entries(sheet.days).forEach(([nextDayKey, nextDayData]) => {
    const nextDay = parseInt(nextDayKey, 10);
    if (!Number.isFinite(nextDay) || nextDay <= startDay) return;
    if (!nextDayData?.activityOverrides || nextDayData.activityOverrides[personId] !== override) return;

    delete nextDayData.activityOverrides[personId];
    if (Object.keys(nextDayData.activityOverrides).length === 0) {
      delete nextDayData.activityOverrides;
    }
  });
}

function setMonthlySheetPersonActivityOverride(sheet, dayKey, personId, override = null) {
  const dayData = sheet?.days?.[dayKey];
  if (!dayData) return;

  if (override === null) {
    if (!dayData.activityOverrides) return;
    delete dayData.activityOverrides[personId];
    if (Object.keys(dayData.activityOverrides).length === 0) {
      delete dayData.activityOverrides;
    }
    return;
  }

  if (hasPreviousContinuousMonthlySheetPersonActivityOverride(sheet, dayKey, personId, override)) {
    if (dayData.activityOverrides) {
      delete dayData.activityOverrides[personId];
      if (Object.keys(dayData.activityOverrides).length === 0) {
        delete dayData.activityOverrides;
      }
    }
    return;
  }

  if (!dayData.activityOverrides) dayData.activityOverrides = {};

  dayData.activityOverrides[personId] = override;
  pruneFutureMonthlySheetPersonActivityOverrides(sheet, dayKey, personId, override);
}

function isMonthlySheetPersonInactiveOnDay(sheet, personId, dayNumber) {
  if (!sheet?.days || !personId) return false;

  let isInactive = false;
  for (let day = 1; day <= dayNumber; day++) {
    const override = getMonthlySheetPersonActivityOverride(sheet, String(day), personId);
    if (override === 'inactive') {
      isInactive = true;
    } else if (override === 'active') {
      isInactive = false;
    }
  }

  return isInactive;
}

function clearMonthlySheetPersonManualFlag(dayData, personId) {
  if (!dayData?.manual) return;

  delete dayData.manual[personId];
  if (Object.keys(dayData.manual).length === 0) {
    delete dayData.manual;
  }
}

function syncMonthlySheetDayPersonHours(sheet, dayKey, personId) {
  const dayData = sheet?.days?.[dayKey];
  if (!dayData) return;
  if (!dayData.hours) dayData.hours = {};

  const dayNumber = parseInt(dayKey, 10);
  if (isMonthlySheetPersonInactiveOnDay(sheet, personId, dayNumber)) {
    dayData.hours[personId] = 0;
    return;
  }

  if (dayData.isWholeTeamChecked && dayData.globalStart && dayData.globalEnd) {
    const calcH = Calculations.calculateHours(dayData.globalStart, dayData.globalEnd);
    if (calcH > 0) {
      dayData.hours[personId] = calcH;
      return;
    }
  }

  delete dayData.hours[personId];
}

function sheetHasPersonData(sheet, personId) {
  if (sheet?.personsConfig?.[personId]) return true;
  if (!sheet?.days) return false;

  return Object.values(sheet.days).some(day => day?.hours && day.hours[personId] !== undefined && day.hours[personId] !== '');
}

function getVisiblePersonsForSheet(state, sheet) {
  const month = sheet?.month || getSelectedMonthKey();
  const configuredActivePersonIds = Array.isArray(sheet?.activePersons)
    ? new Set(sheet.activePersons)
    : null;
  const fallbackActivePersonIds = new Set(
    (state.persons || [])
      .filter(person => Store.isPersonActiveInMonth(person.id, month))
      .map(person => person.id)
  );

  return state.persons.filter(person => {
    if (configuredActivePersonIds) {
      return configuredActivePersonIds.has(person.id);
    }
    return fallbackActivePersonIds.has(person.id) || sheetHasPersonData(sheet, person.id);
  });
}

function getEmployerCandidates(state) {
  return (state.persons || []).filter(person => person.type === 'PARTNER' || person.type === 'WORKING_PARTNER' || person.type === 'SEPARATE_COMPANY');
}

function getPersonDisplayName(person) {
  return Calculations.getPersonDisplayName(person) || '';
}

function getPersonFirstName(person) {
  return (person?.name || '').toString().trim();
}

function getPersonLastName(person) {
  return (person?.lastName || '').toString().trim();
}

function normalizePersonNamePart(value) {
  return (value || '')
    .toString()
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/(^|[\s'-])(\p{L})/gu, (match, prefix, letter) => `${prefix}${letter.toLocaleUpperCase('pl-PL')}`);
}

function getPersonCompactHeaderHtml(person) {
  const firstName = getPersonFirstName(person);
  const lastName = getPersonLastName(person);
  const title = getPersonDisplayName(person);
  if (!lastName) {
    return `<div title="${title}" style="line-height:1.1; white-space:normal; word-break:break-word; overflow-wrap:anywhere;">${firstName}</div>`;
  }

  return `<div title="${title}" style="line-height:1.05; display:flex; flex-direction:column; align-items:center; white-space:normal; word-break:break-word; overflow-wrap:anywhere;"><span>${firstName}</span><span style="font-size:0.62rem; opacity:0.9; text-transform:none;">${lastName}</span></div>`;
}

function getEmployerNameById(state, employerId) {
  if (!employerId) return '';
  return getPersonDisplayName(getEmployerCandidates(state).find(person => person.id === employerId)) || '';
}

function getDefaultCostParticipationForType(type) {
  return type === 'PARTNER' || type === 'SEPARATE_COMPANY';
}

function getDefaultProfitSharePercent() {
  return 100;
}

function getDefaultCountsEmployeeAccountingRefundForType(type) {
  return type === 'SEPARATE_COMPANY';
}

function normalizeProfitSharePercent(value, fallback = getDefaultProfitSharePercent()) {
  const parsed = parseFloat(value);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.max(0, Math.min(100, parsed));
}

function getDefaultContractTaxAmountForType(type) {
  return type === 'EMPLOYEE' || type === 'WORKING_PARTNER' ? 104 : 0;
}

function getDefaultContractZusAmountForType(type) {
  return type === 'EMPLOYEE' || type === 'WORKING_PARTNER' ? 490.27 : 0;
}

function getDefaultContractChargesPaidByEmployerForType(type) {
  if (type === 'EMPLOYEE') return true;
  if (type === 'WORKING_PARTNER') return false;
  return false;
}

function doesEmployerPayPersonContractCharges(person) {
  return !!person && person.contractChargesPaidByEmployer === true;
}

function doesPersonParticipateInCosts(person) {
  if (!person) return false;
  if (person.type === 'PARTNER') return person.participatesInCosts !== false;
  if (person.type === 'SEPARATE_COMPANY') return person.participatesInCosts !== false;
  if (person.type === 'WORKING_PARTNER') return person.participatesInCosts === true;
  return false;
}

function populatePersonEmployerSelect(selectedEmployerId = '') {
  const employerSelect = document.getElementById('person-employer');
  if (!employerSelect) return;

  const state = Store.getState();
  employerSelect.innerHTML = '<option value="">-- Wybierz wspólnika / firmę --</option>';

  getEmployerCandidates(state).forEach(person => {
    const option = document.createElement('option');
    option.value = person.id;
    option.textContent = getPersonDisplayName(person);
    if (person.id === selectedEmployerId) option.selected = true;
    employerSelect.appendChild(option);
  });
}

function updatePersonFormTypeUI(type, selectedEmployerId = '', participatesInCosts = null, contractTaxAmount = null, contractZusAmount = null, contractChargesPaidByEmployer = null, sharesEmployeeProfits = null, employeeProfitSharePercent = null, receivesCompanyEmployeeProfits = null, companyEmployeeProfitSharePercent = null, countsEmployeeAccountingRefund = null) {
  const hint = document.getElementById('person-rate-hint');
  const employerGroup = document.getElementById('person-employer-group');
  const employerSelect = document.getElementById('person-employer');
  const costShareGroup = document.getElementById('person-cost-share-group');
  const costShareCheckbox = document.getElementById('person-participates-in-costs');
  const contractConfigGroup = document.getElementById('person-contract-config-group');
  const contractTaxInput = document.getElementById('person-contract-tax');
  const contractZusInput = document.getElementById('person-contract-zus');
  const contractPaidByEmployerCheckbox = document.getElementById('person-contract-paid-by-employer');
  const separateCompanyConfigGroup = document.getElementById('person-separate-company-config-group');
  const sharesEmployeeProfitsCheckbox = document.getElementById('person-shares-employee-profits');
  const employeeProfitSharePercentGroup = document.getElementById('person-employee-profit-share-percent-group');
  const employeeProfitSharePercentInput = document.getElementById('person-employee-profit-share-percent');
  const receivesCompanyEmployeeProfitsCheckbox = document.getElementById('person-receives-company-employee-profits');
  const companyEmployeeProfitSharePercentGroup = document.getElementById('person-company-employee-profit-share-percent-group');
  const companyEmployeeProfitSharePercentInput = document.getElementById('person-company-employee-profit-share-percent');
  const countsEmployeeAccountingRefundCheckbox = document.getElementById('person-counts-employee-accounting-refund');
  const rateInput = document.getElementById('person-rate');
  const typeSelect = document.getElementById('person-type');
  const isEmployee = type === 'EMPLOYEE';
  const isSeparateCompany = type === 'SEPARATE_COMPANY';
  const canHaveEmployer = type === 'EMPLOYEE' || type === 'WORKING_PARTNER';
  const canParticipateInCosts = type === 'PARTNER' || type === 'WORKING_PARTNER' || type === 'SEPARATE_COMPANY';
  const usesContractCharges = type === 'EMPLOYEE' || type === 'WORKING_PARTNER';

  if (typeSelect) {
    typeSelect.dataset.lastType = type;
  }

  if (hint) {
    if (type === 'WORKING_PARTNER') {
      hint.textContent = '(otrzyma 100% stawki klienta lub wpisaną)';
    } else if (type === 'PARTNER' || type === 'SEPARATE_COMPANY') {
      hint.textContent = '(stawka klienta lub wpisana)';
    } else {
      hint.textContent = '';
    }
  }

  if (rateInput) {
    rateInput.placeholder = isEmployee ? 'Np. 30' : 'Stawka Klienta';
  }

  if (employerGroup && employerSelect) {
    if (canHaveEmployer) {
      populatePersonEmployerSelect(selectedEmployerId);
      employerGroup.style.display = 'block';
    } else {
      employerGroup.style.display = 'none';
      employerSelect.value = '';
    }
  }

  if (costShareGroup && costShareCheckbox) {
    if (canParticipateInCosts) {
      costShareGroup.style.display = 'block';
      costShareCheckbox.checked = participatesInCosts === null
        ? getDefaultCostParticipationForType(type)
        : participatesInCosts === true;
    } else {
      costShareGroup.style.display = 'none';
      costShareCheckbox.checked = false;
    }
  }

  if (contractConfigGroup && contractTaxInput && contractZusInput && contractPaidByEmployerCheckbox) {
    if (usesContractCharges) {
      contractConfigGroup.style.display = 'block';
      contractTaxInput.value = contractTaxAmount === null ? getDefaultContractTaxAmountForType(type) : contractTaxAmount;
      contractZusInput.value = contractZusAmount === null ? getDefaultContractZusAmountForType(type) : contractZusAmount;
      contractPaidByEmployerCheckbox.checked = contractChargesPaidByEmployer === null
        ? getDefaultContractChargesPaidByEmployerForType(type)
        : contractChargesPaidByEmployer === true;
    } else {
      contractConfigGroup.style.display = 'none';
      contractTaxInput.value = '0';
      contractZusInput.value = '0';
      contractPaidByEmployerCheckbox.checked = false;
    }
  }

  if (separateCompanyConfigGroup && sharesEmployeeProfitsCheckbox && employeeProfitSharePercentInput && receivesCompanyEmployeeProfitsCheckbox && companyEmployeeProfitSharePercentInput && countsEmployeeAccountingRefundCheckbox) {
    if (isSeparateCompany) {
      separateCompanyConfigGroup.style.display = 'block';
      sharesEmployeeProfitsCheckbox.checked = sharesEmployeeProfits === true;
      employeeProfitSharePercentInput.value = normalizeProfitSharePercent(employeeProfitSharePercent, getDefaultProfitSharePercent());
      receivesCompanyEmployeeProfitsCheckbox.checked = receivesCompanyEmployeeProfits === true;
      companyEmployeeProfitSharePercentInput.value = normalizeProfitSharePercent(companyEmployeeProfitSharePercent, getDefaultProfitSharePercent());
      countsEmployeeAccountingRefundCheckbox.checked = countsEmployeeAccountingRefund === null
        ? getDefaultCountsEmployeeAccountingRefundForType(type)
        : countsEmployeeAccountingRefund === true;
      if (employeeProfitSharePercentGroup) employeeProfitSharePercentGroup.style.display = sharesEmployeeProfitsCheckbox.checked ? 'block' : 'none';
      if (companyEmployeeProfitSharePercentGroup) companyEmployeeProfitSharePercentGroup.style.display = receivesCompanyEmployeeProfitsCheckbox.checked ? 'block' : 'none';
    } else {
      separateCompanyConfigGroup.style.display = 'none';
      sharesEmployeeProfitsCheckbox.checked = false;
      employeeProfitSharePercentInput.value = getDefaultProfitSharePercent();
      receivesCompanyEmployeeProfitsCheckbox.checked = false;
      companyEmployeeProfitSharePercentInput.value = getDefaultProfitSharePercent();
      countsEmployeeAccountingRefundCheckbox.checked = false;
      if (employeeProfitSharePercentGroup) employeeProfitSharePercentGroup.style.display = 'none';
      if (companyEmployeeProfitSharePercentGroup) companyEmployeeProfitSharePercentGroup.style.display = 'none';
    }
  }
}

// ==========================================
// SUMMARY VIEW
// ==========================================
function getSortedMonthlySheetsForSummary(state) {
  const selectedMonth = getSelectedMonthKey();
  const clientOrder = new Map((state.clients || []).map((client, index) => [client.name, index]));

  return (state.monthlySheets || [])
    .filter(sheet => sheet.month === selectedMonth)
    .sort((a, b) => {
      const clientDiff = (clientOrder.get(a.client) ?? Number.MAX_SAFE_INTEGER) - (clientOrder.get(b.client) ?? Number.MAX_SAFE_INTEGER);
      if (clientDiff !== 0) return clientDiff;
      return (a.site || '').localeCompare(b.site || '', 'pl-PL') || a.id.localeCompare(b.id, 'pl-PL');
    });
}

function getSortedWorksSheetsForSummary(state) {
  const selectedMonth = getSelectedMonthKey();
  const clientOrder = new Map((state.clients || []).map((client, index) => [client.name, index]));

  return (state.worksSheets || [])
    .filter(sheet => sheet.month === selectedMonth)
    .sort((a, b) => {
      const clientDiff = (clientOrder.get(a.client) ?? Number.MAX_SAFE_INTEGER) - (clientOrder.get(b.client) ?? Number.MAX_SAFE_INTEGER);
      if (clientDiff !== 0) return clientDiff;
      return (a.site || '').localeCompare(b.site || '', 'pl-PL') || a.id.localeCompare(b.id, 'pl-PL');
    });
}

function getSummaryShortcutItems(state) {
  const monthlySheets = getSortedMonthlySheetsForSummary(state);
  const worksSheets = getSortedWorksSheetsForSummary(state);

  const items = [];

  items.push({
      type: 'expense-form',
      targetId: '',
      icon: 'plus',
      title: 'Nowy Koszt / Zaliczka',
      meta: 'Koszty i Zaliczki',
      description: 'Otwórz panel „Nowy Dowód/Koszt”.'
    });

  if (monthlySheets.length > 0) {
    items.push(...monthlySheets.map(sheet => ({
          type: 'monthly-sheet',
          targetId: sheet.id,
          icon: 'clock',
          title: 'Tabela z Godzinami',
          meta: sheet.site ? `${sheet.client} • ${sheet.site}` : (sheet.client || 'Bez klienta'),
          description: 'Przejdź do konkretnego arkusza godzin.'
        })));
  } else {
    items.push({
          type: 'monthly-sheet-create',
          targetId: '',
          icon: 'plus',
          title: 'Nowy Arkusz Godzin',
          meta: 'Godziny Pracy',
          description: 'Utwórz pierwszy arkusz godzin.'
        });
  }

  if (worksSheets.length > 0) {
    items.push(...worksSheets.map(sheet => ({
          type: 'works-sheet',
          targetId: sheet.id,
          icon: 'settings',
          title: 'Wykonane Prace',
          meta: sheet.site ? `${sheet.client} • ${sheet.site}` : (sheet.client || 'Bez klienta'),
          description: 'Przejdź do konkretnego arkusza prac.'
        })));
  } else {
    items.push({
          type: 'works-sheet-create',
          targetId: '',
          icon: 'plus',
          title: 'Nowy Arkusz Wykonanych Pr.',
          meta: 'Wykonane Prace',
          description: 'Utwórz pierwszy arkusz prac.'
        });
  }

  return items;
}

function scrollCurrentMonthlySheetToTodayIfNeeded(sheetId = window.currentSheetId) {
  const sheet = Store.getMonthlySheet(sheetId);
  if (!sheet) return;

  const today = new Date();
  const todayMonthKey = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;
  if (sheet.month !== todayMonthKey) return;

  const scrollToToday = () => {
    const todayRow = document.querySelector(`#monthly-hours-body .fill-day-cb[data-day="${today.getDate()}"]`)?.closest('tr');
    if (todayRow) {
      todayRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  };

  requestAnimationFrame(scrollToToday);
  setTimeout(scrollToToday, 250);
}

function openSummaryShortcutTarget(type, targetId = '') {
  if (type === 'expense-form') {
    window.currentSheetId = null;
    window.currentWorksSheetId = null;
    activateNavigationTarget('expenses-view');
    openExpenseFormForCreate();
    return;
  }

  if (type === 'monthly-sheet-create') {
    window.currentSheetId = null;
    window.currentWorksSheetId = null;
    activateNavigationTarget('hours-view');
    document.getElementById('btn-create-sheet')?.click();
    return;
  }

  if (type === 'monthly-sheet' && targetId) {
    window.currentSheetId = targetId;
    activateNavigationTarget('hours-view');

    const tableContainer = document.querySelector('#sheet-detail-container .table-container');
    if (tableContainer) {
      tableContainer.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    scrollCurrentMonthlySheetToTodayIfNeeded(targetId);
    return;
  }

  if (type === 'works-sheet-create') {
    window.currentSheetId = null;
    window.currentWorksSheetId = null;
    activateNavigationTarget('works-view');
    document.getElementById('btn-create-works-sheet')?.click();
    return;
  }

  if (type === 'works-sheet' && targetId) {
    window.currentWorksSheetId = targetId;
    activateNavigationTarget('works-view');

    const tableContainer = document.querySelector('#works-detail-container .table-container');
    if (tableContainer) {
      tableContainer.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }
}

function renderSummaryShortcuts(state) {
  const container = document.getElementById('summary-shortcuts-grid');
  if (!container) return;

  const items = getSummaryShortcutItems(state);
  container.innerHTML = items.map(item => `
    <button type="button" class="glass-panel stat-card summary-shortcut-card" data-shortcut-type="${item.type}" data-shortcut-target-id="${item.targetId}" ${item.isWarning ? 'style="border-color: var(--danger); background: rgba(239, 68, 68, 0.05);"' : ''}>
      <span class="summary-shortcut-kicker" ${item.isWarning ? 'style="color: var(--danger); font-weight: 700;"' : ''}>${item.isWarning ? 'Wymagana akcja' : 'Szybki skrót'}</span>
      <span class="summary-shortcut-title" ${item.isWarning ? 'style="color: var(--danger);"' : ''}>${item.title}</span>
      <span class="summary-shortcut-meta">${item.meta}</span>
      <span class="summary-shortcut-description" ${item.isWarning ? 'style="color: var(--danger); font-weight: 500;"' : ''}>${item.description}</span>
      <i data-lucide="${item.icon}" class="stat-icon" style="width:24px;height:24px;${item.isWarning ? 'color: var(--danger);' : ''}"></i>
    </button>
  `).join('');

  container.querySelectorAll('.summary-shortcut-card').forEach(button => {
    button.addEventListener('click', () => {
      openSummaryShortcutTarget(
        button.getAttribute('data-shortcut-type'),
        button.getAttribute('data-shortcut-target-id') || ''
      );
    });
  });
}

function renderSummary() {
  const state = Store.getState();
  const settlement = Calculations.generateSettlement(state);
  const activePersons = Calculations.getActivePersons(state);
  const summaryRowsById = new Map();

  renderSummaryShortcuts(state);

  settlement.partners.forEach(p => {
    summaryRowsById.set(p.person.id, {
      name: getPersonDisplayName(p.person),
      badgeClass: 'badge-partner',
      label: 'Wspólnik',
      hours: p.hours,
      amount: p.toPayout
    });
  });

  (settlement.separateCompanies || []).forEach(company => {
    summaryRowsById.set(company.person.id, {
      name: getPersonDisplayName(company.person),
      badgeClass: 'badge-partner',
      label: 'Osobna Firma',
      hours: company.hours,
      amount: company.toPayout
    });
  });

  settlement.workingPartners.forEach(wp => {
    summaryRowsById.set(wp.person.id, {
      name: getPersonDisplayName(wp.person),
      badgeClass: 'badge-working-partner',
      label: 'Wspólnik Pracujący',
      hours: wp.hours,
      amount: wp.toPayout
    });
  });

  settlement.employees.forEach(e => {
    summaryRowsById.set(e.person.id, {
      name: getPersonDisplayName(e.person),
      badgeClass: 'badge-employee',
      label: 'Pracownik',
      hours: e.hours,
      amount: e.toPayout
    });
  });

  document.getElementById('dash-people-count').textContent = activePersons.length;
  document.getElementById('dash-total-hours').textContent = `${settlement.totalTeamHours.toFixed(1)}h`;
  document.getElementById('dash-total-revenue').textContent = `${settlement.commonRevenue.toFixed(2)} zł`;
  document.getElementById('dash-profit').textContent = `${settlement.profitToSplit.toFixed(2)} zł`;

  const tbody = document.getElementById('dash-quick-preview');
  tbody.innerHTML = '';

  if (activePersons.length === 0) {
    tbody.innerHTML = `<tr><td colspan="4" style="text-align:center;color:var(--text-muted)">Brak dodanych osób. Dodaj osoby w zakładce "Osoby".</td></tr>`;
    return;
  }

  activePersons.forEach(person => {
    const row = summaryRowsById.get(person.id);
    if (!row) return;

    tbody.innerHTML += `
      <tr>
        <td>${row.name}</td>
        <td><span class="badge ${row.badgeClass}">${row.label}</span></td>
        <td>${row.hours.toFixed(1)}h</td>
        <td style="color: var(--success);">${row.amount.toFixed(2)} zł</td>
      </tr>
    `;
  });
}

// ==========================================
// CLIENTS VIEW
// ==========================================
function getClientRegistryFieldValue(id) {
  const input = document.getElementById(id);
  return input ? input.value.trim() : '';
}

function setClientRegistryStatus(message = '', type = '') {
  const statusEl = document.getElementById('client-registry-status');
  if (!statusEl) return;

  statusEl.textContent = message;
  statusEl.classList.remove('is-success', 'is-error');
  if (type === 'success') statusEl.classList.add('is-success');
  if (type === 'error') statusEl.classList.add('is-error');
}

function fillClientRegistryFields(client = {}) {
  const form = document.getElementById('client-form');
  const mapping = {
    'client-full-company-name': client.fullCompanyName || '',
    'client-address': client.address || '',
    'client-nip': client.nip || '',
    'client-krs': client.krs || '',
    'client-regon': client.regon || ''
  };

  Object.entries(mapping).forEach(([id, value]) => {
    const input = document.getElementById(id);
    if (input) input.value = value;
  });

  if (form) {
    form.dataset.accountNumbers = JSON.stringify(Array.isArray(client.accountNumbers) ? client.accountNumbers : []);
  }
}

function normalizeClientRegistryDigits(value) {
  return (value || '').toString().replace(/\D+/g, '');
}

function getTodayIsoDate() {
  return new Date().toISOString().split('T')[0];
}

function isValidPolishNip(nip) {
  return /^\d{10}$/.test(normalizeClientRegistryDigits(nip));
}

function normalizeWhitelistSubject(subject = {}, requestedNip = '') {
  const address = (subject.workingAddress || subject.residenceAddress || '').toString().trim();
  return {
    fullCompanyName: (subject.name || '').toString().trim(),
    address,
    nip: normalizeClientRegistryDigits(subject.nip || requestedNip),
    krs: '',
    regon: normalizeClientRegistryDigits(subject.regon || ''),
    accountNumbers: Array.isArray(subject.accountNumbers) ? subject.accountNumbers.filter(Boolean) : [],
    fetchedFromRegistryAt: new Date().toISOString(),
    registrySourceUrl: `https://wl-api.mf.gov.pl/api/search/nip/${normalizeClientRegistryDigits(subject.nip || requestedNip)}?date=${getTodayIsoDate()}`
  };
}

async function fetchClientRegistryDataByNip(nip) {
  const normalizedNip = normalizeClientRegistryDigits(nip);
  if (!isValidPolishNip(normalizedNip)) {
    throw new Error('NIP musi mieć dokładnie 10 cyfr.');
  }

  const today = getTodayIsoDate();
  const url = `https://wl-api.mf.gov.pl/api/search/nip/${normalizedNip}?date=${today}`;

  let response;
  try {
    response = await fetch(url);
  } catch {
    throw new Error('Nie udało się połączyć z API Białej Listy Ministerstwa Finansów.');
  }

  if (response.status === 404) {
    throw new Error('Nie znaleziono firmy o podanym NIP w Wykazie Podatników VAT.');
  }

  if (response.status === 429) {
    throw new Error('Przekroczono limit zapytań do API Białej Listy MF. Spróbuj ponownie za chwilę.');
  }

  if (!response.ok) {
    throw new Error(`API Białej Listy MF zwróciło błąd (${response.status}).`);
  }

  let payload;
  try {
    payload = await response.json();
  } catch {
    throw new Error('API Białej Listy MF zwróciło nieprawidłową odpowiedź JSON.');
  }

  const subject = payload?.result?.subject || (Array.isArray(payload?.result?.subjects) ? payload.result.subjects[0] : null);
  if (!subject) {
    throw new Error('Brak danych firmy w odpowiedzi API Białej Listy MF.');
  }

  const parsed = normalizeWhitelistSubject(subject, normalizedNip);
  if (!parsed.fullCompanyName && !parsed.address && !parsed.regon) {
    throw new Error('API Białej Listy MF nie zwróciło wystarczających danych firmy.');
  }

  return parsed;
}

function initClientsTracker() {
  const form = document.getElementById('client-form');
  const btnAdd = document.getElementById('btn-add-client');
  const btnCancel = document.getElementById('btn-cancel-client');
  const btnFetchRegistry = document.getElementById('btn-fetch-client-registry');

  btnAdd.addEventListener('click', () => {
    document.getElementById('client-form-title').textContent = 'Dodaj Klienta';
    document.getElementById('client-id').value = '';
    document.getElementById('client-name').value = '';
    document.getElementById('client-rate').value = '';
    fillClientRegistryFields();
    setClientRegistryStatus('');

    const grid = document.getElementById('client-works-prices-grid');
    grid.innerHTML = '';
    Store.getState().worksCatalog.forEach(w => {
      grid.innerHTML += `
        <div class="form-group">
          <label>${w.name} (${w.unit})</label>
          <input type="number" step="0.01" data-work-id="${w.id}" class="client-work-price-input" placeholder="Domyślnie ${parseFloat(w.defaultPrice).toFixed(2)}">
        </div>
      `;
    });

    document.querySelector('#clients-view .table-container').style.display = 'none';
    form.style.display = 'block';
    form.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });

  btnCancel.addEventListener('click', () => {
    form.style.display = 'none';
    document.querySelector('#clients-view .table-container').style.display = 'block';
    setClientRegistryStatus('');
  });

  if (btnFetchRegistry) {
    btnFetchRegistry.addEventListener('click', async () => {
      const nip = getClientRegistryFieldValue('client-nip');
      if (!nip) {
        setClientRegistryStatus('Najpierw wpisz NIP klienta.', 'error');
        return;
      }

      const originalText = btnFetchRegistry.textContent;
      btnFetchRegistry.disabled = true;
      btnFetchRegistry.textContent = 'Pobieranie...';
      setClientRegistryStatus('Trwa pobieranie danych firmy z Białej Listy MF...', '');

      try {
        const registryData = await fetchClientRegistryDataByNip(nip);
        fillClientRegistryFields(registryData);
        setClientRegistryStatus('Pobrano dane firmy z API Białej Listy MF.', 'success');
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Nie udało się pobrać danych firmy.';
        setClientRegistryStatus(message, 'error');
        alert(message);
      } finally {
        btnFetchRegistry.disabled = false;
        btnFetchRegistry.textContent = originalText;
      }
    });
  }

  form.addEventListener('submit', (e) => {
    e.preventDefault();
    const id = document.getElementById('client-id').value;
    const name = document.getElementById('client-name').value.trim();
    const rate = parseFloat(document.getElementById('client-rate').value);
    const fullCompanyName = getClientRegistryFieldValue('client-full-company-name');
    const address = getClientRegistryFieldValue('client-address');
    const nip = normalizeClientRegistryDigits(getClientRegistryFieldValue('client-nip'));
    const krs = normalizeClientRegistryDigits(getClientRegistryFieldValue('client-krs'));
    const regon = normalizeClientRegistryDigits(getClientRegistryFieldValue('client-regon'));
    const accountNumbers = (() => {
      try {
        return JSON.parse(form.dataset.accountNumbers || '[]');
      } catch {
        return [];
      }
    })();

    if (!name) {
      alert('Podaj nazwę klienta.');
      return;
    }

    if (!Number.isFinite(rate) || rate <= 0) {
      alert('Klient musi mieć stawkę większą od zera.');
      document.getElementById('client-rate').focus();
      return;
    }

    const customWorkPrices = {};
    document.querySelectorAll('.client-work-price-input').forEach(input => {
      if (input.value !== '') {
        customWorkPrices[input.getAttribute('data-work-id')] = parseFloat(input.value);
      }
    });

    if (id) {
      Store.updateClient(id, { name, hourlyRate: rate, customWorkPrices, fullCompanyName, address, nip, krs, regon, accountNumbers });
    } else {
      Store.addClient({ name, hourlyRate: rate, customWorkPrices, fullCompanyName, address, nip, krs, regon, accountNumbers });
    }
    form.style.display = 'none';
    document.querySelector('#clients-view .table-container').style.display = 'block';
    setClientRegistryStatus('');
  });

  new Sortable(document.getElementById('clients-table-body'), {
    handle: '.drag-handle',
    animation: 150,
    ghostClass: 'sortable-ghost',
    onEnd: () => {
      const ids = Array.from(document.querySelectorAll('#clients-table-body tr')).map(tr => tr.dataset.id);
      Store.reorderClients(ids);
    }
  });
}

function renderClients() {
  const state = Store.getState();
  const tbody = document.getElementById('clients-table-body');

  document.getElementById('client-form').style.display = 'none';
  const tableContainer = document.querySelector('#clients-view .table-container');
  if (tableContainer) tableContainer.style.display = 'block';

  tbody.innerHTML = '';

  if (state.clients.length === 0) {
    tbody.innerHTML = `<tr><td colspan="4" style="text-align:center;color:var(--text-muted)">Brak klientów. Dodaj pierwszego klienta.</td></tr>`;
    return;
  }

  state.clients.forEach(c => {
    const clientRate = parseFloat(c.hourlyRate);
    const isClientActive = Store.isClientActiveInMonth(c.id);
    const tr = document.createElement('tr');
    tr.dataset.id = c.id;
    if (!isClientActive) {
      tr.style.opacity = '0.5';
    }
    tr.innerHTML = `
      <td><div class="drag-handle"><i data-lucide="grip-vertical" style="width:16px;height:16px;color:var(--text-muted)"></i></div></td>
      <td style="font-weight: 500">${c.name}</td>
      <td>${Number.isFinite(clientRate) && clientRate > 0 ? `${clientRate.toFixed(2)} zł/h` : '<span style="color: var(--danger);">Brak stawki</span>'}</td>
      <td>
        <button class="btn-status btn-status-client ${!isClientActive ? 'inactive' : 'active'}" data-id="${c.id}">
          ${!isClientActive ? 'Nieaktywny' : 'Aktywny'}
        </button>
      </td>
      <td>
        <div style="display: flex; gap: 0.5rem; flex-wrap: wrap;">
          <button class="btn btn-secondary btn-icon drag-handle btn-mobile-only" title="Przeciągnij">
            <i data-lucide="grip-horizontal" style="width:16px;height:16px"></i>
          </button>
          <button class="btn btn-secondary btn-icon btn-edit-client" data-id="${c.id}">
            <i data-lucide="edit-2" style="width:16px;height:16px"></i>
          </button>
          <button class="btn btn-danger btn-icon btn-delete-client" data-id="${c.id}">
            <i data-lucide="trash-2" style="width:16px;height:16px"></i>
          </button>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  });

  lucide.createIcons();

  document.querySelectorAll('.btn-edit-client').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = e.currentTarget.getAttribute('data-id');
      const client = state.clients.find(x => x.id === id);
      if (client) {
        document.getElementById('client-form-title').textContent = 'Edytuj Klienta';
        document.getElementById('client-id').value = client.id;
        document.getElementById('client-name').value = client.name;
        document.getElementById('client-rate').value = client.hourlyRate;
        fillClientRegistryFields(client);
        setClientRegistryStatus('');

        const grid = document.getElementById('client-works-prices-grid');
        grid.innerHTML = '';
        Store.getState().worksCatalog.forEach(w => {
          const priceVal = (client.customWorkPrices && client.customWorkPrices[w.id]) ? client.customWorkPrices[w.id] : '';
          grid.innerHTML += `
            <div class="form-group">
              <label>${w.name} (${w.unit})</label>
              <input type="number" step="0.01" data-work-id="${w.id}" class="client-work-price-input" placeholder="Domyślnie ${parseFloat(w.defaultPrice).toFixed(2)}" value="${priceVal}">
            </div>
          `;
        });

        document.querySelector('#clients-view .table-container').style.display = 'none';
        document.getElementById('client-form').style.display = 'block';
        document.getElementById('client-form').scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    });
  });

  document.querySelectorAll('.btn-status-client').forEach(btn => {
    btn.addEventListener('click', (e) => {
      Store.toggleClientStatus(e.currentTarget.getAttribute('data-id'));
    });
  });

  document.querySelectorAll('.btn-delete-client').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = e.currentTarget.getAttribute('data-id');
      if (confirm(`Czy na pewno chcesz usunąć klienta "${state.clients.find(c => c.id === id).name}"?`)) {
        Store.deleteClient(id);
      }
    });
  });
}

// ==========================================
// PEOPLE VIEW
// ==========================================
function initPeopleManager() {
  const form = document.getElementById('person-form');
  const btnAdd = document.getElementById('btn-add-person');
  const btnCancel = document.getElementById('btn-cancel-person');
  const nameInput = document.getElementById('person-name');
  const lastNameInput = document.getElementById('person-last-name');
  const emailInput = document.getElementById('person-email');
  const personTypeSelect = document.getElementById('person-type');
  const employerSelect = document.getElementById('person-employer');
  const costShareCheckbox = document.getElementById('person-participates-in-costs');
  const contractTaxInput = document.getElementById('person-contract-tax');
  const contractZusInput = document.getElementById('person-contract-zus');
  const contractPaidByEmployerCheckbox = document.getElementById('person-contract-paid-by-employer');
  const sharesEmployeeProfitsCheckbox = document.getElementById('person-shares-employee-profits');
  const employeeProfitSharePercentInput = document.getElementById('person-employee-profit-share-percent');
  const receivesCompanyEmployeeProfitsCheckbox = document.getElementById('person-receives-company-employee-profits');
  const companyEmployeeProfitSharePercentInput = document.getElementById('person-company-employee-profit-share-percent');
  const countsEmployeeAccountingRefundCheckbox = document.getElementById('person-counts-employee-accounting-refund');

  const normalizePersonIdentityInput = (input) => {
    if (!input) return;
    input.value = normalizePersonNamePart(input.value);
  };

  if (nameInput) {
    nameInput.addEventListener('blur', () => normalizePersonIdentityInput(nameInput));
  }

  if (lastNameInput) {
    lastNameInput.addEventListener('blur', () => normalizePersonIdentityInput(lastNameInput));
  }

  btnAdd.addEventListener('click', () => {
    document.getElementById('person-id').value = '';
    document.getElementById('person-name').value = '';
    if (lastNameInput) lastNameInput.value = '';
    if (emailInput) emailInput.value = '';
    document.getElementById('person-type').value = 'EMPLOYEE';
    document.getElementById('person-rate').value = '';
    if (employerSelect) employerSelect.value = '';
    if (costShareCheckbox) costShareCheckbox.checked = false;
    if (contractTaxInput) contractTaxInput.value = getDefaultContractTaxAmountForType('EMPLOYEE');
    if (contractZusInput) contractZusInput.value = getDefaultContractZusAmountForType('EMPLOYEE');
    if (contractPaidByEmployerCheckbox) contractPaidByEmployerCheckbox.checked = getDefaultContractChargesPaidByEmployerForType('EMPLOYEE');
    if (sharesEmployeeProfitsCheckbox) sharesEmployeeProfitsCheckbox.checked = false;
    if (employeeProfitSharePercentInput) employeeProfitSharePercentInput.value = getDefaultProfitSharePercent();
    if (receivesCompanyEmployeeProfitsCheckbox) receivesCompanyEmployeeProfitsCheckbox.checked = false;
    if (companyEmployeeProfitSharePercentInput) companyEmployeeProfitSharePercentInput.value = getDefaultProfitSharePercent();
    if (countsEmployeeAccountingRefundCheckbox) countsEmployeeAccountingRefundCheckbox.checked = getDefaultCountsEmployeeAccountingRefundForType('SEPARATE_COMPANY');
    document.getElementById('person-form-title').textContent = 'Dodaj Osobę';
    updatePersonFormTypeUI('EMPLOYEE', '', false, getDefaultContractTaxAmountForType('EMPLOYEE'), getDefaultContractZusAmountForType('EMPLOYEE'), getDefaultContractChargesPaidByEmployerForType('EMPLOYEE'), false, getDefaultProfitSharePercent(), false, getDefaultProfitSharePercent(), getDefaultCountsEmployeeAccountingRefundForType('EMPLOYEE'));
    document.querySelector('#people-view .table-container').style.display = 'none';
    form.style.display = 'block';
    form.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });

  btnCancel.addEventListener('click', () => {
    form.style.display = 'none';
    document.querySelector('#people-view .table-container').style.display = 'block';
  });

  personTypeSelect.addEventListener('change', (e) => {
    const previousType = personTypeSelect.dataset.lastType || '';
    const keepCurrentCostSetting = previousType === 'PARTNER' || previousType === 'WORKING_PARTNER' || previousType === 'SEPARATE_COMPANY';
    const keepCurrentContractSettings = previousType === 'EMPLOYEE' || previousType === 'WORKING_PARTNER';
    const keepSeparateCompanySettings = previousType === 'SEPARATE_COMPANY';
    updatePersonFormTypeUI(
      e.target.value,
      employerSelect ? employerSelect.value : '',
      keepCurrentCostSetting && costShareCheckbox ? costShareCheckbox.checked : null,
      keepCurrentContractSettings && contractTaxInput ? contractTaxInput.value : null,
      keepCurrentContractSettings && contractZusInput ? contractZusInput.value : null,
      keepCurrentContractSettings && contractPaidByEmployerCheckbox ? contractPaidByEmployerCheckbox.checked : null,
      keepSeparateCompanySettings && sharesEmployeeProfitsCheckbox ? sharesEmployeeProfitsCheckbox.checked : null,
      keepSeparateCompanySettings && employeeProfitSharePercentInput ? employeeProfitSharePercentInput.value : null,
      keepSeparateCompanySettings && receivesCompanyEmployeeProfitsCheckbox ? receivesCompanyEmployeeProfitsCheckbox.checked : null,
      keepSeparateCompanySettings && companyEmployeeProfitSharePercentInput ? companyEmployeeProfitSharePercentInput.value : null,
      keepSeparateCompanySettings && countsEmployeeAccountingRefundCheckbox ? countsEmployeeAccountingRefundCheckbox.checked : null
    );
  });

  if (sharesEmployeeProfitsCheckbox) {
    sharesEmployeeProfitsCheckbox.addEventListener('change', () => {
      updatePersonFormTypeUI(
        personTypeSelect.value,
        employerSelect ? employerSelect.value : '',
        costShareCheckbox ? costShareCheckbox.checked : null,
        contractTaxInput ? contractTaxInput.value : null,
        contractZusInput ? contractZusInput.value : null,
        contractPaidByEmployerCheckbox ? contractPaidByEmployerCheckbox.checked : null,
        sharesEmployeeProfitsCheckbox.checked,
        employeeProfitSharePercentInput ? employeeProfitSharePercentInput.value : null,
        receivesCompanyEmployeeProfitsCheckbox ? receivesCompanyEmployeeProfitsCheckbox.checked : null,
        companyEmployeeProfitSharePercentInput ? companyEmployeeProfitSharePercentInput.value : null,
        countsEmployeeAccountingRefundCheckbox ? countsEmployeeAccountingRefundCheckbox.checked : null
      );
    });
  }

  if (receivesCompanyEmployeeProfitsCheckbox) {
    receivesCompanyEmployeeProfitsCheckbox.addEventListener('change', () => {
      updatePersonFormTypeUI(
        personTypeSelect.value,
        employerSelect ? employerSelect.value : '',
        costShareCheckbox ? costShareCheckbox.checked : null,
        contractTaxInput ? contractTaxInput.value : null,
        contractZusInput ? contractZusInput.value : null,
        contractPaidByEmployerCheckbox ? contractPaidByEmployerCheckbox.checked : null,
        sharesEmployeeProfitsCheckbox ? sharesEmployeeProfitsCheckbox.checked : null,
        employeeProfitSharePercentInput ? employeeProfitSharePercentInput.value : null,
        receivesCompanyEmployeeProfitsCheckbox.checked,
        companyEmployeeProfitSharePercentInput ? companyEmployeeProfitSharePercentInput.value : null,
        countsEmployeeAccountingRefundCheckbox ? countsEmployeeAccountingRefundCheckbox.checked : null
      );
    });
  }

  form.addEventListener('submit', (e) => {
    e.preventDefault();
    const id = document.getElementById('person-id').value;
    const name = normalizePersonNamePart(document.getElementById('person-name').value);
    const lastName = lastNameInput ? normalizePersonNamePart(lastNameInput.value) : '';
    const email = emailInput ? emailInput.value.trim() : '';
    if (nameInput) nameInput.value = name;
    if (lastNameInput) lastNameInput.value = lastName;
    const type = document.getElementById('person-type').value;
    const rateVal = document.getElementById('person-rate').value;
    const hourlyRate = parseFloat(rateVal) || 0;
    const employerId = (type === 'EMPLOYEE' || type === 'WORKING_PARTNER') && employerSelect ? (employerSelect.value || null) : null;
    const participatesInCosts = (type === 'PARTNER' || type === 'WORKING_PARTNER') && costShareCheckbox
      ? costShareCheckbox.checked
      : false;
    const contractTaxAmount = (type === 'EMPLOYEE' || type === 'WORKING_PARTNER') && contractTaxInput
      ? (parseFloat(contractTaxInput.value) || 0)
      : 0;
    const contractZusAmount = (type === 'EMPLOYEE' || type === 'WORKING_PARTNER') && contractZusInput
      ? (parseFloat(contractZusInput.value) || 0)
      : 0;
    const contractChargesPaidByEmployer = (type === 'EMPLOYEE' || type === 'WORKING_PARTNER') && contractPaidByEmployerCheckbox
      ? contractPaidByEmployerCheckbox.checked
      : false;
    const sharesEmployeeProfits = type === 'SEPARATE_COMPANY' && sharesEmployeeProfitsCheckbox
      ? sharesEmployeeProfitsCheckbox.checked
      : false;
    const employeeProfitSharePercent = type === 'SEPARATE_COMPANY' && employeeProfitSharePercentInput
      ? normalizeProfitSharePercent(employeeProfitSharePercentInput.value)
      : getDefaultProfitSharePercent();
    const receivesCompanyEmployeeProfits = type === 'SEPARATE_COMPANY' && receivesCompanyEmployeeProfitsCheckbox
      ? receivesCompanyEmployeeProfitsCheckbox.checked
      : false;
    const companyEmployeeProfitSharePercent = type === 'SEPARATE_COMPANY' && companyEmployeeProfitSharePercentInput
      ? normalizeProfitSharePercent(companyEmployeeProfitSharePercentInput.value)
      : getDefaultProfitSharePercent();
    const countsEmployeeAccountingRefund = type === 'SEPARATE_COMPANY' && countsEmployeeAccountingRefundCheckbox
      ? countsEmployeeAccountingRefundCheckbox.checked
      : false;

    if (id) {
      Store.updatePerson(id, { name, lastName, email, type, hourlyRate, employerId, participatesInCosts, contractTaxAmount, contractZusAmount, contractChargesPaidByEmployer, sharesEmployeeProfits, employeeProfitSharePercent, receivesCompanyEmployeeProfits, companyEmployeeProfitSharePercent, countsEmployeeAccountingRefund });
    } else {
      Store.addPerson({ name, lastName, email, type, hourlyRate, employerId, participatesInCosts, contractTaxAmount, contractZusAmount, contractChargesPaidByEmployer, sharesEmployeeProfits, employeeProfitSharePercent, receivesCompanyEmployeeProfits, companyEmployeeProfitSharePercent, countsEmployeeAccountingRefund });
    }
    form.style.display = 'none';
    document.querySelector('#people-view .table-container').style.display = 'block';
  });

  new Sortable(document.getElementById('people-table-body'), {
    handle: '.drag-handle',
    animation: 150,
    ghostClass: 'sortable-ghost',
    onEnd: () => {
      const ids = Array.from(document.querySelectorAll('#people-table-body tr')).map(tr => tr.dataset.id);
      Store.reorderPersons(ids);
    }
  });
}

function renderPeople() {
  const state = Store.getState();
  const tbody = document.getElementById('people-table-body');

  const form = document.getElementById('person-form');
  if (form) form.style.display = 'none';
  const tableContainer = document.querySelector('#people-view .table-container');
  if (tableContainer) tableContainer.style.display = 'block';

  tbody.innerHTML = '';

  if (state.persons.length === 0) {
    tbody.innerHTML = `<tr><td colspan="8" style="text-align:center;color:var(--text-muted)">Brak rekordów</td></tr>`;
    return;
  }

  state.persons.forEach(p => {
    const isPersonActive = Store.isPersonActiveInMonth(p.id);
    const tr = document.createElement('tr');
    tr.dataset.id = p.id;
    if (!isPersonActive) {
      tr.style.opacity = '0.5';
    }
    tr.innerHTML = `
      <td><div class="drag-handle"><i data-lucide="grip-vertical" style="width:16px;height:16px;color:var(--text-muted)"></i></div></td>
      <td style="font-weight: 500;">${getPersonDisplayName(p)}</td>
      <td>
        <span class="badge ${p.type === 'PARTNER' || p.type === 'SEPARATE_COMPANY' ? 'badge-partner' : (p.type === 'WORKING_PARTNER' ? 'badge-working-partner' : 'badge-employee')}">
          ${p.type === 'PARTNER' ? 'Wspólnik' : (p.type === 'SEPARATE_COMPANY' ? 'Osobna Firma' : (p.type === 'WORKING_PARTNER' ? 'Wspólnik Pracujący' : 'Pracownik'))}
        </span>
      </td>
      <td>${p.type === 'EMPLOYEE' ? (p.hourlyRate ? p.hourlyRate.toFixed(2) + ' zł' : '-') : (p.hourlyRate ? p.hourlyRate.toFixed(2) + ' zł' : 'Stawka Klienta')}</td>
      <td>${(p.type === 'EMPLOYEE' || p.type === 'WORKING_PARTNER') ? (getEmployerNameById(state, p.employerId) || '<span style="color: var(--text-muted)">Nie przypisano</span>') : '-'}</td>
      <td>${(p.type === 'PARTNER' || p.type === 'WORKING_PARTNER' || p.type === 'SEPARATE_COMPANY') ? `<span class="badge ${doesPersonParticipateInCosts(p) ? 'badge-working-partner' : 'badge-employee'}">${doesPersonParticipateInCosts(p) ? 'Tak' : 'Nie'}</span>` : '-'}</td>
      <td>
        <button class="btn-status btn-status-person ${!isPersonActive ? 'inactive' : 'active'}" data-id="${p.id}">
          ${!isPersonActive ? 'Nieaktywny' : 'Aktywny'}
        </button>
      </td>
      <td>
        <div style="display: flex; gap: 0.5rem; flex-wrap: wrap;">
          <button class="btn btn-secondary btn-icon drag-handle btn-mobile-only" title="Przeciągnij">
            <i data-lucide="grip-horizontal" style="width:16px;height:16px"></i>
          </button>
          <button class="btn btn-secondary btn-icon btn-edit-person" data-id="${p.id}">
            <i data-lucide="edit-2" style="width:16px;height:16px"></i>
          </button>
          <button class="btn btn-danger btn-icon btn-delete-person" data-id="${p.id}">
            <i data-lucide="trash-2" style="width:16px;height:16px"></i>
          </button>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  });

  // Attach event listeners for edit and delete
  document.querySelectorAll('.btn-edit-person').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = e.currentTarget.getAttribute('data-id');
      const person = Store.getState().persons.find(p => p.id === id);
      if (person) {
        const lastNameInput = document.getElementById('person-last-name');
        const emailInput = document.getElementById('person-email');
        const employerSelect = document.getElementById('person-employer');
        const costShareCheckbox = document.getElementById('person-participates-in-costs');
        const contractTaxInput = document.getElementById('person-contract-tax');
        const contractZusInput = document.getElementById('person-contract-zus');
        const contractPaidByEmployerCheckbox = document.getElementById('person-contract-paid-by-employer');
        const sharesEmployeeProfitsCheckbox = document.getElementById('person-shares-employee-profits');
        const employeeProfitSharePercentInput = document.getElementById('person-employee-profit-share-percent');
        const receivesCompanyEmployeeProfitsCheckbox = document.getElementById('person-receives-company-employee-profits');
        const companyEmployeeProfitSharePercentInput = document.getElementById('person-company-employee-profit-share-percent');
        const countsEmployeeAccountingRefundCheckbox = document.getElementById('person-counts-employee-accounting-refund');
        document.getElementById('person-id').value = person.id;
        document.getElementById('person-name').value = person.name;
        if (lastNameInput) lastNameInput.value = person.lastName || '';
        if (emailInput) emailInput.value = person.email || '';
        document.getElementById('person-type').value = person.type;
        document.getElementById('person-rate').value = person.hourlyRate || '';
        if (employerSelect) employerSelect.value = person.employerId || '';
        if (costShareCheckbox) costShareCheckbox.checked = doesPersonParticipateInCosts(person);
        if (contractTaxInput) contractTaxInput.value = person.contractTaxAmount ?? getDefaultContractTaxAmountForType(person.type);
        if (contractZusInput) contractZusInput.value = person.contractZusAmount ?? getDefaultContractZusAmountForType(person.type);
        if (contractPaidByEmployerCheckbox) contractPaidByEmployerCheckbox.checked = doesEmployerPayPersonContractCharges(person);
        if (sharesEmployeeProfitsCheckbox) sharesEmployeeProfitsCheckbox.checked = person.sharesEmployeeProfits === true;
        if (employeeProfitSharePercentInput) employeeProfitSharePercentInput.value = normalizeProfitSharePercent(person.employeeProfitSharePercent, getDefaultProfitSharePercent());
        if (receivesCompanyEmployeeProfitsCheckbox) receivesCompanyEmployeeProfitsCheckbox.checked = person.receivesCompanyEmployeeProfits === true;
        if (companyEmployeeProfitSharePercentInput) companyEmployeeProfitSharePercentInput.value = normalizeProfitSharePercent(person.companyEmployeeProfitSharePercent, getDefaultProfitSharePercent());
        if (countsEmployeeAccountingRefundCheckbox) countsEmployeeAccountingRefundCheckbox.checked = person.countsEmployeeAccountingRefund ?? getDefaultCountsEmployeeAccountingRefundForType(person.type);
        updatePersonFormTypeUI(
          person.type,
          person.employerId || '',
          doesPersonParticipateInCosts(person),
          person.contractTaxAmount ?? getDefaultContractTaxAmountForType(person.type),
          person.contractZusAmount ?? getDefaultContractZusAmountForType(person.type),
          doesEmployerPayPersonContractCharges(person),
          person.sharesEmployeeProfits === true,
          person.employeeProfitSharePercent ?? getDefaultProfitSharePercent(),
          person.receivesCompanyEmployeeProfits === true,
          person.companyEmployeeProfitSharePercent ?? getDefaultProfitSharePercent(),
          person.countsEmployeeAccountingRefund ?? getDefaultCountsEmployeeAccountingRefundForType(person.type)
        );
        document.getElementById('person-form-title').textContent = 'Edytuj Osobę';
        document.querySelector('#people-view .table-container').style.display = 'none';
        document.getElementById('person-form').style.display = 'block';
        document.getElementById('person-form').scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    });
  });

  document.querySelectorAll('.btn-status-person').forEach(btn => {
    btn.addEventListener('click', (e) => {
      Store.togglePersonStatus(e.currentTarget.getAttribute('data-id'));
    });
  });

  document.querySelectorAll('.btn-delete-person').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = e.currentTarget.getAttribute('data-id');
      if (confirm('Czy na pewno chcesz usunąć tę osobę? Rekordy godzin mogą zostać zaburzone.')) {
        Store.deletePerson(id);
      }
    });
  });
}

// ==========================================
// MONTHLY SHEETS (Godziny)
// ==========================================

function getEaster(year) {
  const a = year % 19, b = Math.floor(year / 100), c = year % 100, d = Math.floor(b / 4), e = b % 4;
  const f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const p = (h + l - 7 * m + 114) % 31;
  return new Date(year, Math.floor((h + l - 7 * m + 114) / 31) - 1, p + 1);
}

function getPolishHolidays(year) {
  const easter = getEaster(year);
  const easterMonday = new Date(easter);
  easterMonday.setDate(easter.getDate() + 1);
  const corpusChristi = new Date(easter);
  corpusChristi.setDate(easter.getDate() + 60);

  const format = (date) => `${String(date.getMonth() + 1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`;

  return {
    '01-01': 'Nowy Rok',
    '01-06': 'Trzech Króli',
    '05-01': 'Święto Pracy',
    '05-03': 'Święto Konst. 3 Maja',
    '08-15': 'Wniebowzi. NMP',
    '11-01': 'Wszystkich Św.',
    '11-11': 'Św. Niepodleg.',
    '12-25': 'Boże Narodzenie',
    '12-26': 'Drugi dz. Świąt',
    [format(easter)]: 'Wielkanoc',
    [format(easterMonday)]: 'Pon. Wielk.',
    [format(corpusChristi)]: 'Boże Ciało'
  };
}

function getActiveClients() {
  const state = Store.getState();
  return state.clients.filter(client =>
    Store.isClientActiveInMonth(client.id)
    && Number.isFinite(parseFloat(client.hourlyRate))
    && parseFloat(client.hourlyRate) > 0
  );
}

function populateSheetClientSelect(selectedClient = '') {
  const select = document.getElementById('sheet-client-select');
  const activeClients = getActiveClients();
  const selectedClientName = typeof selectedClient === 'string'
    ? selectedClient
    : (selectedClient?.client || '');
  select.innerHTML = '<option value="">-- Wybierz z listy --</option>';

  activeClients.forEach(client => {
    const opt = document.createElement('option');
    opt.value = client.name;
    opt.textContent = `${client.name} (${parseFloat(client.hourlyRate).toFixed(2)} zł/h)`;
    if (client.name === selectedClientName) opt.selected = true;
    select.appendChild(opt);
  });

  if (selectedClientName && !activeClients.some(client => client.name === selectedClientName)) {
    const currentClient = Store.getState().clients.find(client => client.name === selectedClientName);
    if (currentClient) {
      const opt = document.createElement('option');
      opt.value = currentClient.name;
      opt.textContent = `${currentClient.name} (${parseFloat(currentClient.hourlyRate).toFixed(2)} zł/h) - nieaktywny`;
      opt.selected = true;
      select.appendChild(opt);
    }
  }
}

function parsePolishCurrency(value) {
  return parseFloat((value || '').toString().replace(/\s/g, '').replace(',', '.'));
}

function normalizePolishText(value) {
  return (value || '')
    .toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();
}

function extractCurrencyValues(value) {
  return Array.from((value || '').matchAll(/(\d[\d\s]*,\d{2})/g)).map(match => parsePolishCurrency(match[1]));
}

function extractSocialRatesFromSource(normalizedSocialText) {
  const withDirect = normalizedSocialText.match(/razem spoleczne z chorobowym\s*=\s*(\d[\d\s]*,\d{2})/i)?.[1];
  const withoutDirect = normalizedSocialText.match(/razem spoleczne bez chorobowego\s*=\s*(\d[\d\s]*,\d{2})/i)?.[1];

  if (withDirect && withoutDirect) {
    return {
      withSick: parsePolishCurrency(withDirect),
      withoutSick: parsePolishCurrency(withoutDirect),
      mode: 'direct'
    };
  }

  const socialBlockMatch = normalizedSocialText.match(/razem skladki zus bez skladki zdrowotnej([\s\S]{0,400})/i);
  if (!socialBlockMatch) return null;

  const values = extractCurrencyValues(socialBlockMatch[1]).filter(value => Number.isFinite(value) && value > 0);
  if (values.length < 2) return null;

  const uniqueValues = [...new Set(values.map(value => value.toFixed(2)))].map(value => parseFloat(value));
  const sorted = uniqueValues.sort((a, b) => a - b);
  if (sorted.length < 2) return null;

  return {
    withoutSick: sorted[0],
    withSick: sorted[sorted.length - 1],
    mode: 'block'
  };
}

function extractHealthRatesFromSource(normalizedHealthText) {
  const ryczaltBlockMatch = normalizedHealthText.match(/ryczalt od przychodow ewidencjonowanych([\s\S]{0,500})/i);
  const sourceText = ryczaltBlockMatch ? ryczaltBlockMatch[1] : normalizedHealthText;
  const healthLineMatch = sourceText.match(/skladka zdrowotna([^\n]{0,120}|[\s\S]{0,120}?)(\d[\d\s]*,\d{2})\s+(\d[\d\s]*,\d{2})\s+(\d[\d\s]*,\d{2})/i);

  if (healthLineMatch) {
    return {
      LOW: parsePolishCurrency(healthLineMatch[2]),
      MID: parsePolishCurrency(healthLineMatch[3]),
      HIGH: parsePolishCurrency(healthLineMatch[4]),
      mode: 'direct'
    };
  }

  const values = extractCurrencyValues(sourceText).filter(value => Number.isFinite(value) && value > 0);
  const sorted = [...new Set(values.map(value => value.toFixed(2)))].map(value => parseFloat(value)).sort((a, b) => a - b);
  if (sorted.length < 3) return null;

  return {
    LOW: sorted[0],
    MID: sorted[1],
    HIGH: sorted[2],
    mode: 'block'
  };
}

const FALLBACK_ZUS_RYCZALT_RATES = {
  2026: {
    social: { withSick: 1926.76, withoutSick: 1788.29 },
    health: { LOW: 498.35, MID: 830.58, HIGH: 1495.04 },
    source: 'fallback-2026'
  },
  2025: {
    social: { withSick: 1773.96, withoutSick: 1646.47 },
    health: { LOW: 461.66, MID: 769.43, HIGH: 1384.97 },
    source: 'fallback-2025'
  }
};

function extractZusSocialRates(normalizedSocialText) {
  const withSickMatches = Array.from(normalizedSocialText.matchAll(/razem spoleczne z chorobowym\s*=\s*(\d[\d\s]*,\d{2})/g));
  const withoutSickMatches = Array.from(normalizedSocialText.matchAll(/razem spoleczne bez chorobowego\s*=\s*(\d[\d\s]*,\d{2})/g));

  if (withSickMatches.length > 0 && withoutSickMatches.length > 0) {
    const withSick = parsePolishCurrency(withSickMatches[0][1]);
    const withoutSick = parsePolishCurrency(withoutSickMatches[0][1]);
    if (Number.isFinite(withSick) && Number.isFinite(withoutSick)) {
      return { withSick, withoutSick };
    }
  }

  const socialLine = normalizedSocialText
    .split(/\r?\n/)
    .map(line => line.trim())
    .find(line => line.includes('razem skladki zus bez skladki zdrowotnej') && line.includes('chorobowa'));

  if (!socialLine) return null;

  const values = extractCurrencyValues(socialLine).filter(value => Number.isFinite(value) && value > 0);
  if (values.length < 2) return null;

  const sorted = [...values].sort((a, b) => a - b);
  return {
    withoutSick: sorted[0],
    withSick: sorted[sorted.length - 1]
  };
}

function extractRyczaltHealthRates(normalizedHealthText) {
  const healthLines = normalizedHealthText
    .split(/\r?\n/)
    .map(line => line.trim())
    .filter(line => line.includes('skladka zdrowotna'));

  for (const line of healthLines) {
    const values = extractCurrencyValues(line).filter(value => Number.isFinite(value) && value > 0);
    if (values.length >= 3) {
      const sorted = [...values].sort((a, b) => a - b);
      return {
        LOW: sorted[0],
        MID: sorted[1],
        HIGH: sorted[2]
      };
    }
  }

  return null;
}

async function fetchCurrentZusRateForRyczalt(options = {}) {
  const currentYear = new Date().getFullYear();
  const includeSick = options.includeSick === true;
  const threshold = options.threshold || 'MID';
  const debugAttempts = [];
  const thresholdLabels = {
    LOW: 'Do 60 000 zł przychodu',
    MID: '60 000 – 300 000 zł przychodu',
    HIGH: 'Powyżej 300 000 zł przychodu'
  };

  for (const year of [currentYear, currentYear - 1]) {
    const fallback = FALLBACK_ZUS_RYCZALT_RATES[year];
    const [socialResponse, healthResponse] = await Promise.all([
      fetch(`https://r.jina.ai/http://zus.pox.pl/skladki-zus-${year}.htm`),
      fetch(`https://r.jina.ai/http://zus.pox.pl/aktualne-skladki-zus-ryczalt.htm`)
    ]);

    if (!socialResponse.ok || !healthResponse.ok) {
      debugAttempts.push(`rok ${year}: social=${socialResponse.status}, health=${healthResponse.status}`);
      if (fallback) {
        const fallbackSocial = includeSick ? fallback.social.withSick : fallback.social.withoutSick;
        const fallbackHealth = fallback.health[threshold];
        debugAttempts.push(`rok ${year}: użyto fallback po statusie social=${fallbackSocial}, health=${fallbackHealth}`);
        return {
          amount: Number((fallbackSocial + fallbackHealth).toFixed(2)),
          year,
          source: fallback.source,
          socialAmount: fallbackSocial,
          healthAmount: fallbackHealth,
          includeSick,
          thresholdLabel: thresholdLabels[threshold] || thresholdLabels.MID,
          debugInfo: debugAttempts
        };
      }
      continue;
    }

    const socialText = await socialResponse.text();
    const healthText = await healthResponse.text();

    const directWithSick = socialText.match(/Razem społeczne z chorobowym\s*=\s*(\d[\d\s]*,\d{2})/i)?.[1];
    const directWithoutSick = socialText.match(/Razem społeczne bez chorobowego\s*=\s*(\d[\d\s]*,\d{2})/i)?.[1];
    const directHealthMatch = healthText.match(/\[Składka zdrowotna\][^\n]*?(\d[\d\s]*,\d{2})\s+(\d[\d\s]*,\d{2})\s+(\d[\d\s]*,\d{2})/i)
      || healthText.match(/Składka zdrowotna[^\n]*?(\d[\d\s]*,\d{2})\s+(\d[\d\s]*,\d{2})\s+(\d[\d\s]*,\d{2})/i);

    const normalizedSocialText = normalizePolishText(socialText);
    const normalizedHealthText = normalizePolishText(healthText);
    const socialRates = (directWithSick && directWithoutSick)
      ? {
          withSick: parsePolishCurrency(directWithSick),
          withoutSick: parsePolishCurrency(directWithoutSick),
          mode: 'raw-direct'
        }
      : extractSocialRatesFromSource(normalizedSocialText);
    const healthRates = directHealthMatch
      ? {
          LOW: parsePolishCurrency(directHealthMatch[1]),
          MID: parsePolishCurrency(directHealthMatch[2]),
          HIGH: parsePolishCurrency(directHealthMatch[3]),
          mode: 'raw-direct'
        }
      : extractHealthRatesFromSource(normalizedHealthText);

    if (!socialRates || !healthRates) {
      debugAttempts.push(`rok ${year}: social=${socialRates ? socialRates.mode : 'brak'}, health=${healthRates ? healthRates.mode : 'brak'}`);
      if (fallback) {
        const fallbackSocial = includeSick ? fallback.social.withSick : fallback.social.withoutSick;
        const fallbackHealth = fallback.health[threshold];
        debugAttempts.push(`rok ${year}: użyto fallback po braku dopasowania social=${fallbackSocial}, health=${fallbackHealth}`);
        return {
          amount: Number((fallbackSocial + fallbackHealth).toFixed(2)),
          year,
          source: fallback.source,
          socialAmount: fallbackSocial,
          healthAmount: fallbackHealth,
          includeSick,
          thresholdLabel: thresholdLabels[threshold] || thresholdLabels.MID,
          debugInfo: debugAttempts
        };
      }
      continue;
    }

    const socialAmount = includeSick ? socialRates.withSick : socialRates.withoutSick;
    const healthAmount = healthRates[threshold];

    debugAttempts.push(`rok ${year}: socialMode=${socialRates.mode}, healthMode=${healthRates.mode}, with=${socialRates.withSick}, without=${socialRates.withoutSick}, healthLow=${healthRates.LOW}, healthMid=${healthRates.MID}, healthHigh=${healthRates.HIGH}`);

    if (Number.isFinite(socialAmount) && socialAmount > 0 && Number.isFinite(healthAmount) && healthAmount > 0) {
      return {
        amount: Number((socialAmount + healthAmount).toFixed(2)),
        year,
        source: `zus.pox.pl/${year}`,
        socialAmount,
        healthAmount,
        includeSick,
        thresholdLabel: thresholdLabels[threshold] || thresholdLabels.MID,
        debugInfo: debugAttempts
      };
    }

    debugAttempts.push(`rok ${year}: wynik niepoprawny social=${socialAmount}, health=${healthAmount}`);

    if (fallback) {
      const fallbackSocial = includeSick ? fallback.social.withSick : fallback.social.withoutSick;
      const fallbackHealth = fallback.health[threshold];
      if (Number.isFinite(fallbackSocial) && Number.isFinite(fallbackHealth)) {
        debugAttempts.push(`rok ${year}: użyto fallback social=${fallbackSocial}, health=${fallbackHealth}`);
        return {
          amount: Number((fallbackSocial + fallbackHealth).toFixed(2)),
          year,
          source: fallback.source,
          socialAmount: fallbackSocial,
          healthAmount: fallbackHealth,
          includeSick,
          thresholdLabel: thresholdLabels[threshold] || thresholdLabels.MID,
          debugInfo: debugAttempts
        };
      }
    }
  }

  for (const year of [currentYear, currentYear - 1]) {
    const fallback = FALLBACK_ZUS_RYCZALT_RATES[year];
    if (!fallback) continue;

    const fallbackSocial = includeSick ? fallback.social.withSick : fallback.social.withoutSick;
    const fallbackHealth = fallback.health[threshold];
    if (Number.isFinite(fallbackSocial) && Number.isFinite(fallbackHealth)) {
      debugAttempts.push(`rok ${year}: użyto końcowego fallback social=${fallbackSocial}, health=${fallbackHealth}`);
      return {
        amount: Number((fallbackSocial + fallbackHealth).toFixed(2)),
        year,
        source: fallback.source,
        socialAmount: fallbackSocial,
        healthAmount: fallbackHealth,
        includeSick,
        thresholdLabel: thresholdLabels[threshold] || thresholdLabels.MID,
        debugInfo: debugAttempts
      };
    }
  }

  const error = new Error('Nie udało się odczytać aktualnej stawki ZUS dla ryczałtu.');
  error.debugInfo = debugAttempts;
  throw error;
}

window.currentSheetId = null;

function isMobilePortraitLayout() {
  return window.matchMedia('(max-width: 768px)').matches;
}

function revealFocusedField(element) {
  if (!element || !isMobilePortraitLayout()) return;

  const performScroll = () => {
    const rect = element.getBoundingClientRect();
    const viewportHeight = window.visualViewport ? window.visualViewport.height : window.innerHeight;
    const topInset = 84;
    const bottomInset = 24;

    if (rect.top < topInset || rect.bottom > viewportHeight - bottomInset) {
      element.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'nearest' });
    }
  };

  requestAnimationFrame(performScroll);
  setTimeout(performScroll, 250);
  setTimeout(performScroll, 500);
}

function initMobileInputVisibilityHandling() {
  document.addEventListener('focusin', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    if (!target.matches('input, select, textarea')) return;

    revealFocusedField(target);
  });

  if (window.visualViewport) {
    window.visualViewport.addEventListener('resize', () => {
      const activeElement = document.activeElement;
      if (activeElement instanceof HTMLElement && activeElement.matches('input, select, textarea')) {
        revealFocusedField(activeElement);
      }
    });
  }
}

function bindSheetEditingUX() {
  document.querySelectorAll('#monthly-hours-body td').forEach(cell => {
    cell.addEventListener('focusout', () => {
      requestAnimationFrame(() => {
        if (!cell.contains(document.activeElement)) {
          cell.classList.remove('is-editing');
        }
      });
    });
  });

  document.querySelectorAll('.hour-in').forEach(input => {
    const activateCell = () => {
      const cell = input.closest('td');
      if (!cell) return;
      cell.classList.add('is-editing');
      revealFocusedField(input);
    };

    input.addEventListener('focus', activateCell);
    input.addEventListener('click', activateCell);
    input.addEventListener('input', activateCell);
  });
}

function populateRatesGridForMetaForm(sheet) {
  const state = Store.getState();
  const ratesGrid = document.getElementById('sheet-rates-grid');
  ratesGrid.innerHTML = '';
  
  const clientInput = document.getElementById('sheet-client-select');
  const selectedClientName = clientInput.value;
  let clientRate = 0;
  if (selectedClientName) {
    const client = state.clients.find(c => c.name === selectedClientName);
    if (client) {
      clientRate = parseFloat(client.hourlyRate) || 0;
    }
  }

  // Create a dummy sheet if null for calculations, but using current client
  const dummySheet = sheet
    ? { ...sheet }
    : { month: getSelectedMonthKey(), client: selectedClientName, personsConfig: {}, days: {} };
  const activePersonsContainer = document.getElementById('sheet-meta-active-persons-list');
  if (activePersonsContainer && activePersonsContainer.querySelector('.cb-sheet-participant')) {
    dummySheet.activePersons = collectMonthlySheetActivePersons(activePersonsContainer);
  }
  const sheetClientRateOverride = parseFloat(document.getElementById('sheet-client-rate').value);
  const effectiveSheetClientRate = (Number.isFinite(sheetClientRateOverride) && sheetClientRateOverride > 0) ? sheetClientRateOverride : clientRate;

  const activePersonsForConfig = getVisiblePersonsForSheet(state, dummySheet);
  activePersonsForConfig.forEach(p => {
    const custom = (dummySheet.personsConfig && dummySheet.personsConfig[p.id] && dummySheet.personsConfig[p.id].customRate) ? dummySheet.personsConfig[p.id].customRate : '';
    
    // Quick default rate calc without full sheet simulation
    let defaultRate = p.hourlyRate || 0;
    if (p.type === 'WORKING_PARTNER') {
      defaultRate = p.hourlyRate || effectiveSheetClientRate;
    } else if (p.type === 'PARTNER') {
      defaultRate = p.hourlyRate || effectiveSheetClientRate;
    }

    ratesGrid.innerHTML += `
      <div style="background: rgba(255,255,255,0.05); padding: 0.5rem; border-radius: 4px;">
        <label style="font-size: 0.8rem; display: block; margin-bottom: 0.3rem;">${getPersonDisplayName(p)} (Dom: ${defaultRate.toFixed(2)} zł/h${(!p.hourlyRate && effectiveSheetClientRate > 0) ? ' - stawka arkusza' : ''})</label>
        <input type="number" step="0.1" class="rate-input-custom" data-id="${p.id}" value="${custom}" placeholder="Inna stawka">
      </div>
    `;
  });
}

function getDefaultMonthlySheetActivePersonIds(month = getSelectedMonthKey()) {
  return Store.getState().persons
    .filter(person => Store.isPersonActiveInMonth(person.id, month))
    .map(person => person.id);
}

function getMonthlySheetActivePersonIds(sheet = null, month = getSelectedMonthKey()) {
  return Array.isArray(sheet?.activePersons)
    ? [...new Set(sheet.activePersons)]
    : getDefaultMonthlySheetActivePersonIds(month);
}

function renderMonthlySheetActivePersonsList(container, sheet = null) {
  if (!container) return;

  const month = sheet?.month || getSelectedMonthKey();
  const allPersons = Store.getState().persons.filter(person => Store.isPersonActiveInMonth(person.id, month));
  const activePersonIds = new Set(getMonthlySheetActivePersonIds(sheet, month));

  if (allPersons.length === 0) {
    container.innerHTML = '<p style="color: var(--text-muted);">Brak aktywnych osób w tym miesiącu.</p>';
    return;
  }

  container.innerHTML = allPersons.map(person => {
    const { badgeClass, badgeLabel } = getWorksPersonBadgeMeta(person);
    return `
      <div class="active-person-card">
        <div class="active-person-card-header">
          <label class="active-person-toggle">
            <input type="checkbox" class="cb-sheet-participant" data-pid="${person.id}" ${activePersonIds.has(person.id) ? 'checked' : ''} style="width: 19px; height: 19px; margin: 0;">
            <span class="active-person-name">${getPersonDisplayName(person)}</span>
          </label>
          <span class="badge ${badgeClass}" style="padding: 0.3rem 0.7rem; font-size: 0.72rem; letter-spacing: 0.02em;">${badgeLabel}</span>
        </div>
      </div>
    `;
  }).join('');
}

function collectMonthlySheetActivePersons(container) {
  const scope = container || document;
  return Array.from(scope.querySelectorAll('.cb-sheet-participant:checked'))
    .map(checkbox => checkbox.getAttribute('data-pid'));
}

function openMonthlySheetMetaForm(sheet = null) {
  const listContainer = document.getElementById('hours-sheet-list');
  const form = document.getElementById('sheet-meta-form');
  const detailContainer = document.getElementById('sheet-detail-container');
  const activePersonsForm = document.getElementById('sheet-active-persons-form');
  const activePersonsList = document.getElementById('sheet-meta-active-persons-list');
  const isEditing = !!sheet;
  const titleEl = document.getElementById('sheet-meta-title') || document.getElementById('sheet-form-title');

  if (titleEl) titleEl.textContent = isEditing ? 'Edytuj Arkusz Godzin' : 'Nowy Arkusz Godzin';
  document.getElementById('sheet-id').value = isEditing ? sheet.id : '';
  document.getElementById('sheet-client-select').value = '';
  document.getElementById('sheet-site').value = isEditing ? (sheet.site || '') : '';
  document.getElementById('sheet-client-rate').value = isEditing && sheet.clientRateOverride ? sheet.clientRateOverride : '';

  populateSheetClientSelect(isEditing ? sheet.client : '');
  renderMonthlySheetActivePersonsList(activePersonsList, sheet);
  populateRatesGridForMetaForm(sheet);

  listContainer.style.display = 'none';
  detailContainer.style.display = 'none';
  if (activePersonsForm) activePersonsForm.style.display = 'none';
  form.style.display = 'block';
}

function initMonthlySheets() {
  const listContainer = document.getElementById('hours-sheet-list');
  const form = document.getElementById('sheet-meta-form');
  const detailContainer = document.getElementById('sheet-detail-container');
  const activePersonsForm = document.getElementById('sheet-active-persons-form');
  const activePersonsList = document.getElementById('sheet-active-persons-list');
  const metaActivePersonsList = document.getElementById('sheet-meta-active-persons-list');

  document.getElementById('btn-create-sheet').addEventListener('click', () => {
    const activeClients = getActiveClients();
    if (activeClients.length === 0) {
      alert('Dodaj aktywnego klienta ze stawką, zanim utworzysz arkusz godzin.');
      return;
    }

    openMonthlySheetMetaForm(null);
  });

  const clientSelect = document.getElementById('sheet-client-select');
  const clientRateInput = document.getElementById('sheet-client-rate');
  clientSelect.addEventListener('change', () => {
    const sheetId = document.getElementById('sheet-id').value;
    const sheet = sheetId ? Store.getMonthlySheet(sheetId) : null;
    populateRatesGridForMetaForm(sheet);
  });
  clientRateInput.addEventListener('input', () => {
    const sheetId = document.getElementById('sheet-id').value;
    const sheet = sheetId ? Store.getMonthlySheet(sheetId) : null;
    populateRatesGridForMetaForm(sheet);
  });

  if (metaActivePersonsList) {
    metaActivePersonsList.addEventListener('change', (e) => {
      if (!e.target.classList.contains('cb-sheet-participant')) return;
      const sheetId = document.getElementById('sheet-id').value;
      const sheet = sheetId ? Store.getMonthlySheet(sheetId) : null;
      populateRatesGridForMetaForm(sheet);
    });
  }

  document.getElementById('btn-cancel-sheet').addEventListener('click', () => {
    form.style.display = 'none';
    listContainer.style.display = 'block';
  });

  document.getElementById('btn-back-to-sheets').addEventListener('click', () => {
    window.currentSheetId = null;
    detailContainer.style.display = 'none';
    if (activePersonsForm) activePersonsForm.style.display = 'none';
    listContainer.style.display = 'block';
    renderMonthlySheets();
  });

  const btnSheetActivePersons = document.getElementById('btn-sheet-active-persons');
  if (btnSheetActivePersons) {
    btnSheetActivePersons.addEventListener('click', () => {
      if (!window.currentSheetId) return;

      const sheet = Store.getMonthlySheet(window.currentSheetId);
      if (!sheet) return;

      renderMonthlySheetActivePersonsList(activePersonsList, sheet);
      listContainer.style.display = 'none';
      detailContainer.style.display = 'none';
      form.style.display = 'none';
      activePersonsForm.style.display = 'block';
      activePersonsForm.scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  }

  const btnCloseSheetActivePersons = document.getElementById('btn-close-sheet-active-persons');
  if (btnCloseSheetActivePersons) {
    btnCloseSheetActivePersons.addEventListener('click', () => {
      if (activePersonsForm) activePersonsForm.style.display = 'none';
      if (window.currentSheetId) {
        detailContainer.style.display = 'block';
      } else {
        listContainer.style.display = 'block';
      }
    });
  }

  if (activePersonsForm) {
    activePersonsForm.addEventListener('submit', (e) => {
      e.preventDefault();
      if (!window.currentSheetId) return;

      const activePersons = collectMonthlySheetActivePersons(activePersonsForm);
      Store.updateMonthlySheet(window.currentSheetId, { activePersons });
      activePersonsForm.style.display = 'none';
      renderSheetDetail(window.currentSheetId);
    });
  }

  form.addEventListener('submit', (e) => {
    e.preventDefault();
    const id = document.getElementById('sheet-id').value;
    const month = getSelectedMonthKey();
    const client = document.getElementById('sheet-client-select').value;
    const site = document.getElementById('sheet-site').value;

    const clientRateRaw = document.getElementById('sheet-client-rate').value.trim();
    const parsedClientRate = clientRateRaw === '' ? null : parseFloat(clientRateRaw);
    const activePersons = collectMonthlySheetActivePersons(metaActivePersonsList);

    if (!client) {
      alert('Wybierz klienta z listy przed zapisaniem arkusza.');
      document.getElementById('sheet-client-select').focus();
      return;
    }
    
    if (clientRateRaw !== '' && (!Number.isFinite(parsedClientRate) || parsedClientRate <= 0)) {
       alert('Stawka klienta dla arkusza musi być większa od zera albo pusta.');
       document.getElementById('sheet-client-rate').focus();
       return;
    }

    const state = Store.getState();
    const existingSheet = id ? state.monthlySheets.find(s => s.id === id) : null;
    const personsConfig = existingSheet ? { ...existingSheet.personsConfig } : {};
    const inputs = document.querySelectorAll('.rate-input-custom');
    
    inputs.forEach(inp => {
      const pid = inp.getAttribute('data-id');
      const val = parseFloat(inp.value);
      if (!personsConfig[pid]) personsConfig[pid] = {};
      if (!isNaN(val)) personsConfig[pid].customRate = val;
      else personsConfig[pid].customRate = null;
    });
    
    if (id) {
      Store.updateMonthlySheet(id, { month, client, site, clientRateOverride: parsedClientRate, personsConfig });
      window.currentSheetId = id;
    } else {
      const newSheet = { month, client, site, personsConfig, clientRateOverride: parsedClientRate, days: {} };
      Store.addMonthlySheet(newSheet);
      const newState = Store.getState();
      window.currentSheetId = newState.monthlySheets[newState.monthlySheets.length - 1].id;
    }
    
    form.style.display = 'none';
    renderMonthlySheets();
    renderSheetDetail(window.currentSheetId);
  });
}

function renderMonthlySheets() {
  const state = Store.getState();
  const tbody = document.getElementById('hours-sheets-body');
  const selectedMonth = getSelectedMonthKey();
  
  // Only render if we are currently not in detail view, or if we just want to update list
  if (window.currentSheetId && document.getElementById('sheet-detail-container').style.display === 'block') {
    return; // Already viewing detail
  }

  tbody.innerHTML = '';
  document.getElementById('hours-sheet-list').style.display = 'block';
  document.getElementById('sheet-detail-container').style.display = 'none';
  document.getElementById('sheet-meta-form').style.display = 'none';
  const activePersonsForm = document.getElementById('sheet-active-persons-form');
  if (activePersonsForm) activePersonsForm.style.display = 'none';

  const sheetsForMonth = state.monthlySheets.filter(sheet => sheet.month === selectedMonth);

  if (sheetsForMonth.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;color:var(--text-muted)">Brak arkuszy. Kliknij "Nowy Arkusz".</td></tr>`;
    return;
  }

  const clientOrder = new Map(state.clients.map((client, index) => [client.name, index]));
  const sorted = [...sheetsForMonth].sort((a, b) => {
    const clientDiff = (clientOrder.get(a.client) ?? Number.MAX_SAFE_INTEGER) - (clientOrder.get(b.client) ?? Number.MAX_SAFE_INTEGER);
    if (clientDiff !== 0) return clientDiff;
    return (a.site || '').localeCompare(b.site || '') || a.id.localeCompare(b.id);
  });

  sorted.forEach(s => {
    const visiblePersonIds = new Set(getVisiblePersonsForSheet(state, s).map(person => person.id));
    const totalH = Calculations.getSheetTotalHours(s, visiblePersonIds);

    const sheetRevenue = totalH * Calculations.getSheetClientRate(s, state);

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${s.client || '-'}</td>
      <td>${s.site || '-'}</td>
      <td style="color: var(--primary); font-weight: bold;">${totalH.toFixed(1)}h</td>
      <td style="color: var(--success); font-weight: bold;">${sheetRevenue.toFixed(2)} zł</td>
      <td>
        <button class="btn btn-primary btn-sm btn-open-sheet" data-id="${s.id}" style="margin-right:0.5rem;">Otwórz</button>
        <button class="btn btn-secondary btn-icon btn-edit-sheet" data-id="${s.id}" style="margin-right:0.5rem;">
          <i data-lucide="edit-2" style="width:16px;height:16px"></i>
        </button>
        <button class="btn btn-danger btn-icon btn-delete-sheet" data-id="${s.id}">
          <i data-lucide="trash-2" style="width:16px;height:16px"></i>
        </button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  document.querySelectorAll('.btn-open-sheet').forEach(btn => {
  btn.addEventListener('click', (e) => {
    const id = e.currentTarget.getAttribute('data-id');
    window.currentSheetId = id;
    renderSheetDetail(id);

    // Scroll the sheet detail table into view for all screen sizes
    const containerElement = document.querySelector('#sheet-detail-container .table-container');
    if (containerElement) {
      containerElement.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  });
});

  document.querySelectorAll('.btn-edit-sheet').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = e.currentTarget.getAttribute('data-id');
      const sheet = state.monthlySheets.find(s => s.id === id);
      if (sheet) {
        openMonthlySheetMetaForm(sheet);
      }
    });
  });

  document.querySelectorAll('.btn-delete-sheet').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = e.currentTarget.getAttribute('data-id');
      if (confirm('Usunąć ten arkusz? Wszystkie zapisane w nim godziny przepadną.')) {
        Store.deleteMonthlySheet(id);
        renderMonthlySheets();
      }
    });
  });
  
  lucide.createIcons();
}

function renderSheetDetail(sheetId) {
  const state = Store.getState();
  const sheet = Store.getMonthlySheet(sheetId);
  if (!sheet) {
    renderMonthlySheets();
    return;
  }

  const sheetClientRate = Calculations.getSheetClientRate(sheet, state);

  document.getElementById('hours-sheet-list').style.display = 'none';
  const activePersonsForm = document.getElementById('sheet-active-persons-form');
  if (activePersonsForm) activePersonsForm.style.display = 'none';
  document.getElementById('sheet-detail-container').style.display = 'block';

  document.getElementById('detail-sheet-title').textContent = `${sheet.month} - ${sheet.client}`;
  document.getElementById('detail-sheet-site').textContent = sheet.site || 'Brak wpisanej budowy';

  // Build table
  const thead = document.getElementById('monthly-hours-head');
  const tbody = document.getElementById('monthly-hours-body');
  const tfoot = document.getElementById('monthly-hours-foot');
  
  // Filter people for the table: only active ones
  const activePersons = getVisiblePersonsForSheet(state, sheet);

  // Header
  let headHtml = `<th style="width: 40px; text-align: center;" title="Wypełnij dla wszystkich"></th>
                  <th style="text-align: center;">Dzień</th>
                  <th style="min-width: 120px; text-align: center;">Czas</th>`;
  activePersons.forEach(p => {
    headHtml += `<th style="text-align: center;">${getPersonCompactHeaderHtml(p)}</th>`;
  });
  thead.innerHTML = headHtml;

  tbody.innerHTML = '';
  
  const parts = sheet.month.split('-');
  const year = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10);
  const daysInMonth = new Date(year, month, 0).getDate();

  const sums = {};
  activePersons.forEach(p => sums[p.id] = 0);

  const holidays = getPolishHolidays(year);

  for (let d = 1; d <= daysInMonth; d++) {
    const dayData = (sheet.days && sheet.days[d]) ? sheet.days[d] : {};
    const globalStart = dayData.globalStart || '';
    const globalEnd = dayData.globalEnd || '';
    const isChecked = dayData.isWholeTeamChecked ? 'checked' : '';
    
    // Check if it's Sunday or holiday
    const dateObj = new Date(year, month - 1, d);
    const dateStr = `${String(month).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    const holidayName = holidays[dateStr];
    const isHoliday = !!holidayName;
    const isSunday = dateObj.getDay() === 0;
    const isSaturday = dateObj.getDay() === 6;
    
    const dayNames = ['niedziela', 'poniedziałek', 'wtorek', 'środa', 'czwartek', 'piątek', 'sobota'];
    const smallText = isHoliday ? holidayName : dayNames[dateObj.getDay()];

    let rowBg = '';
    let rowClass = 'day-row';
    if (isSunday || isHoliday) {
      rowBg = 'background: rgba(255, 165, 0, 0.4);';
      rowClass += ' is-holiday';
    } else if (isSaturday) {
      rowBg = 'background: rgba(255, 235, 153, 0.2);';
      rowClass += ' is-saturday';
    }

    let tr = document.createElement('tr');
    tr.style = rowBg;
    tr.className = rowClass;
    
    let html = `
      <td style="text-align: center; padding: 0.35rem 0.25rem;">
        <input type="checkbox" class="fill-day-cb" data-day="${d}" ${isChecked} style="width: 18px; height: 18px; cursor: pointer;">
      </td>
      <td style="font-weight: bold; line-height: 1.1; text-align: center; padding: 0.35rem 0.3rem;">
        ${d}<br>
        <span style="font-size: 0.6rem; color: var(--text-secondary); font-weight: normal; white-space: nowrap;">${smallText}</span>
      </td>
      <td style="font-size: 0.72rem; display: flex; gap: 2px; justify-content: center; align-items: center; border: none; padding: 0.35rem 0.25rem;">
        <input type="text" class="time-in" data-day="${d}" data-type="start" readonly value="${globalStart}" placeholder="--:--" style="width: 44px; padding: 0.05rem; background: transparent; text-align: center; cursor: pointer; border: 1px solid transparent; border-radius: 4px;">
        <span style="opacity: 0.5">-</span>
        <input type="text" class="time-in" data-day="${d}" data-type="end" readonly value="${globalEnd}" placeholder="--:--" style="width: 44px; padding: 0.05rem; background: transparent; text-align: center; cursor: pointer; border: 1px solid transparent; border-radius: 4px;">
      </td>
    `;

    activePersons.forEach(p => {
      const hours = (dayData.hours && dayData.hours[p.id] !== undefined) ? dayData.hours[p.id] : '';
      const isInactive = isMonthlySheetPersonInactiveOnDay(sheet, p.id, d);
      if (!isInactive && hours !== '') sums[p.id] += parseFloat(hours);
      
      const isManual = dayData.manual && dayData.manual[p.id];
      const isZero = hours === 0 && isManual;
      
      let valToShow = hours;
      let extraClass = '';
      if (isInactive) {
        valToShow = 0;
        extraClass = 'is-inactive';
      } else if (isZero) {
        valToShow = 0;
        extraClass = 'is-zero';
      }

      let manualStyle = isManual ? 'background: rgba(255, 255, 255, 0.2);' : 'background: transparent; border: none;';
      if (isInactive) {
        manualStyle = 'background: rgba(148, 163, 184, 0.18); color: var(--text-secondary); border: 1px dashed var(--border-color);';
      } else if (isZero) {
        manualStyle = 'background: rgba(239, 68, 68, 0.2);'; // Light red for 0
      }
      
      html += `<td style="text-align: center; position: relative; padding: 0.25rem 0.35rem;">
                <input type="number" step="0.5" class="hour-in ${extraClass}" data-day="${d}" data-id="${p.id}" value="${valToShow}" style="${manualStyle} text-align: center; width: 100%; border-radius: 4px; padding: 1px 2px;">
                <div class="cell-actions">
                  <button class="action-btn btn-set-zero" data-day="${d}" data-id="${p.id}" title="Ustaw 0 (nieobecność)">X</button>
                  <button class="action-btn btn-reset-manual" data-day="${d}" data-id="${p.id}" title="Przywróć automat">
                    <svg style="width: 12px; height: 12px;" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 0 0-9-9 9.75 9 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/><path d="M3 12a9 9 0 0 0 9 9 9.75 9 0 0 0 6.74-2.74L21 16"/><path d="M16 16h5v5"/></svg>
                  </button>
                </div>
               </td>`;
    });

    tr.innerHTML = html;
    tbody.appendChild(tr);
  }

  // Footer sums and earnings combined
  let footHtml = `<td></td><td><strong>RAZEM:</strong></td><td></td>`;
  
  activePersons.forEach(p => {
    const s = sums[p.id];
    const effectiveRate = Calculations.getSheetPersonRate(p, sheet, state);
    const earnings = s * effectiveRate;
    
    footHtml += `<td style="text-align: center; line-height: 1.15; padding: 0.35rem 0.1rem;">
                   <div style="color: var(--success); font-weight: bold; font-size: 1rem;">${s.toFixed(1)}h</div>
                   <div style="color: var(--primary); font-size: 0.8rem; opacity: 0.9;">${earnings.toFixed(2)} zł</div>
                 </td>`;
  });
  tfoot.innerHTML = `<tr>${footHtml}</tr>`;

  const totalSheetHours = activePersons.reduce((sum, person) => sum + sums[person.id], 0);
  const totalSheetRevenue = totalSheetHours * sheetClientRate;
  document.getElementById('sheet-total-hours-summary').textContent = `${totalSheetHours.toFixed(1)}h`;
  document.getElementById('sheet-total-revenue-summary').textContent = `${totalSheetRevenue.toFixed(2)} zł`;

  // Attach auto-save events
  document.querySelectorAll('.time-in').forEach(inp => {
    inp.addEventListener('click', (e) => {
       showCustomTimePicker(e.target);
    });
  });

  document.querySelectorAll('.hour-in').forEach(inp => {
    inp.addEventListener('change', (e) => {
      const d = e.target.getAttribute('data-day');
      const dayNumber = parseInt(d, 10);
      const pid = e.target.getAttribute('data-id');
      const val = e.target.value;
      
      const s = Store.getMonthlySheet(window.currentSheetId);
      if (!s.days) s.days = {};
      if (!s.days[d]) s.days[d] = { hours: {} };
      if (!s.days[d].manual) s.days[d].manual = {};
      
      if (val !== '') {
        s.days[d].hours[pid] = parseFloat(val);
        s.days[d].manual[pid] = true; // Mark as manually overridden
        if (isMonthlySheetPersonInactiveOnDay(s, pid, dayNumber) || getMonthlySheetPersonActivityOverride(s, d, pid) === 'inactive') {
          setMonthlySheetPersonActivityOverride(s, d, pid, 'active');
        }
        e.target.style.background = 'rgba(255, 255, 255, 0.25)';
        e.target.style.textAlign = 'center';
      } else {
        delete s.days[d].hours[pid];
        clearMonthlySheetPersonManualFlag(s.days[d], pid);
        e.target.style.background = 'transparent';
        e.target.style.border = 'none';
      }
      
      Store.updateMonthlySheet(s.id, { days: s.days });
      renderSheetDetail(window.currentSheetId);
    });
  });

  bindSheetEditingUX();

  document.querySelectorAll('.btn-set-zero').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const target = e.target.closest('.btn-set-zero');
      const d = target.getAttribute('data-day');
      const dayNumber = parseInt(d, 10);
      const pid = target.getAttribute('data-id');
      
      const s = Store.getMonthlySheet(window.currentSheetId);
      if (!s.days) s.days = {};
      if (!s.days[d]) s.days[d] = { hours: {} };
      if (!s.days[d].manual) s.days[d].manual = {};

      const currentHours = s.days[d].hours ? s.days[d].hours[pid] : undefined;
      const isManualZero = currentHours === 0 && s.days[d].manual[pid] === true;
      const isInactive = isMonthlySheetPersonInactiveOnDay(s, pid, dayNumber);
      
      s.days[d].hours[pid] = 0;
      s.days[d].manual[pid] = true;

      if (!isInactive && isManualZero) {
        setMonthlySheetPersonActivityOverride(s, d, pid, 'inactive');
      }
      
      Store.updateMonthlySheet(s.id, { days: s.days });
      renderSheetDetail(window.currentSheetId);
    });
  });

  document.querySelectorAll('.btn-reset-manual').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const target = e.target.closest('.btn-reset-manual');
      const d = target.getAttribute('data-day');
      const dayNumber = parseInt(d, 10);
      const pid = target.getAttribute('data-id');
      
      const s = Store.getMonthlySheet(window.currentSheetId);
      if (!s.days) s.days = {};
      if (!s.days[d]) s.days[d] = { hours: {} };

      const todayOverride = getMonthlySheetPersonActivityOverride(s, d, pid);
      const isInactive = isMonthlySheetPersonInactiveOnDay(s, pid, dayNumber);

      clearMonthlySheetPersonManualFlag(s.days[d], pid);

      if (isInactive) {
        if (todayOverride === 'inactive') {
          setMonthlySheetPersonActivityOverride(s, d, pid, null);
        } else {
          setMonthlySheetPersonActivityOverride(s, d, pid, 'active');
        }
      }

      syncMonthlySheetDayPersonHours(s, d, pid);

      Store.updateMonthlySheet(s.id, { days: s.days });
      renderSheetDetail(window.currentSheetId);
    });
  });

  document.querySelectorAll('.fill-day-cb').forEach(cb => {
    cb.addEventListener('change', (e) => {
      const d = e.target.getAttribute('data-day');
      const isChecked = e.target.checked;
      
      const s = Store.getMonthlySheet(window.currentSheetId);
      if (!s.days) s.days = {};
      if (!s.days[d]) s.days[d] = { hours: {} };
      
      s.days[d].isWholeTeamChecked = isChecked;
      
      if (isChecked) {
        const start = s.days[d].globalStart || '07:00';
        const end = s.days[d].globalEnd || '17:00';
        s.days[d].globalStart = start;
        s.days[d].globalEnd = end;
        
        const calcH = Calculations.calculateHours(start, end);
        if (calcH > 0) {
          activePersons.forEach(p => {
            if (!s.days[d].manual || !s.days[d].manual[p.id]) {
              syncMonthlySheetDayPersonHours(s, d, p.id);
            }
          });
        }
      } else {
        s.days[d].globalStart = '';
        s.days[d].globalEnd = '';
        activePersons.forEach(p => {
          if (!s.days[d].manual || !s.days[d].manual[p.id]) {
            delete s.days[d].hours[p.id];
          }
        });
      }
      
      Store.updateMonthlySheet(s.id, { days: s.days });
      renderSheetDetail(window.currentSheetId); // Force re-render to update inputs
    });
  });
}

function updateColumnSum(personId, activePersons) {
  const sheet = Store.getMonthlySheet(window.currentSheetId);
  if (!sheet) return;
  let s = 0;
  if (sheet.days) {
    Object.values(sheet.days).forEach(day => {
      if (day.hours && day.hours[personId]) {
        s += parseFloat(day.hours[personId]);
      }
    });
  }
  
  // Find index of person to update tfoot
  const idx = activePersons.findIndex(p => p.id === personId);
  if (idx !== -1) {
    const tfootRow = document.querySelector('#monthly-hours-foot tr:first-child');
    if (tfootRow && tfootRow.children[idx + 3]) { // +3 because V, Dzień, Czas
      const p = activePersons[idx];
      const state = Store.getState();
      const effectiveRate = Calculations.getSheetPersonRate(p, sheet, state);
      const earnings = s * effectiveRate;
      
      tfootRow.children[idx + 3].innerHTML = `
        <div style="color: var(--success); font-weight: bold; font-size: 1rem;">${s.toFixed(1)}h</div>
        <div style="color: var(--primary); font-size: 0.8rem; opacity: 0.9;">${earnings.toFixed(2)} zł</div>
      `;
    }
  }
}

// ==========================================
// EXPENSES VIEW
// ==========================================
const BONUS_EXPENSE_PAYER_ID = 'all_partners';
const EXPENSE_NAME_SUGGESTIONS = [
  'Paliwo',
  'Hotel',
  'Badania',
  'BHP',
  'Woda',
  'Kask',
  'Młotek',
  'Cęgi',
  'Narzędzia',
  'Rękawiczki',
  'Spodnie',
  'Buty',
  'Obiady',
  'Bilety',
  'Naprawa Auta',
  'Do Auta'
];

function renderExpenseNameSuggestions(filter = '') {
  const dropdown = document.getElementById('expense-name-suggestions');
  if (!dropdown) return;

  const normalizedFilter = (filter || '').toLocaleLowerCase('pl-PL').trim();
  const items = EXPENSE_NAME_SUGGESTIONS.filter(name =>
    !normalizedFilter || name.toLocaleLowerCase('pl-PL').includes(normalizedFilter)
  );

  if (items.length === 0) {
    dropdown.innerHTML = '<div class="input-dropdown-item" style="cursor: default; opacity: 0.7;">Brak podpowiedzi</div>';
    return;
  }

  dropdown.innerHTML = items.map(name => `
    <button type="button" class="input-dropdown-item expense-name-suggestion-item" data-value="${name}">${name}</button>
  `).join('');
}

function setExpenseSuggestionsVisibility(isVisible) {
  const dropdown = document.getElementById('expense-name-suggestions');
  if (!dropdown) return;
  dropdown.style.display = isVisible ? 'block' : 'none';
}

function initExpenseNameSuggestions() {
  const input = document.getElementById('expense-name');
  const button = document.getElementById('btn-expense-name-suggestions');
  const dropdown = document.getElementById('expense-name-suggestions');
  const container = document.getElementById('expense-name-container');
  if (!input || !button || !dropdown || !container) return;

  const openDropdown = () => {
    renderExpenseNameSuggestions(input.value);
    setExpenseSuggestionsVisibility(true);
  };

  const closeDropdown = () => {
    setExpenseSuggestionsVisibility(false);
  };

  button.addEventListener('click', (e) => {
    e.preventDefault();
    e.stopPropagation();

    const isOpen = dropdown.style.display === 'block';
    if (isOpen) {
      closeDropdown();
      return;
    }

    openDropdown();
  });

  input.addEventListener('input', () => {
    if (dropdown.style.display === 'block') {
      renderExpenseNameSuggestions(input.value);
    }
  });

  dropdown.addEventListener('click', (e) => {
    const item = e.target.closest('.expense-name-suggestion-item');
    if (!item) return;

    input.value = item.getAttribute('data-value') || '';
    closeDropdown();
    input.focus();
  });

  document.addEventListener('click', (e) => {
    if (!container.contains(e.target)) {
      closeDropdown();
    }
  });
}

function getExpenseRecipientOptionsHtml(type = 'ADVANCE', selectedRecipientId = '') {
  const recipients = (Store.getState().persons || []).filter(person => type !== 'BONUS' || person.type === 'EMPLOYEE');
  const placeholder = type === 'BONUS'
    ? '-- Wybierz pracownika --'
    : '-- Wybierz odbiorcę --';

  return `<option value="">${placeholder}</option>` + recipients.map(person => `
    <option value="${person.id}" ${person.id === selectedRecipientId ? 'selected' : ''}>${getPersonDisplayName(person)}</option>
  `).join('');
}

function populateExpenseRecipientSelect(type = 'ADVANCE', selectedRecipientId = '') {
  const selectAdvanceFor = document.getElementById('expense-advance-for');
  if (!selectAdvanceFor) return;

  selectAdvanceFor.innerHTML = getExpenseRecipientOptionsHtml(type, selectedRecipientId);
}

function updateExpenseFormTypeUI(type, selectedRecipientId = '') {
  const advanceForContainer = document.getElementById('expense-advance-for-container');
  const expenseNameContainer = document.getElementById('expense-name-container');
  const expenseNameInput = document.getElementById('expense-name');
  const paidBySelect = document.getElementById('expense-paid-by');
  const paidByContainer = paidBySelect ? paidBySelect.closest('.form-group') : null;
  const recipientLabel = document.getElementById('expense-recipient-label');
  const isAdvance = type === 'ADVANCE';
  const isBonus = type === 'BONUS';

  if (advanceForContainer) {
    advanceForContainer.style.display = (isAdvance || isBonus) ? 'block' : 'none';
  }

  if (expenseNameContainer && expenseNameInput) {
    expenseNameContainer.style.display = (isAdvance || isBonus) ? 'none' : 'block';
    expenseNameInput.required = !(isAdvance || isBonus);
    if (isBonus) expenseNameInput.value = 'Premia';
    if (isAdvance) expenseNameInput.value = 'Zaliczka';
  }

  if (paidByContainer && paidBySelect) {
    paidByContainer.style.display = isBonus ? 'none' : 'block';
    paidBySelect.required = !isBonus;
    if (isBonus) {
      paidBySelect.value = '';
    }
  }

  if (recipientLabel) {
    recipientLabel.textContent = isBonus ? 'Dla kogo Premia?' : 'Dla kogo zaliczka?';
  }

  populateExpenseRecipientSelect(type, selectedRecipientId);
}

function openExpenseFormForCreate() {
  const form = document.getElementById('expense-form');
  const typeSelect = document.getElementById('expense-type');
  if (!form || !typeSelect) return;

  document.getElementById('expense-form-title').textContent = 'Nowy Dowód/Koszt';
  document.getElementById('expense-id').value = '';
  document.getElementById('expense-date').value = getDefaultDateForSelectedMonth();
  document.getElementById('expense-name').value = '';
  document.getElementById('expense-amount').value = '';
  document.getElementById('expense-paid-by').value = '';
  document.getElementById('expense-advance-for').value = '';
  typeSelect.value = 'COST';
  updateExpenseFormTypeUI('COST');
  document.querySelector('#expenses-view .table-container').style.display = 'none';
  form.style.display = 'block';
  form.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function initExpensesTracker() {
  const form = document.getElementById('expense-form');
  const btnAdd = document.getElementById('btn-add-expense');
  const btnCancel = document.getElementById('btn-cancel-expense');
  const typeSelect = document.getElementById('expense-type');

  initExpenseNameSuggestions();

  document.getElementById('expense-date').value = getDefaultDateForSelectedMonth();

  btnAdd.addEventListener('click', openExpenseFormForCreate);

  btnCancel.addEventListener('click', () => {
    form.style.display = 'none';
    document.querySelector('#expenses-view .table-container').style.display = 'block';
  });

  typeSelect.addEventListener('change', (e) => {
    updateExpenseFormTypeUI(e.target.value);
  });

  form.addEventListener('submit', (e) => {
    e.preventDefault();
    const id = document.getElementById('expense-id').value;
    const type = typeSelect.value;
    const isAdvance = type === 'ADVANCE';
    const isBonus = type === 'BONUS';
    const date = document.getElementById('expense-date').value;
    const name = isAdvance ? 'Zaliczka' : (isBonus ? 'Premia' : document.getElementById('expense-name').value);
    const amount = parseFloat(document.getElementById('expense-amount').value);
    const paidById = isBonus ? BONUS_EXPENSE_PAYER_ID : document.getElementById('expense-paid-by').value;
    const advanceForId = (isAdvance || isBonus) ? document.getElementById('expense-advance-for').value : '';
    const selectedMonth = getSelectedMonthKey();

    if (!isBonus && !paidById) {
      alert('Wybierz osobę dokonującą płatności');
      return;
    }
    if ((isAdvance || isBonus) && !advanceForId) {
      alert(isBonus ? 'Wybierz pracownika, który otrzymuje premię.' : 'Wybierz osobę która otrzymuje zaliczkę');
      return;
    }
    if (!date || !date.startsWith(`${selectedMonth}-`)) {
      alert('Data wpisu musi należeć do wybranego miesiąca.');
      document.getElementById('expense-date').focus();
      return;
    }

    if (id) {
      Store.updateExpense(id, { type, date, name, amount, paidById, advanceForId });
    } else {
      Store.addExpense({ type, date, name, amount, paidById, advanceForId });
    }
    form.style.display = 'none';
    document.querySelector('#expenses-view .table-container').style.display = 'block';
  });

  updateExpenseFormTypeUI(typeSelect.value);
}

function renderExpenses() {
  const state = Store.getState();

  document.getElementById('expense-form').style.display = 'none';
  const tableContainer = document.querySelector('#expenses-view .table-container');
  if (tableContainer) tableContainer.style.display = 'block';

  const selectedMonth = getSelectedMonthKey();
  const selectPaidBy = document.getElementById('expense-paid-by');
  const currentExpenseType = document.getElementById('expense-type')?.value || 'COST';
  
  // Re-populate dropdowns
  const personsOptionsHtml = state.persons.map(p => `<option value="${p.id}">${getPersonDisplayName(p)}</option>`).join('');
  const clientsOptionsHtml = state.clients.map(c => `<option value="client_${c.id}">${c.name}</option>`).join('');
  
  selectPaidBy.innerHTML = `<option value="">-- Wybierz płatnika --</option>` + 
    `<optgroup label="Zespół">${personsOptionsHtml}</optgroup>` + 
    `<optgroup label="Klienci">${clientsOptionsHtml}</optgroup>`;
  populateExpenseRecipientSelect(currentExpenseType);

  const tbody = document.getElementById('expenses-table-body');
  tbody.innerHTML = '';

  const expensesForMonth = state.expenses.filter(expense => expense.date && expense.date.startsWith(`${selectedMonth}-`));

  if (expensesForMonth.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;color:var(--text-muted)">Brak kosztów, zaliczek i premii.</td></tr>`;
    return;
  }

  // Sort by date (ascending - oldest first)
  const sortedExpenses = [...expensesForMonth].sort((a, b) => new Date(a.date) - new Date(b.date));

  const getPayerName = (id) => {
    if (!id) return 'Nieznany';
    if (id === BONUS_EXPENSE_PAYER_ID) return 'Od Wszystkich Wspólników';
    if (id.startsWith('client_')) {
      const cId = id.substring(7);
      return (state.clients.find(c => c.id === cId)?.name || 'Nieznany') + ' (Klient)';
    }
    return getPersonDisplayName(state.persons.find(p => p.id === id)) || 'Nieznany';
  };

  const getPersonName = (id) => getPersonDisplayName(state.persons.find(p => p.id === id)) || 'Nieznany';

  const getExpenseBadgeMeta = (type) => {
    if (type === 'BONUS') {
      return { badgeClass: 'badge-bonus', label: 'Premia' };
    }
    if (type === 'ADVANCE') {
      return { badgeClass: 'badge-employee', label: 'Zaliczka' };
    }
    return { badgeClass: 'badge-partner', label: 'Koszt' };
  };

  sortedExpenses.forEach(e => {
    const badgeMeta = getExpenseBadgeMeta(e.type);
    const expenseTitle = e.type === 'ADVANCE' ? 'Zaliczka' : (e.type === 'BONUS' ? 'Premia' : e.name);
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${e.date}</td>
      <td style="font-weight: 500">${expenseTitle}</td>
      <td>
        <span class="badge ${badgeMeta.badgeClass}">
          ${badgeMeta.label}
        </span>
      </td>
      <td style="color: ${e.type === 'ADVANCE' ? 'var(--warning)' : (e.type === 'BONUS' ? '#fbbf24' : 'var(--danger)')}">
        ${e.amount.toFixed(2)} zł
      </td>
      <td>${e.type === 'BONUS' ? 'Od Wszystkich Wspólników' : getPayerName(e.paidById)}</td>
      <td>${e.type === 'ADVANCE' || e.type === 'BONUS' ? getPersonName(e.advanceForId) : '-'}</td>
      <td>
        <div style="display: flex; gap: 0.5rem;">
          <button class="btn btn-secondary btn-icon btn-edit-expense" data-id="${e.id}">
            <i data-lucide="edit-2" style="width:16px;height:16px"></i>
          </button>
          <button class="btn btn-danger btn-icon btn-delete-expense" data-id="${e.id}">
            <i data-lucide="trash-2" style="width:16px;height:16px"></i>
          </button>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  });

  lucide.createIcons();

  document.querySelectorAll('.btn-edit-expense').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = e.currentTarget.getAttribute('data-id');
      const expense = state.expenses.find(x => x.id === id);
      if (expense) {
        document.getElementById('expense-form-title').textContent = 'Edytuj Wpis';
        document.getElementById('expense-id').value = expense.id;
        document.getElementById('expense-date').value = expense.date;
        document.getElementById('expense-name').value = expense.name;
        document.getElementById('expense-amount').value = expense.amount;
        document.getElementById('expense-paid-by').value = expense.paidById;
        const typeSelect = document.getElementById('expense-type');
        typeSelect.value = expense.type;
        updateExpenseFormTypeUI(expense.type, expense.advanceForId || '');
        if (expense.type !== 'BONUS') {
          document.getElementById('expense-paid-by').value = expense.paidById;
        }
        
        document.querySelector('#expenses-view .table-container').style.display = 'none';
        document.getElementById('expense-form').style.display = 'block';
        document.getElementById('expense-form').scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    });
  });

  document.querySelectorAll('.btn-delete-expense').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = e.currentTarget.getAttribute('data-id');
      if (confirm('Czy na pewno chcesz usunąć ten wpis?')) {
        Store.deleteExpense(id);
      }
    });
  });
}

// ==========================================
// SETTLEMENT VIEW
// ==========================================
function formatSettlementCurrency(value) {
  const amount = parseFloat(value);
  return `${(Number.isFinite(amount) ? amount : 0).toFixed(2)} zł`;
}

function formatSettlementHours(value) {
  const amount = parseFloat(value);
  return `${(Number.isFinite(amount) ? amount : 0).toFixed(1)}h`;
}

function formatPercentConfigInputValue(decimalValue) {
  const percentValue = (parseFloat(decimalValue) || 0) * 100;
  return percentValue === 0
    ? '0'
    : percentValue.toFixed(3).replace(/\.0+$|(?<=\.[0-9]*[1-9])0+$/g, '');
}

function populateSettlementPersonContractGrid() {
  const state = Store.getState();
  const selectedMonth = getSelectedMonthKey();
  const monthSettings = Store.getMonthSettings(selectedMonth);
  const grid = document.getElementById('settlement-person-contract-grid');
  if (!grid) return;

  const persons = Calculations.getSettlementPersons(state, selectedMonth)
    .filter(person => person.type === 'EMPLOYEE' || person.type === 'WORKING_PARTNER');

  if (persons.length === 0) {
    grid.innerHTML = '<p style="color: var(--text-muted); grid-column: 1 / -1;">Brak pracowników lub wspólników pracujących do konfiguracji UZ w tym miesiącu.</p>';
    return;
  }

  grid.innerHTML = persons.map(person => {
    const monthOverride = monthSettings.personContractCharges?.[person.id] || {};
    const defaultTaxAmount = parseFloat(person.contractTaxAmount) || 0;
    const defaultZusAmount = parseFloat(person.contractZusAmount) || 0;
    const typeLabel = person.type === 'WORKING_PARTNER' ? 'Wspólnik pracujący' : 'Pracownik';

    return `
      <div class="glass-panel" data-person-id="${person.id}" style="padding: 0.9rem; border-radius: 12px;">
        <div style="font-weight: 600; margin-bottom: 0.25rem;">${getPersonDisplayName(person)}</div>
        <div style="color: var(--text-secondary); font-size: 0.8rem; margin-bottom: 0.75rem;">${typeLabel} • Domyślnie: Podatek UZ ${defaultTaxAmount.toFixed(2)} zł • ZUS UZ ${defaultZusAmount.toFixed(2)} zł</div>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 0.75rem;">
          <div class="form-group" style="margin-bottom: 0;">
            <label>Podatek UZ (zł)</label>
            <input type="number" step="0.01" class="config-month-contract-tax" data-person-id="${person.id}" value="${monthOverride.contractTaxAmount ?? ''}" placeholder="Domyślnie ${defaultTaxAmount.toFixed(2)}">
          </div>
          <div class="form-group" style="margin-bottom: 0;">
            <label>ZUS UZ (zł)</label>
            <input type="number" step="0.01" class="config-month-contract-zus" data-person-id="${person.id}" value="${monthOverride.contractZusAmount ?? ''}" placeholder="Domyślnie ${defaultZusAmount.toFixed(2)}">
          </div>
        </div>
      </div>
    `;
  }).join('');
}

function populateSettlementConfigPanel() {
  const state = Store.getState();
  const selectedMonth = getSelectedMonthKey();
  const monthSettings = Store.getMonthSettings(selectedMonth);
  const taxRatePercent = parseFloat(state.config.taxRate) || 0;
  const monthTaxRate = parseFloat(monthSettings.settlementConfig?.taxRate);
  const monthZusFixedAmount = parseFloat(monthSettings.settlementConfig?.zusFixedAmount);
  const monthInfo = document.getElementById('config-month-info');

  document.getElementById('config-tax').value = formatPercentConfigInputValue(taxRatePercent);
  document.getElementById('config-zus').value = parseFloat(state.config.zusFixedAmount) || 0;
  document.getElementById('config-month-tax').value = Number.isFinite(monthTaxRate)
    ? formatPercentConfigInputValue(monthTaxRate)
    : '';
  document.getElementById('config-month-zus').value = Number.isFinite(monthZusFixedAmount)
    ? monthZusFixedAmount
    : '';

  if (monthInfo) {
    monthInfo.textContent = `Nadpisz domyślne stawki tylko dla miesiąca ${formatMonthLabel(selectedMonth)}.`;
  }

  populateSettlementPersonContractGrid();
}

function getEmployeeGeneratedProfitDisplay(entry) {
  const generatedProfit = Number.isFinite(parseFloat(entry?.generatedProfit)) ? parseFloat(entry.generatedProfit) : 0;
  const employerCoveredUz = parseFloat(entry?.employerPaidContractCharges) || 0;

  if (employerCoveredUz <= 0) {
    return {
      label: 'Wypracowany zysk z godzin:',
      valueText: formatSettlementCurrency(generatedProfit)
    };
  }

  const profitAfterUz = generatedProfit - employerCoveredUz;

  return {
    label: 'Wypracowany zysk z godzin (bez UZ):',
    valueText: `${formatSettlementCurrency(generatedProfit)} (${formatSettlementCurrency(profitAfterUz)})`
  };
}

function getSettlementEntityName(state, id) {
  if (!id) return 'Nieznany';
  if (id.startsWith('client_')) {
    const clientId = id.substring(7);
    const client = (state.clients || []).find(c => c.id === clientId);
    return client ? `${client.name} (Klient)` : 'Nieznany klient';
  }

  return getPersonDisplayName((state.persons || []).find(person => person.id === id)) || 'Nieznana osoba';
}

function getSettlementSheetLabel(client, site) {
  if (client && site) return `${client} / ${site}`;
  return client || site || 'Bez przypisania';
}

function getSettlementPersonTypeLabel(type) {
  if (type === 'PARTNER') return 'Wspólnik';
  if (type === 'SEPARATE_COMPANY') return 'Osobna firma';
  if (type === 'WORKING_PARTNER') return 'Wspólnik pracujący';
  return 'Pracownik';
}

function getSettlementPersonDisplayName(person, state) {
  if (!person) return 'Nieznana osoba';
  const employer = person?.employerId
    ? (state?.persons || []).find(candidate => candidate.id === person.employerId)
    : null;
  const separateCompanySuffix = person?.type === 'EMPLOYEE' && Calculations.isSeparateCompany(employer)
    ? ' (z osobnej firmy)'
    : '';
  const baseName = `${getPersonDisplayName(person)}${separateCompanySuffix}`;
  return Calculations.isPersonActiveInMonth(person, state)
    ? baseName
    : `${baseName} (nieaktywny)`;
}

function getSettlementPersonDisplayNameHtml(person, state) {
  if (!person) return 'Nieznana osoba';

  const employer = person?.employerId
    ? (state?.persons || []).find(candidate => candidate.id === person.employerId)
    : null;
  const suffixes = [];

  if (person?.type === 'EMPLOYEE' && Calculations.isSeparateCompany(employer)) {
    suffixes.push('(z osobnej firmy)');
  }

  if (!Calculations.isPersonActiveInMonth(person, state)) {
    suffixes.push('(nieaktywny)');
  }

  if (suffixes.length === 0) {
    return getPersonDisplayName(person);
  }

  return `${getPersonDisplayName(person)} <span style="font-weight:400; color: var(--text-secondary);">${suffixes.join(' ')}</span>`;
}

function getSettlementPersonSources(person, state) {
  const selectedMonth = Calculations.getSelectedMonth(state);
  const sources = [];
  let totalHours = 0;
  let totalSalary = 0;
  const includeOnlyRoboczogodzinyFromWorks = person.type === 'PARTNER' || person.type === 'WORKING_PARTNER';

  (state.monthlySheets || []).forEach(sheet => {
    if (sheet.month !== selectedMonth) return;
    if (Array.isArray(sheet.activePersons) && !sheet.activePersons.includes(person.id)) return;

    let sheetHours = 0;
    Object.entries(sheet.days || {}).forEach(([dayKey, day]) => {
      if (isMonthlySheetPersonInactiveOnDay(sheet, person.id, parseInt(dayKey, 10))) return;
      if (day.hours && day.hours[person.id] !== undefined && day.hours[person.id] !== '') {
        const hours = parseFloat(day.hours[person.id]);
        if (Number.isFinite(hours)) sheetHours += hours;
      }
    });

    if (sheetHours <= 0) return;

    const rate = Calculations.getSheetPersonRate(person, sheet, state, 0);
    const salary = sheetHours * rate;
    totalHours += sheetHours;
    totalSalary += salary;

    sources.push({
      label: `Arkusz godzin: ${getSettlementSheetLabel(sheet.client, sheet.site)}`,
      hours: sheetHours,
      rate,
      salary
    });
  });

  (state.worksSheets || []).forEach(sheet => {
    if (sheet.month !== selectedMonth || !(sheet.activePersons || []).includes(person.id)) return;

    let sheetHours = 0;
    (sheet.entries || []).forEach(entry => {
      if (includeOnlyRoboczogodzinyFromWorks && !Calculations.isRoboczogodzinyEntry(entry, state)) return;
      if (entry.hours && entry.hours[person.id] !== undefined && entry.hours[person.id] !== '') {
        const hours = parseFloat(entry.hours[person.id]);
        if (Number.isFinite(hours)) sheetHours += hours;
      }
    });

    if (sheetHours <= 0) return;

    const rate = Calculations.getSheetPersonRate(person, sheet, state, 0);
    const salary = sheetHours * rate;
    totalHours += sheetHours;
    totalSalary += salary;

    sources.push({
      label: `Arkusz prac: ${getSettlementSheetLabel(sheet.client, sheet.site)}`,
      hours: sheetHours,
      rate,
      salary
    });
  });

  return { sources, totalHours, totalSalary };
}

function getSettlementPersonExpenseDetails(personId, selectedExpenses, state) {
  const paidCosts = selectedExpenses
    .filter(expense => expense.type === 'COST' && expense.paidById === personId)
    .map(expense => ({
      label: `${expense.date} • ${expense.name || 'Koszt'}`,
      amount: expense.amount,
      meta: 'Koszt zapłacony przez tę osobę'
    }));

  const paidAdvances = selectedExpenses
    .filter(expense => expense.type === 'ADVANCE' && expense.paidById === personId)
    .map(expense => ({
      label: `${expense.date} • Zaliczka dla ${getSettlementEntityName(state, expense.advanceForId)}`,
      amount: expense.amount,
      meta: 'Zwrot środków wypłaconych przez tę osobę'
    }));

  const advancesTaken = selectedExpenses
    .filter(expense => expense.type === 'ADVANCE' && expense.advanceForId === personId)
    .map(expense => ({
      label: `${expense.date} • Zaliczka od ${getSettlementEntityName(state, expense.paidById)}`,
      amount: expense.amount,
      meta: 'Kwota już pobrana wcześniej'
    }));

  const bonusesReceived = selectedExpenses
    .filter(expense => expense.type === 'BONUS' && expense.advanceForId === personId)
    .map(expense => ({
      label: `${expense.date} • Premia`,
      amount: expense.amount,
      meta: 'Premia od wszystkich wspólników'
    }));

  return { paidCosts, paidAdvances, advancesTaken, bonusesReceived };
}

function getWorksProfitSplitDetails(sheet, profit, state, activePersonIds) {
  const allPartners = (state.persons || []).filter(person =>
    (person.type === 'PARTNER' || person.type === 'WORKING_PARTNER' || person.type === 'SEPARATE_COMPANY')
    && activePersonIds.has(person.id)
  );

  const eligiblePartners = allPartners.filter(person => {
    const isParticipant = sheet.activePersons && sheet.activePersons.includes(person.id);
    const hasOverride = person.type === 'PARTNER'
      && sheet.partnerProfitOverrides
      && sheet.partnerProfitOverrides.includes(person.id);
    return isParticipant || hasOverride;
  });

  const shareGroup = eligiblePartners.length > 0 ? eligiblePartners : allPartners;
  if (shareGroup.length === 0) {
    return [];
  }

  const shareAmount = profit / shareGroup.length;
  return shareGroup.map(person => ({
    name: getPersonDisplayName(person),
    typeLabel: getSettlementPersonTypeLabel(person.type),
    amount: shareAmount
  }));
}

function getSettlementDetailsData(state, res) {
  const selectedMonth = Calculations.getSelectedMonth(state);
  const settlementPersons = Calculations.getSettlementPersons(state, selectedMonth);
  const activePersons = Calculations.getActivePersons(state, selectedMonth);
  const settlementPersonIds = new Set(settlementPersons.map(person => person.id));
  const settlementEmployeeIds = new Set(settlementPersons.filter(person => person.type === 'EMPLOYEE').map(person => person.id));
  const activePersonIds = new Set(activePersons.map(person => person.id));
  const activeCostShareCount = activePersons.filter(person =>
    (person.type === 'PARTNER' || person.type === 'WORKING_PARTNER')
    && Calculations.personParticipatesInCosts(person)
  ).length;
  const selectedExpenses = (state.expenses || []).filter(expense => expense.date && expense.date.startsWith(`${selectedMonth}-`));

  const hourlyRevenueItems = (state.monthlySheets || [])
    .filter(sheet => sheet.month === selectedMonth)
    .map(sheet => {
      const clientRate = Calculations.getSheetClientRate(sheet, state);
      const totalHours = Calculations.getSheetTotalHours(sheet, settlementPersonIds);
      const employeeHours = Calculations.getSheetTotalHours(sheet, settlementEmployeeIds);
      return {
        label: getSettlementSheetLabel(sheet.client, sheet.site),
        hours: totalHours,
        employeeHours,
        clientRate,
        revenue: totalHours * clientRate
      };
    })
    .filter(item => item.hours > 0 || item.revenue > 0);

  const worksRevenueItems = (state.worksSheets || [])
    .filter(sheet => sheet.month === selectedMonth)
    .map(sheet => {
      let revenue = 0;
      let employeeCost = 0;
      let profitBaseRevenue = 0;

      const entries = (sheet.entries || []).map(entry => {
        const metrics = Calculations.getWorksEntryMetrics(entry, sheet, state);

        revenue += metrics.revenue;
        employeeCost += metrics.employeeCost;
        profitBaseRevenue += metrics.profitBaseRevenue;

        return {
          label: `${entry.date || 'Brak daty'} • ${entry.name || 'Pozycja'}`,
          description: metrics.isRoboczogodziny
            ? `${formatSettlementHours(metrics.totalLoggedHours)} × ${formatSettlementCurrency(metrics.clientRate).replace(' zł', ' zł/h')}`
            : `${(parseFloat(entry.quantity) || 0).toFixed(2)} ${entry.unit || ''} × ${formatSettlementCurrency(parseFloat(entry.price) || 0)}`,
          revenue: metrics.revenue,
          employeeCost: metrics.employeeCost,
          profitBaseRevenue: metrics.profitBaseRevenue,
          profitBaseLabel: metrics.isRoboczogodziny ? 'Przychód z roboczogodzin pracowników' : 'Wartość do podziału',
          profit: metrics.profit
        };
      }).filter(entry => entry.revenue !== 0 || entry.employeeCost !== 0);

      return {
        label: getSettlementSheetLabel(sheet.client, sheet.site),
        revenue,
        employeeCost,
        profitBaseRevenue,
        profit: profitBaseRevenue - employeeCost,
        profitSplitItems: getWorksProfitSplitDetails(sheet, profitBaseRevenue - employeeCost, state, activePersonIds),
        entries
      };
    })
    .filter(item => item.entries.length > 0 || item.revenue !== 0 || item.employeeCost !== 0);

  const totalEmployeesSalary = res.employees.reduce((sum, employee) => sum + employee.salary, 0);
  const totalEmployeeBonuses = res.employees.reduce((sum, employee) => sum + (employee.bonusAmount || 0), 0);
  const totalCommonCosts = selectedExpenses
    .filter(expense => expense.type === 'COST')
    .reduce((sum, expense) => sum + expense.amount, 0);
  const clientAdvanceItems = selectedExpenses.filter(expense => expense.type === 'ADVANCE' && expense.paidById && expense.paidById.startsWith('client_'));
  const otherAdvanceItems = selectedExpenses.filter(expense => expense.type === 'ADVANCE' && (!expense.paidById || !expense.paidById.startsWith('client_')));
  const bonusItems = selectedExpenses.filter(expense => expense.type === 'BONUS');

  return {
    selectedMonth,
    hourlyRevenueItems,
    worksRevenueItems,
    totalHourlyRevenue: hourlyRevenueItems.reduce((sum, item) => sum + item.revenue, 0),
    totalWorksRevenue: worksRevenueItems.reduce((sum, item) => sum + item.revenue, 0),
    totalWorksProfitBaseRevenue: worksRevenueItems.reduce((sum, item) => sum + item.profitBaseRevenue, 0),
    totalWorksEmployeeCost: worksRevenueItems.reduce((sum, item) => sum + item.employeeCost, 0),
    totalEmployeesSalary,
    totalEmployeesSalaryFromHours: res.employeeSalaryFromHours || 0,
    totalEmployeeBonuses,
    totalGrossPayouts: [...res.partners, ...(res.separateCompanies || []), ...res.workingPartners, ...res.employees].reduce((sum, person) => sum + (person.toPayout || 0), 0),
    totalCommonCosts,
    activeCostShareCount,
    costShare: activeCostShareCount > 0 ? totalCommonCosts / activeCostShareCount : 0,
    clientAdvanceItems,
    otherAdvanceItems,
    bonusItems,
    selectedExpenses
  };
}

function buildSettlementEntryListHtml(items, valueKey, emptyMessage, metaKey = 'meta') {
  if (!items.length) {
    return `<p class="settlement-detail-empty">${emptyMessage}</p>`;
  }

  return items.map(item => `
    <div class="settlement-detail-item">
      <div class="settlement-detail-item-header">
        <div>
          <h4>${item.label}</h4>
          ${item[metaKey] ? `<div class="settlement-detail-meta">${item[metaKey]}</div>` : ''}
        </div>
        <strong>${formatSettlementCurrency(item[valueKey])}</strong>
      </div>
      ${item.description ? `<div class="settlement-detail-meta">${item.description}</div>` : ''}
      ${item.employeeCost > 0 ? `<div class="settlement-detail-row" style="margin-top: 0.45rem;"><span>Koszt pracowników w tej pozycji</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(item.employeeCost)}</strong></div>` : ''}
      ${item.profit !== undefined ? `<div class="settlement-detail-row" style="margin-top: 0.25rem;"><span>Zysk po kosztach pracowników</span><strong class="settlement-accent-primary">${formatSettlementCurrency(item.profit)}</strong></div>` : ''}
    </div>
  `).join('');
}

function buildSettlementProfitRecipientSummary(items) {
  if (!items || items.length === 0) return 'Brak udziału osobnych firm.';
  return items
    .map(item => `${item.name}: ${item.percent.toFixed(2)}%`)
    .join(' • ');
}

function buildSettlementProfitAllocationSummary(res) {
  const recipients = [
    ...(res?.partners || []).map(item => ({
      name: getPersonDisplayName(item.person) || 'Wspólnik',
      typeLabel: 'Wspólnik',
      amount: parseFloat(item.revenueShare) || 0
    })),
    ...((res?.separateCompanies || []).map(item => ({
      name: getPersonDisplayName(item.person) || 'Osobna firma',
      typeLabel: 'Osobna firma',
      amount: parseFloat(item.revenueShare) || 0
    })))
  ].filter(item => Math.abs(item.amount) >= 0.005);

  if (!recipients.length) {
    return '<p class="settlement-detail-empty">Brak przydzielonego zysku dla wspólników i osobnych firm.</p>';
  }

  return recipients.map(item => `
    <div class="settlement-detail-item">
      <div class="settlement-detail-item-header">
        <div>
          <h4>${item.name}</h4>
          <div class="settlement-detail-meta">${item.typeLabel}</div>
        </div>
        <strong>${formatSettlementCurrency(item.amount)}</strong>
      </div>
    </div>
  `).join('');
}

function buildSettlementProfitBreakdownSectionHtml(res, taxDetailsHtml = '') {
  const breakdown = res.profitBreakdown || {};
  const ownEmployees = breakdown.ownEmployees || [];
  const separateCompanyEmployees = breakdown.separateCompanyEmployees || [];
  const separateCompanyRateDifferences = breakdown.separateCompanyRateDifferences || [];
  const recipientSummary = buildSettlementProfitRecipientSummary(breakdown.separateCompanyRecipientPercents || []);

  const ownEmployeesHtml = ownEmployees.length
    ? ownEmployees.map(item => `
      <div class="settlement-detail-item">
        <div class="settlement-detail-item-header">
          <div>
            <h4>${item.employeeName}</h4>
            <div class="settlement-detail-meta">Udział osobnych firm w tym zysku: ${recipientSummary}</div>
          </div>
          <strong>${formatSettlementCurrency(item.profit)}</strong>
        </div>
      </div>
    `).join('')
    : '<p class="settlement-detail-empty">Brak zysków z naszych pracowników w wybranym miesiącu.</p>';

  const separateCompanyEmployeesHtml = separateCompanyEmployees.length
    ? separateCompanyEmployees.map(item => `
      <div class="settlement-detail-item">
        <div class="settlement-detail-item-header">
          <div>
            <h4>${item.employeeName}</h4>
            <div class="settlement-detail-meta">Firma: ${item.companyName}</div>
          </div>
          <strong>${formatSettlementCurrency(item.profit)}</strong>
        </div>
        <div class="settlement-detail-row" style="margin-top: 0.45rem;"><span>Do podziału</span><strong>${item.sharePercent.toFixed(2)}% = ${formatSettlementCurrency(item.distributedProfit)}</strong></div>
        ${Math.abs(item.retainedProfit || 0) >= 0.005 ? `<div class="settlement-detail-row" style="margin-top: 0.25rem;"><span>Zostaje w osobnej firmie</span><strong>${formatSettlementCurrency(item.retainedProfit || 0)}</strong></div>` : ''}
      </div>
    `).join('')
    : '<p class="settlement-detail-empty">Brak zysków z pracowników osobnych firm.</p>';

  const rateDifferencesHtml = separateCompanyRateDifferences.length
    ? separateCompanyRateDifferences.map(item => `
      <div class="settlement-detail-item">
        <div class="settlement-detail-item-header">
          <div>
            <h4>${item.sourceName}</h4>
            <div class="settlement-detail-meta">Firma: ${item.companyName}</div>
          </div>
          <strong>${formatSettlementCurrency(item.profit)}</strong>
        </div>
      </div>
    `).join('')
    : '<p class="settlement-detail-empty">Brak zysków z różnicy stawek firmy.</p>';

  return `
    <div class="settlement-detail-section">
      <h3>Składniki zarobku do podziału</h3>
      <div class="settlement-details-grid">
        <div>
          <h4 style="margin-bottom: 0.75rem;">Zyski z naszych pracowników</h4>
          <div class="settlement-detail-list">${ownEmployeesHtml}</div>
          <div class="settlement-detail-formula" style="margin-top: 0.75rem;">Łącznie: <strong>${formatSettlementCurrency(breakdown.ownEmployeesTotal || 0)}</strong></div>
        </div>
        <div>
          <h4 style="margin-bottom: 0.75rem;">Zyski z pracowników osobnej firmy</h4>
          <div class="settlement-detail-list">${separateCompanyEmployeesHtml}</div>
          <div class="settlement-detail-formula" style="margin-top: 0.75rem;">Łącznie do podziału: <strong>${formatSettlementCurrency(breakdown.separateCompanyEmployeesTotal || 0)}</strong></div>
        </div>
        <div>
          <h4 style="margin-bottom: 0.75rem;">Zyski z różnicy stawki firmy</h4>
          <div class="settlement-detail-list">${rateDifferencesHtml}</div>
          <div class="settlement-detail-formula" style="margin-top: 0.75rem;">Łącznie: <strong>${formatSettlementCurrency(breakdown.separateCompanyRateDifferencesTotal || 0)}</strong></div>
        </div>
      </div>
      <div class="settlement-detail-formula" style="margin-top: 1rem;">${formatSettlementCurrency(breakdown.ownEmployeesTotal || 0)} + ${formatSettlementCurrency(breakdown.separateCompanyEmployeesTotal || 0)} + ${formatSettlementCurrency(breakdown.separateCompanyRateDifferencesTotal || 0)} + ${formatSettlementCurrency(res.totalWorksProfit || 0)} = <strong>${formatSettlementCurrency(res.profitToSplit || 0)}</strong></div>
    </div>
    <div class="settlement-detail-section">
      <h3>Kto ile dostał z zysku</h3>
      <div class="settlement-detail-list">${buildSettlementProfitAllocationSummary(res)}</div>
    </div>
    ${taxDetailsHtml}
  `;
}

function buildSettlementPersonCardHtml(entry, role, state, details) {
  const personSources = getSettlementPersonSources(entry.person, state);
  const expenseDetails = getSettlementPersonExpenseDetails(entry.person.id, details.selectedExpenses, state);
  const employeeGeneratedProfit = role === 'employee' ? getEmployeeGeneratedProfitDisplay(entry) : null;
  const ownGrossAmount = parseFloat(entry.ownGrossAmount) || 0;
  const totalTaxAmount = parseFloat(entry.taxAmount) || 0;
  const sharedCompanyTaxAmount = parseFloat(entry.sharedCompanyTaxAmount) || 0;
  const bonusAmount = parseFloat(entry.bonusAmount) || 0;
  const sourcesHtml = personSources.sources.length > 0
    ? personSources.sources.map(source => `<li>${source.label}: ${formatSettlementHours(source.hours)} × ${formatSettlementCurrency(source.rate).replace(' zł', ' zł/h')} = ${formatSettlementCurrency(source.salary)}</li>`).join('')
    : '<li>Brak godzin w wybranym miesiącu.</li>';
  const contractTaxLabel = entry.contractChargesPaidByEmployer ? 'Podatek umowy zlecenie (opłaca pracodawca)' : 'Podatek umowy zlecenie';
  const contractZusLabel = entry.contractChargesPaidByEmployer ? 'ZUS umowy zlecenie (opłaca pracodawca)' : 'ZUS umowy zlecenie';

  let ownGrossFormula = `${formatSettlementCurrency(entry.salary)}`;
  let ownGrossText = 'Przychód własny(Brutto) = zarobek z własnych godzin';
  let grossFormula = `${formatSettlementCurrency(ownGrossAmount)}`;
  let grossText = 'Przychód (Brutto) = przychód własny(Brutto)';
  let netFormulaHtml = '';

  if (role === 'partner') {
    ownGrossFormula = `${formatSettlementCurrency(entry.salary)} + ${formatSettlementCurrency(entry.revenueShare)} + ${formatSettlementCurrency(entry.worksShare)}`;
    ownGrossText = 'Przychód własny(Brutto) = własne godziny + udział z godzin pracowników + udział z wykonanych prac';
    grossFormula += ` + ${formatSettlementCurrency(entry.paidCosts)} + ${formatSettlementCurrency(entry.paidAdvances)} - ${formatSettlementCurrency(entry.costShareApplied || 0)} - ${formatSettlementCurrency(entry.advancesTaken)}`;
    grossText = 'Przychód (Brutto) = przychód własny(Brutto) + zwrot kosztów + zwrot zaliczek - udział w kosztach - pobrane zaliczki';
    netFormulaHtml = `<div style="margin-top: 0.45rem;">Netto = ${formatSettlementCurrency(entry.toPayout)} - podatek własny ${formatSettlementCurrency(entry.ownTaxAmount || 0)}${sharedCompanyTaxAmount > 0 ? ` - podatek wspólny ${formatSettlementCurrency(sharedCompanyTaxAmount)}` : ''} - ZUS ${formatSettlementCurrency(entry.zusAmount)} = <strong>${formatSettlementCurrency(entry.netAfterAccounting)}</strong></div>`;
  } else if (role === 'separateCompany') {
    ownGrossFormula = `${formatSettlementCurrency(entry.salary)} + ${formatSettlementCurrency(entry.revenueShare)} + ${formatSettlementCurrency(entry.worksShare)}`;
    ownGrossText = 'Przychód własny(Brutto) = własne godziny + udział z godzin pracowników + udział z wykonanych prac';
    grossFormula += ` + ${formatSettlementCurrency(entry.paidCosts)} + ${formatSettlementCurrency(entry.paidAdvances)} - ${formatSettlementCurrency(entry.costShareApplied || 0)} - ${formatSettlementCurrency(entry.advancesTaken)}`;
    grossText = 'Przychód (Brutto) = przychód własny(Brutto) + zwrot kosztów + zwrot zaliczek - udział w kosztach - pobrane zaliczki';
  } else if (role === 'workingPartner') {
    ownGrossFormula = `${formatSettlementCurrency(entry.salary)} + ${formatSettlementCurrency(entry.worksShare)}`;
    ownGrossText = 'Przychód własny(Brutto) = własne godziny + udział z wykonanych prac';
    grossFormula += ` + ${formatSettlementCurrency(entry.paidCosts)} + ${formatSettlementCurrency(entry.paidAdvances)} - ${formatSettlementCurrency(entry.costShareApplied || 0)} - ${formatSettlementCurrency(entry.advancesTaken)}`;
    grossText = 'Przychód (Brutto) = przychód własny(Brutto) + zwrot kosztów + zwrot zaliczek - udział w kosztach - pobrane zaliczki';
    const deductions = [];
    if ((entry.ownTaxAmount || 0) > 0) deductions.push(`- podatek własny ${formatSettlementCurrency(entry.ownTaxAmount || 0)}`);
    if ((entry.deductedContractTaxAmount || 0) > 0) deductions.push(`- podatek UZ ${formatSettlementCurrency(entry.deductedContractTaxAmount || 0)}`);
    if ((entry.deductedContractZusAmount || 0) > 0) deductions.push(`- ZUS UZ ${formatSettlementCurrency(entry.deductedContractZusAmount || 0)}`);
    netFormulaHtml = `<div style="margin-top: 0.45rem;">Netto = ${formatSettlementCurrency(entry.toPayout)}${deductions.length ? ` ${deductions.join(' ')}` : ''} = <strong>${formatSettlementCurrency(entry.netAfterAccounting)}</strong></div>`;
  } else {
    if (bonusAmount > 0) {
      grossFormula += ` + ${formatSettlementCurrency(bonusAmount)}`;
    }
    grossFormula += ` + ${formatSettlementCurrency(entry.paidCosts)} + ${formatSettlementCurrency(entry.paidAdvances)} - ${formatSettlementCurrency(entry.advancesTaken)}`;
    if ((entry.deductedContractTaxAmount || 0) > 0 || (entry.deductedContractZusAmount || 0) > 0) {
      grossFormula += `${(entry.deductedContractTaxAmount || 0) > 0 ? ` - ${formatSettlementCurrency(entry.deductedContractTaxAmount || 0)}` : ''}${(entry.deductedContractZusAmount || 0) > 0 ? ` - ${formatSettlementCurrency(entry.deductedContractZusAmount || 0)}` : ''}`;
      grossText = `Do wypłaty = zarobek${bonusAmount > 0 ? ' + premia' : ''} + zwroty - pobrane zaliczki - podatek UZ - ZUS UZ`;
    } else {
      grossText = `Do wypłaty = zarobek${bonusAmount > 0 ? ' + premia' : ''} + zwroty - pobrane zaliczki`;
    }
  }

  return `
    <div class="settlement-person-card">
      <div class="settlement-person-header">
        <div>
          <div class="settlement-person-group">${getSettlementPersonTypeLabel(entry.person.type)}</div>
          <h4>${getSettlementPersonDisplayNameHtml(entry.person, state)}</h4>
          <div class="settlement-person-summary">Godziny: ${formatSettlementHours(entry.hours)}${entry.effectiveRate > 0 ? ` • efektywna stawka ${formatSettlementCurrency(entry.effectiveRate).replace(' zł', ' zł/h')}` : ''}</div>
        </div>
        <strong class="${(entry.netAfterAccounting ?? entry.toPayout) >= 0 ? 'settlement-accent-positive' : 'settlement-accent-negative'}">${role === 'employee' ? `Do wypłaty ${formatSettlementCurrency(entry.toPayout)}` : (role === 'separateCompany' ? `Brutto ${formatSettlementCurrency(entry.toPayout)}` : `Netto ${formatSettlementCurrency(entry.netAfterAccounting)}`)}</strong>
      </div>

      <div class="settlement-detail-stack">
        <div class="settlement-detail-row"><span>Zarobek z godzin</span><strong>${formatSettlementCurrency(entry.salary)}</strong></div>
        ${(role === 'partner' || role === 'separateCompany') ? `<div class="settlement-detail-row"><span>Udział z godzin pracowników</span><strong>${formatSettlementCurrency(entry.revenueShare)}</strong></div>` : ''}
        ${(role === 'partner' || role === 'workingPartner') ? `<div class="settlement-detail-row"><span>Udział z wykonanych prac</span><strong>${formatSettlementCurrency(entry.worksShare || 0)}</strong></div>` : ''}
        ${role === 'separateCompany' ? `<div class="settlement-detail-row"><span>Udział z wykonanych prac</span><strong>${formatSettlementCurrency(entry.worksShare || 0)}</strong></div>` : ''}
        ${(role === 'partner' || role === 'workingPartner' || role === 'separateCompany') ? `<div class="settlement-detail-row"><span>Przychód własny(Brutto)</span><strong>${formatSettlementCurrency(ownGrossAmount)}</strong></div>` : ''}
        ${(role === 'partner' || role === 'workingPartner') ? `<div class="settlement-detail-row"><span>Pensje pracowników</span><strong>${formatSettlementCurrency(entry.employeeSalaries || 0)}</strong></div>` : ''}
        ${role === 'separateCompany' ? `<div class="settlement-detail-row"><span>Pensje pracowników</span><strong>${formatSettlementCurrency(entry.employeeSalaries || 0)}</strong></div>` : ''}
        ${(role === 'partner' || role === 'workingPartner') && (entry.employeeReceivables || 0) > 0 ? `<div class="settlement-detail-row"><span>Do odebrania</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(entry.employeeReceivables || 0)}</strong></div>` : ''}
        ${role === 'separateCompany' && (entry.employeeReceivables || 0) > 0 ? `<div class="settlement-detail-row"><span>Do odebrania</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(entry.employeeReceivables || 0)}</strong></div>` : ''}
        ${(role === 'partner' || role === 'workingPartner') && (entry.employeeAccountingRefund || 0) > 0 ? `<div class="settlement-detail-row"><span>Zwrot Podatku i ZUS za pracowników</span><strong>${formatSettlementCurrency(entry.employeeAccountingRefund || 0)}</strong></div>` : ''}
        ${role === 'separateCompany' && (entry.employeeAccountingRefund || 0) > 0 ? `<div class="settlement-detail-row"><span>Zwrot Podatku i ZUS za pracowników</span><strong>${formatSettlementCurrency(entry.employeeAccountingRefund || 0)}</strong></div>` : ''}
        ${(role === 'partner' || role === 'workingPartner') && (entry.ownTaxAmount || 0) > 0 ? `<div class="settlement-detail-row"><span>Podatek własny</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(entry.ownTaxAmount || 0)}</strong></div>` : ''}
        ${(role === 'partner' || role === 'workingPartner') && sharedCompanyTaxAmount > 0 ? `<div class="settlement-detail-row"><span>Podatek wspólny</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(sharedCompanyTaxAmount)}</strong></div>` : ''}
        ${(role === 'employee' || role === 'workingPartner') && (entry.contractTaxAmount || 0) > 0 ? `<div class="settlement-detail-row"><span>${contractTaxLabel}</span><strong class="${entry.contractChargesPaidByEmployer ? '' : 'settlement-accent-negative'}">${entry.contractChargesPaidByEmployer ? formatSettlementCurrency(entry.contractTaxAmount || 0) : `-${formatSettlementCurrency(entry.contractTaxAmount || 0)}`}</strong></div>` : ''}
        ${(role === 'employee' || role === 'workingPartner') && (entry.contractZusAmount || 0) > 0 ? `<div class="settlement-detail-row"><span>${contractZusLabel}</span><strong class="${entry.contractChargesPaidByEmployer ? '' : 'settlement-accent-negative'}">${entry.contractChargesPaidByEmployer ? formatSettlementCurrency(entry.contractZusAmount || 0) : `-${formatSettlementCurrency(entry.contractZusAmount || 0)}`}</strong></div>` : ''}
        ${role === 'employee' ? `<div class="settlement-detail-row"><span>${employeeGeneratedProfit.label}</span><strong style="color: #2563eb;">${employeeGeneratedProfit.valueText}</strong></div>` : ''}
        ${bonusAmount > 0 ? `<div class="settlement-detail-row"><span>Premia</span><strong style="color: #fbbf24;">${formatSettlementCurrency(bonusAmount)}</strong></div>` : ''}
        <div class="settlement-detail-row"><span>Zwrot kosztów</span><strong>${formatSettlementCurrency(entry.paidCosts)}</strong></div>
        <div class="settlement-detail-row"><span>Zwrot wypłaconych zaliczek</span><strong>${formatSettlementCurrency(entry.paidAdvances)}</strong></div>
        ${(role === 'partner' || role === 'workingPartner') && (entry.costShareApplied || 0) > 0 ? `<div class="settlement-detail-row"><span>Udział w kosztach wspólnych</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(entry.costShareApplied || 0)}</strong></div>` : ''}
        ${role === 'separateCompany' && (entry.costShareApplied || 0) > 0 ? `<div class="settlement-detail-row"><span>Udział w kosztach wspólnych</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(entry.costShareApplied || 0)}</strong></div>` : ''}
        <div class="settlement-detail-row"><span>Pobrane zaliczki</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(entry.advancesTaken)}</strong></div>
      </div>

      <div class="settlement-detail-formula">
        ${(role === 'partner' || role === 'workingPartner' || role === 'separateCompany') ? `<div>${ownGrossText}</div><strong>${ownGrossFormula} = ${formatSettlementCurrency(ownGrossAmount)}</strong><div style="margin-top: 0.45rem;">${grossText}</div>` : `<div>${grossText}</div>`}
        <strong>${grossFormula} = ${formatSettlementCurrency(entry.toPayout)}</strong>
        ${role !== 'employee' ? `${((entry.employeeSalaries || 0) > 0 || (entry.employeeReceivables || 0) > 0 || (entry.employeeAccountingRefund || 0) > 0) ? `<div style="margin-top: 0.45rem;">Przychód (Brutto) z Pensjami = ${formatSettlementCurrency(entry.toPayout)}${(entry.employeeSalaries || 0) > 0 ? ` + ${formatSettlementCurrency(entry.employeeSalaries || 0)}` : ''}${(entry.employeeAccountingRefund || 0) > 0 ? ` + ${formatSettlementCurrency(entry.employeeAccountingRefund || 0)}` : ''}${(entry.employeeReceivables || 0) > 0 ? ` - ${formatSettlementCurrency(entry.employeeReceivables || 0)}` : ''} = <strong>${formatSettlementCurrency(entry.grossWithEmployeeSalaries || entry.toPayout)}</strong></div>` : ''}${role === 'separateCompany' ? '' : netFormulaHtml}` : ''}
      </div>

      <div class="settlement-detail-item" style="margin-top: 0.85rem;">
        <div class="settlement-detail-item-header">
          <div>
            <h4>Skąd biorą się godziny i stawki</h4>
            <div class="settlement-detail-meta">Każda pozycja pokazuje źródłowy arkusz, godziny i wyliczoną kwotę.</div>
          </div>
          <strong>${formatSettlementCurrency(personSources.totalSalary)}</strong>
        </div>
        <ul class="settlement-inline-list">${sourcesHtml}</ul>
      </div>

      <div class="settlement-details-grid" style="margin-top: 0.85rem; margin-bottom: 0;">
        <div class="settlement-detail-item">
          <div class="settlement-detail-item-header">
            <div>
              <h4>Zapłacone koszty</h4>
            </div>
            <strong>${formatSettlementCurrency(entry.paidCosts)}</strong>
          </div>
          <ul class="settlement-inline-list">${expenseDetails.paidCosts.length ? expenseDetails.paidCosts.map(item => `<li>${item.label}: ${formatSettlementCurrency(item.amount)}</li>`).join('') : '<li>Brak pozycji.</li>'}</ul>
        </div>
        <div class="settlement-detail-item">
          <div class="settlement-detail-item-header">
            <div>
              <h4>Wypłacone zaliczki</h4>
            </div>
            <strong>${formatSettlementCurrency(entry.paidAdvances)}</strong>
          </div>
          <ul class="settlement-inline-list">${expenseDetails.paidAdvances.length ? expenseDetails.paidAdvances.map(item => `<li>${item.label}: ${formatSettlementCurrency(item.amount)}</li>`).join('') : '<li>Brak pozycji.</li>'}</ul>
        </div>
        <div class="settlement-detail-item">
          <div class="settlement-detail-item-header">
            <div>
              <h4>Premie</h4>
            </div>
            <strong>${formatSettlementCurrency(bonusAmount)}</strong>
          </div>
          <ul class="settlement-inline-list">${expenseDetails.bonusesReceived.length ? expenseDetails.bonusesReceived.map(item => `<li>${item.label}: ${formatSettlementCurrency(item.amount)}</li>`).join('') : '<li>Brak pozycji.</li>'}</ul>
        </div>
        <div class="settlement-detail-item">
          <div class="settlement-detail-item-header">
            <div>
              <h4>Pobrane zaliczki</h4>
            </div>
            <strong>${formatSettlementCurrency(entry.advancesTaken)}</strong>
          </div>
          <ul class="settlement-inline-list">${expenseDetails.advancesTaken.length ? expenseDetails.advancesTaken.map(item => `<li>${item.label}: ${formatSettlementCurrency(item.amount)}</li>`).join('') : '<li>Brak pozycji.</li>'}</ul>
        </div>
      </div>
    </div>
  `;
}

function buildSettlementTaxDetailsSectionHtml(state, res, settlementConfig, selectedMonth) {
  const taxRatePercent = (parseFloat(settlementConfig?.taxRate) || 0) * 100;
  const taxEqualization = Calculations.calculateInvoiceTaxEqualization(state, {
    partners: res.partners,
    workingPartners: res.workingPartners,
    separateCompanies: res.separateCompanies
  }, selectedMonth);
  const partnerEntries = [...(res.partners || [])];
  const workingPartnerEntries = [...(res.workingPartners || [])];
  const separateCompanyEntries = [...(res.separateCompanies || [])];
  const employeeEntries = [...(res.employees || [])].filter(entry =>
    (parseFloat(entry.contractTaxAmount) || 0) > 0
    || (parseFloat(entry.contractZusAmount) || 0) > 0
    || (parseFloat(entry.employerPaidContractCharges) || 0) > 0
  );
  const participantById = Object.fromEntries((taxEqualization.participants || []).map(participant => [participant.personId, participant]));
  const allocationsByIssuerId = {};
  const allocationsByBeneficiaryId = {};
  (taxEqualization.allocations || []).forEach(allocation => {
    if (!allocationsByIssuerId[allocation.issuerId]) allocationsByIssuerId[allocation.issuerId] = [];
    if (!allocationsByBeneficiaryId[allocation.beneficiaryId]) allocationsByBeneficiaryId[allocation.beneficiaryId] = [];
    allocationsByIssuerId[allocation.issuerId].push(allocation);
    allocationsByBeneficiaryId[allocation.beneficiaryId].push(allocation);
  });
  const reimbursementsByPayerId = {};
  const reimbursementsByRecipientId = {};
  (taxEqualization.reimbursements || []).forEach(payment => {
    if (!reimbursementsByPayerId[payment.payerId]) reimbursementsByPayerId[payment.payerId] = [];
    if (!reimbursementsByRecipientId[payment.recipientId]) reimbursementsByRecipientId[payment.recipientId] = [];
    reimbursementsByPayerId[payment.payerId].push(payment);
    reimbursementsByRecipientId[payment.recipientId].push(payment);
  });

  const buildPartnerTaxCardHtml = (entry) => {
    const personId = entry.person?.id;
    const participant = participantById[personId] || {};
    const desiredInvoiceAmount = parseFloat(participant.desiredInvoiceAmount) || 0;
    const issuedAmount = parseFloat(taxEqualization.actualIssuedAmountByIssuer?.[personId]) || 0;
    const officeTaxAmount = parseFloat(taxEqualization.invoiceTaxByIssuer?.[personId]) || 0;
    const ownTaxAmount = parseFloat(taxEqualization.ownTaxByPerson?.[personId]) || 0;
    const sharedCompanyTaxAmount = parseFloat(taxEqualization.sharedCompanyTaxByPerson?.[personId]) || 0;
    const targetTaxAmount = parseFloat(taxEqualization.targetTaxByPerson?.[personId]) || 0;
    const actualTaxBurdenAmount = parseFloat(taxEqualization.actualTaxBurdenByPerson?.[personId]) || 0;
    const separateCompanyReimbursements = parseFloat(taxEqualization.separateCompanyReimbursementsReceivedByIssuer?.[personId]) || 0;
    const incomingReimbursements = reimbursementsByRecipientId[personId] || [];
    const outgoingReimbursements = reimbursementsByPayerId[personId] || [];
    const incomingRefundItems = incomingReimbursements.length
      ? incomingReimbursements.map(item => `<li>${item.payerName}: ${formatSettlementCurrency(item.taxAmount)}${item.category === 'partner-tax-equalization' ? ` (${item.type})` : ` za ${formatSettlementCurrency(item.revenueAmount)} przychodu`}</li>`).join('')
      : '<li>Brak zwrotów.</li>';
    const outgoingRefundItems = outgoingReimbursements.length
      ? outgoingReimbursements.map(item => `<li>${item.recipientName}: ${formatSettlementCurrency(item.taxAmount)}${item.category === 'partner-tax-equalization' ? ` (${item.type})` : ` za ${formatSettlementCurrency(item.revenueAmount)} przychodu`}</li>`).join('')
      : '<li>Brak zwrotów.</li>';

    const uzRows = entry.person?.type === 'WORKING_PARTNER'
      ? `
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Podatek UZ</span><strong class="${entry.contractChargesPaidByEmployer ? '' : 'settlement-accent-negative'}">${entry.contractChargesPaidByEmployer ? formatSettlementCurrency(entry.contractTaxAmount || 0) : `-${formatSettlementCurrency(entry.deductedContractTaxAmount || 0)}`}</strong></div>
          <div class="settlement-detail-meta" style="margin-top: -0.1rem;">Kwota UZ z konfiguracji osoby / miesiąca: ${formatSettlementCurrency(entry.contractTaxAmount || 0)}${entry.contractChargesPaidByEmployer ? `, opłaca pracodawca = ${formatSettlementCurrency(entry.employerPaidContractTaxAmount || 0)}` : `, potrącone z wypłaty = ${formatSettlementCurrency(entry.deductedContractTaxAmount || 0)}`}</div>
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>ZUS UZ</span><strong class="${entry.contractChargesPaidByEmployer ? '' : 'settlement-accent-negative'}">${entry.contractChargesPaidByEmployer ? formatSettlementCurrency(entry.contractZusAmount || 0) : `-${formatSettlementCurrency(entry.deductedContractZusAmount || 0)}`}</strong></div>
          <div class="settlement-detail-meta" style="margin-top: -0.1rem;">Kwota UZ z konfiguracji osoby / miesiąca: ${formatSettlementCurrency(entry.contractZusAmount || 0)}${entry.contractChargesPaidByEmployer ? `, opłaca pracodawca = ${formatSettlementCurrency(entry.employerPaidContractZusAmount || 0)}` : `, potrącone z wypłaty = ${formatSettlementCurrency(entry.deductedContractZusAmount || 0)}`}</div>
        `
      : '';

    return `
      <div class="settlement-detail-item">
        <div class="settlement-detail-item-header">
          <div>
            <h4>${getSettlementPersonDisplayNameHtml(entry.person, state)}</h4>
            <div class="settlement-detail-meta">${getSettlementPersonTypeLabel(entry.person?.type)}</div>
          </div>
          <strong>${formatSettlementCurrency(officeTaxAmount)}</strong>
        </div>
        <div class="settlement-detail-stack" style="margin-top: 0.75rem;">
          <div class="settlement-detail-row"><span>Przychód do zafakturowania z rozliczenia</span><strong>${formatSettlementCurrency(desiredInvoiceAmount)}</strong></div>
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Rzeczywiście wystawione faktury</span><strong>${formatSettlementCurrency(issuedAmount)}</strong></div>
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Opłaca w urzędzie od swoich faktur</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(officeTaxAmount)}</strong></div>
          <div class="settlement-detail-meta" style="margin-top: -0.1rem;">${formatSettlementCurrency(issuedAmount)} × ${taxRatePercent.toFixed(2)}% = ${formatSettlementCurrency(officeTaxAmount)}</div>
          ${separateCompanyReimbursements > 0 ? `<div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Zwrot podatku od osobnych firm</span><strong>${formatSettlementCurrency(separateCompanyReimbursements)}</strong></div>` : ''}
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Rzeczywisty ciężar podatku po zwrotach firm</span><strong>${formatSettlementCurrency(actualTaxBurdenAmount)}</strong></div>
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Podatek własny</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(ownTaxAmount)}</strong></div>
          <div class="settlement-detail-meta" style="margin-top: -0.1rem;">Liczony tylko od pola „Przychód własny(Brutto)”.</div>
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Podatek wspólny</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(sharedCompanyTaxAmount)}</strong></div>
          <div class="settlement-detail-meta" style="margin-top: -0.1rem;">Równa część podatku spółki: zwrot zaliczek + pensje pracowników + premie pracowników + zwrot Podatku i ZUS za pracowników.</div>
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Docelowy podatek do poniesienia</span><strong>${formatSettlementCurrency(targetTaxAmount)}</strong></div>
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Dostaje zwrot podatku</span><strong>${formatSettlementCurrency(incomingReimbursements.reduce((sum, item) => sum + (item.taxAmount || 0), 0))}</strong></div>
          <ul class="settlement-inline-list">${incomingRefundItems}</ul>
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Oddaje zwrot podatku</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(outgoingReimbursements.reduce((sum, item) => sum + (item.taxAmount || 0), 0))}</strong></div>
          <ul class="settlement-inline-list">${outgoingRefundItems}</ul>
          ${uzRows}
        </div>
      </div>
    `;
  };

  const buildWorkingPartnerTaxCardHtml = (entry) => `
    <div class="settlement-detail-item">
      <div class="settlement-detail-item-header">
        <div>
          <h4>${getSettlementPersonDisplayName(entry.person, state)}</h4>
          <div class="settlement-detail-meta">Wspólnik pracujący</div>
        </div>
        <strong>${formatSettlementCurrency((parseFloat(entry.ownTaxAmount) || 0) + (parseFloat(entry.contractTaxAmount) || 0) + (parseFloat(entry.contractZusAmount) || 0))}</strong>
      </div>
      <div class="settlement-detail-stack" style="margin-top: 0.75rem;">
        <div class="settlement-detail-row"><span>Podatek własny</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(entry.ownTaxAmount || 0)}</strong></div>
        <div class="settlement-detail-row"><span>Podatek UZ</span><strong class="${entry.contractChargesPaidByEmployer ? '' : 'settlement-accent-negative'}">${entry.contractChargesPaidByEmployer ? formatSettlementCurrency(entry.contractTaxAmount || 0) : `-${formatSettlementCurrency(entry.deductedContractTaxAmount || 0)}`}</strong></div>
        <div class="settlement-detail-row"><span>ZUS UZ</span><strong class="${entry.contractChargesPaidByEmployer ? '' : 'settlement-accent-negative'}">${entry.contractChargesPaidByEmployer ? formatSettlementCurrency(entry.contractZusAmount || 0) : `-${formatSettlementCurrency(entry.deductedContractZusAmount || 0)}`}</strong></div>
      </div>
    </div>
  `;

  const buildCompanyTaxPanelHtml = () => {
    const totalAdvanceRefunds = partnerEntries.reduce((sum, entry) => sum + (parseFloat(entry.paidAdvances) || 0), 0);
    const totalEmployeeSalaries = partnerEntries.reduce((sum, entry) => sum + (parseFloat(entry.employeeSalaries) || 0), 0);
    const totalEmployeeBonuses = partnerEntries.reduce((sum, entry) => sum + (parseFloat(entry.employeeBonuses) || 0), 0);
    const totalEmployeeAccountingRefunds = partnerEntries.reduce((sum, entry) => sum + (parseFloat(entry.employeeAccountingRefund) || 0), 0);
    const totalEmployeeReceivables = partnerEntries.reduce((sum, entry) => sum + (parseFloat(entry.employeeReceivables) || 0), 0);
    const totalBaseEmployeeSalaries = totalEmployeeSalaries - totalEmployeeBonuses;
    const taxBase = parseFloat(res.partnerSharedCompanyTaxBaseTotal) || 0;
    const taxTotal = parseFloat(res.partnerSharedCompanyTaxTotal) || 0;
    const taxPerPartner = parseFloat(res.partnerSharedCompanyTaxPerPerson) || 0;
    const participantCount = parseFloat(res.partnerSharedCompanyTaxParticipantsCount) || 0;

    return `
      <div class="settlement-detail-item">
        <div class="settlement-detail-item-header">
          <div>
            <h4>Podatek spółki</h4>
            <div class="settlement-detail-meta">To część podatku dzielona po równo między aktywnych wspólników.</div>
          </div>
          <strong>${formatSettlementCurrency(taxTotal)}</strong>
        </div>
        <div class="settlement-detail-stack" style="margin-top: 0.75rem;">
          <div class="settlement-detail-row"><span>Zwrot zaliczek</span><strong>${formatSettlementCurrency(totalAdvanceRefunds)}</strong></div>
          <div class="settlement-detail-row"><span>Pensje pracowników</span><strong>${formatSettlementCurrency(totalBaseEmployeeSalaries)}</strong></div>
          ${totalEmployeeBonuses > 0 ? `<div class="settlement-detail-row"><span>Premie pracowników</span><strong>${formatSettlementCurrency(totalEmployeeBonuses)}</strong></div>` : ''}
          <div class="settlement-detail-row"><span>Zwrot Podatku i ZUS za pracowników</span><strong>${formatSettlementCurrency(totalEmployeeAccountingRefunds)}</strong></div>
          ${totalEmployeeReceivables > 0 ? `<div class="settlement-detail-row"><span>Do odebrania od pracowników</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(totalEmployeeReceivables)}</strong></div>` : ''}
        </div>
        <div class="settlement-detail-formula" style="margin-top: 0.85rem;">
          ${formatSettlementCurrency(totalAdvanceRefunds)} + ${formatSettlementCurrency(totalBaseEmployeeSalaries)}${totalEmployeeBonuses > 0 ? ` + ${formatSettlementCurrency(totalEmployeeBonuses)}` : ''} + ${formatSettlementCurrency(totalEmployeeAccountingRefunds)}${totalEmployeeReceivables > 0 ? ` - ${formatSettlementCurrency(totalEmployeeReceivables)}` : ''} = <strong>${formatSettlementCurrency(taxBase)}</strong>
          <div style="margin-top: 0.45rem;">${formatSettlementCurrency(taxBase)} × ${taxRatePercent.toFixed(2)}% = <strong>${formatSettlementCurrency(taxTotal)}</strong></div>
          <div style="margin-top: 0.45rem;">${formatSettlementCurrency(taxTotal)} / ${participantCount || 0} = <strong>${formatSettlementCurrency(taxPerPartner)}</strong> na wspólnika</div>
        </div>
      </div>
    `;
  };

  const buildSeparateCompanyTaxCardHtml = (entry) => {
    const personId = entry.person?.id;
    const participant = participantById[personId] || {};
    const desiredInvoiceAmount = parseFloat(participant.desiredInvoiceAmount) || 0;
    const issuedAmount = parseFloat(taxEqualization.actualIssuedAmountByIssuer?.[personId]) || 0;
    const selfIssuedAllocations = (allocationsByBeneficiaryId[personId] || []).filter(item => item.issuerId === personId);
    const partnerIssuedAllocations = (allocationsByBeneficiaryId[personId] || []).filter(item => item.issuerId !== personId && item.issuerType !== 'SEPARATE_COMPANY');
    const outgoingReimbursements = reimbursementsByPayerId[personId] || [];

    return `
      <div class="settlement-detail-item">
        <div class="settlement-detail-item-header">
          <div>
            <h4>${getSettlementPersonDisplayNameHtml(entry.person, state)}</h4>
            <div class="settlement-detail-meta">Osobna firma</div>
          </div>
          <strong>${formatSettlementCurrency(outgoingReimbursements.reduce((sum, item) => sum + (item.taxAmount || 0), 0))}</strong>
        </div>
        <div class="settlement-detail-stack" style="margin-top: 0.75rem;">
          <div class="settlement-detail-row"><span>Przychód do zafakturowania z rozliczenia</span><strong>${formatSettlementCurrency(desiredInvoiceAmount)}</strong></div>
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Rzeczywiście wystawione przez tę firmę</span><strong>${formatSettlementCurrency(issuedAmount)}</strong></div>
          <div class="settlement-detail-meta" style="margin-top: -0.1rem;">Podatku własnej firmy nie pokazujemy w polu „Opłaca w urzędzie” — firma rozlicza go po swojej stronie.</div>
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Wystawione przez wspólników za tę firmę</span><strong>${formatSettlementCurrency(partnerIssuedAllocations.reduce((sum, item) => sum + (item.amount || 0), 0))}</strong></div>
          <ul class="settlement-inline-list">${partnerIssuedAllocations.length ? partnerIssuedAllocations.map(item => `<li>${item.issuerName}: ${formatSettlementCurrency(item.amount)} przychodu</li>`).join('') : '<li>Brak przychodu tej firmy na fakturach wspólników.</li>'}</ul>
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Zwrot podatku dla wspólników</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(outgoingReimbursements.reduce((sum, item) => sum + (item.taxAmount || 0), 0))}</strong></div>
          <ul class="settlement-inline-list">${outgoingReimbursements.length ? outgoingReimbursements.map(item => `<li>${item.recipientName}: ${formatSettlementCurrency(item.taxAmount)}${item.category === 'partner-tax-equalization' ? ` (${item.type})` : ` za ${formatSettlementCurrency(item.revenueAmount)} przychodu`}</li>`).join('') : '<li>Brak zwrotów podatku.</li>'}</ul>
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Własna część wystawiona na firmę</span><strong>${formatSettlementCurrency(selfIssuedAllocations.reduce((sum, item) => sum + (item.amount || 0), 0))}</strong></div>
        </div>
      </div>
    `;
  };

  const buildEmployeeUzCardHtml = (entry) => {
    const employerName = entry.person?.employerId ? getEmployerNameById(state, entry.person.employerId) : '';
    const employerPaidCharges = (parseFloat(entry.employerPaidContractTaxAmount) || 0) + (parseFloat(entry.employerPaidContractZusAmount) || 0);
    return `
      <div class="settlement-detail-item">
        <div class="settlement-detail-item-header">
          <div>
            <h4>${getSettlementPersonDisplayNameHtml(entry.person, state)}</h4>
            <div class="settlement-detail-meta">${employerName ? `Pracownik • pracodawca: ${employerName}` : 'Pracownik'}</div>
          </div>
          <strong>${formatSettlementCurrency((parseFloat(entry.contractTaxAmount) || 0) + (parseFloat(entry.contractZusAmount) || 0))}</strong>
        </div>
        <div class="settlement-detail-stack" style="margin-top: 0.75rem;">
          <div class="settlement-detail-row"><span>Podatek UZ</span><strong class="${entry.contractChargesPaidByEmployer ? '' : 'settlement-accent-negative'}">${entry.contractChargesPaidByEmployer ? formatSettlementCurrency(entry.contractTaxAmount || 0) : `-${formatSettlementCurrency(entry.deductedContractTaxAmount || 0)}`}</strong></div>
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>ZUS UZ</span><strong class="${entry.contractChargesPaidByEmployer ? '' : 'settlement-accent-negative'}">${entry.contractChargesPaidByEmployer ? formatSettlementCurrency(entry.contractZusAmount || 0) : `-${formatSettlementCurrency(entry.deductedContractZusAmount || 0)}`}</strong></div>
          ${entry.contractChargesPaidByEmployer && employerPaidCharges > 0 ? `<div class="settlement-detail-meta" style="margin-top: 0.35rem;">Podatki i ZUS opłacane przez pracodawcę (${formatSettlementCurrency(employerPaidCharges)}) są odejmowane od zysku, który wypracował pracownik.</div>` : ''}
        </div>
      </div>
    `;
  };

  return `
    <div class="settlement-detail-section">
      <h3>Podatki i ZUS — skąd biorą się kwoty</h3>
      <div style="margin-bottom: 1rem;">
        ${buildCompanyTaxPanelHtml()}
      </div>
      <div class="settlement-details-grid" style="grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); align-items: start;">
        ${partnerEntries.length ? partnerEntries.map(buildPartnerTaxCardHtml).join('') : '<p class="settlement-detail-empty">Brak wspólników do rozpisania podatków.</p>'}
      </div>
      <div class="settlement-details-grid" style="margin-top: 1rem; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); align-items: start;">
        <div>
          <h4 style="margin-bottom: 0.75rem;">Wspólnicy pracujący</h4>
          <div class="settlement-detail-list">${workingPartnerEntries.length ? workingPartnerEntries.map(buildWorkingPartnerTaxCardHtml).join('') : '<p class="settlement-detail-empty">Brak wspólników pracujących do rozpisania podatków.</p>'}</div>
        </div>
        <div>
          <h4 style="margin-bottom: 0.75rem;">Osobne firmy</h4>
          <div class="settlement-detail-list">${separateCompanyEntries.length ? separateCompanyEntries.map(buildSeparateCompanyTaxCardHtml).join('') : '<p class="settlement-detail-empty">Brak osobnych firm do rozpisania podatków.</p>'}</div>
        </div>
        <div>
          <h4 style="margin-bottom: 0.75rem;">Pracownicy — podatki UZ</h4>
          <div class="settlement-detail-list">${employeeEntries.length ? employeeEntries.map(buildEmployeeUzCardHtml).join('') : '<p class="settlement-detail-empty">Brak podatków UZ dla pracowników w wybranym miesiącu.</p>'}</div>
        </div>
      </div>
    </div>
  `;
}

function renderSettlementDetails(state, res) {
  const content = document.getElementById('settlement-details-content');
  if (!content) return;

  const details = getSettlementDetailsData(state, res);
  const settlementConfig = res.settlementConfig || Calculations.getSettlementConfig(state, details.selectedMonth);
  const totalUnreduced = res.commonRevenue;
  const advances = res.clientAdvances || 0;
  const finalRevenue = totalUnreduced - advances;
  const employerPaidContractCharges = res.employerPaidContractCharges || 0;
  const employeeBonuses = details.totalEmployeeBonuses || 0;
  const basicEmployeeProfitBeforeEmployerCharges = res.employeeRevenue - details.totalEmployeesSalaryFromHours;
  const basicEmployeeProfit = basicEmployeeProfitBeforeEmployerCharges - employeeBonuses - employerPaidContractCharges;
  const profitFromEmployees = res.employeeProfitShared || 0;
  const extraEmployeeProfitShare = profitFromEmployees - basicEmployeeProfit;
  const accountingPeopleCount = res.partners.length + res.workingPartners.length;
  const clientAdvanceSum = details.clientAdvanceItems.reduce((sum, item) => sum + item.amount, 0);
  const otherAdvanceSum = details.otherAdvanceItems.reduce((sum, item) => sum + item.amount, 0);
  const totalBonusSum = details.bonusItems.reduce((sum, item) => sum + item.amount, 0);
  const totalAdvanceSum = clientAdvanceSum + otherAdvanceSum;
  const totalPartnersGrossWithEmployees = [...res.partners, ...(res.separateCompanies || []), ...res.workingPartners]
    .reduce((sum, person) => {
      if (person.person?.type === 'WORKING_PARTNER' && person.person?.employerId) {
        return sum;
      }

      return sum + (((person.employeeSalaries || 0) > 0 || (person.employeeReceivables || 0) > 0)
        ? (person.grossWithEmployeeSalaries || 0)
        : (person.toPayout || 0));
    }, 0);
  const partnersGrossDifference = totalPartnersGrossWithEmployees - finalRevenue;
  const partnersGrossMatch = Math.abs(partnersGrossDifference) < 0.005;
  const grossPayoutsDifference = details.totalGrossPayouts - finalRevenue;
  const grossPayoutsMatch = Math.abs(grossPayoutsDifference) < 0.005;
  const grossPayoutsDifferenceText = grossPayoutsMatch
    ? 'Suma wszystkich wypłat przed podatkiem i ZUS zgadza się z przychodem brutto z odjętymi zaliczkami klienta.'
    : (grossPayoutsDifference > 0
      ? `Suma wszystkich wypłat przed podatkiem i ZUS jest wyższa od przychodu brutto z odjętymi zaliczkami klienta o ${formatSettlementCurrency(grossPayoutsDifference)}.`
      : `Suma wszystkich wypłat przed podatkiem i ZUS jest niższa od przychodu brutto z odjętymi zaliczkami klienta o ${formatSettlementCurrency(Math.abs(grossPayoutsDifference))}.`);

  content.innerHTML = `
    <div class="settlement-details-grid">
      <div class="settlement-detail-card">
        <h3>1. Cały przychód</h3>
        <div class="settlement-detail-stack">
          <div class="settlement-detail-row"><span>Arkusze godzin</span><strong>${formatSettlementCurrency(details.totalHourlyRevenue)}</strong></div>
          <div class="settlement-detail-row"><span>Arkusze wykonanych prac</span><strong>${formatSettlementCurrency(details.totalWorksRevenue)}</strong></div>
          <div class="settlement-detail-row"><span>Przychód brutto przed zaliczkami klienta</span><strong>${formatSettlementCurrency(totalUnreduced)}</strong></div>
          <div class="settlement-detail-row"><span>Zaliczki od klientów</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(advances)}</strong></div>
        </div>
        <div class="settlement-detail-formula">${formatSettlementCurrency(totalUnreduced)} - ${formatSettlementCurrency(advances)} = <strong>${formatSettlementCurrency(finalRevenue)}</strong></div>
      </div>

      <div class="settlement-detail-card">
        <h3>2. Zysk z godzin pracowników</h3>
        <div class="settlement-detail-stack">
          <div class="settlement-detail-row"><span>Przychód z zakładki Godziny wypracowany przez pracowników</span><strong>${formatSettlementCurrency(res.employeeRevenue)}</strong></div>
          <div class="settlement-detail-row"><span>Pensje pracowników z zakładki Godziny</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(details.totalEmployeesSalaryFromHours)}</strong></div>
          ${employeeBonuses > 0 ? `<div class="settlement-detail-row"><span>Premie pracowników</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(employeeBonuses)}</strong></div>` : ''}
          ${employerPaidContractCharges > 0 ? `<div class="settlement-detail-row"><span>Podatek i ZUS UZ opłacony przez pracodawcę</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(employerPaidContractCharges)}</strong></div>` : ''}
          ${Math.abs(extraEmployeeProfitShare) >= 0.005 ? `<div class="settlement-detail-row"><span>Dodatkowy podział zysków firm i ich pracowników</span><strong>${formatSettlementCurrency(extraEmployeeProfitShare)}</strong></div>` : ''}
        </div>
        <div class="settlement-detail-formula">${formatSettlementCurrency(res.employeeRevenue)} - ${formatSettlementCurrency(details.totalEmployeesSalaryFromHours)}${employeeBonuses > 0 ? ` - ${formatSettlementCurrency(employeeBonuses)}` : ''}${employerPaidContractCharges > 0 ? ` - ${formatSettlementCurrency(employerPaidContractCharges)}` : ''}${Math.abs(extraEmployeeProfitShare) >= 0.005 ? ` ${extraEmployeeProfitShare >= 0 ? '+' : '-'} ${formatSettlementCurrency(Math.abs(extraEmployeeProfitShare))}` : ''} = <strong>${formatSettlementCurrency(profitFromEmployees)}</strong></div>
      </div>

      <div class="settlement-detail-card">
        <h3>3. Zysk z wykonanych prac</h3>
        <div class="settlement-detail-stack">
          <div class="settlement-detail-row"><span>Wartość wykonanych prac do podziału</span><strong>${formatSettlementCurrency(details.totalWorksProfitBaseRevenue)}</strong></div>
          <div class="settlement-detail-row"><span>Koszt pracowników w wykonanych pracach</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(details.totalWorksEmployeeCost)}</strong></div>
        </div>
        <div class="settlement-detail-formula">${formatSettlementCurrency(details.totalWorksProfitBaseRevenue)} - ${formatSettlementCurrency(details.totalWorksEmployeeCost)} = <strong>${formatSettlementCurrency(res.totalWorksProfit)}</strong></div>
      </div>

      <div class="settlement-detail-card">
        <h3>4. Zarobek do podziału</h3>
        <div class="settlement-detail-stack">
          <div class="settlement-detail-row"><span>Zysk z godzin pracowników</span><strong>${formatSettlementCurrency(profitFromEmployees)}</strong></div>
          <div class="settlement-detail-row"><span>Zysk z wykonanych prac</span><strong>${formatSettlementCurrency(res.totalWorksProfit)}</strong></div>
        </div>
        <div class="settlement-detail-formula">${formatSettlementCurrency(profitFromEmployees)} + ${formatSettlementCurrency(res.totalWorksProfit)} = <strong>${formatSettlementCurrency(res.profitToSplit)}</strong></div>
      </div>

      <div class="settlement-detail-card">
        <h3>5. Koszty wspólne</h3>
        <div class="settlement-detail-stack">
          <div class="settlement-detail-row"><span>Suma kosztów wspólnych</span><strong>${formatSettlementCurrency(details.totalCommonCosts)}</strong></div>
          <div class="settlement-detail-row"><span>Liczba osób z udziałem w kosztach</span><strong>${details.activeCostShareCount}</strong></div>
          <div class="settlement-detail-row"><span>Udział w kosztach na 1 osobę</span><strong>${formatSettlementCurrency(details.costShare)}</strong></div>
        </div>
        <div class="settlement-detail-formula">${details.activeCostShareCount > 0 ? `${formatSettlementCurrency(details.totalCommonCosts)} / ${details.activeCostShareCount} = <strong>${formatSettlementCurrency(details.costShare)}</strong>` : 'Brak osób z udziałem w kosztach — koszt wspólny nie jest dzielony.'}</div>
      </div>

      <div class="settlement-detail-card">
        <h3>6. Zaliczki i premie</h3>
        <div class="settlement-detail-stack">
          <div class="settlement-detail-row"><span>Suma zaliczek od klientów</span><strong>${formatSettlementCurrency(clientAdvanceSum)}</strong></div>
          <div class="settlement-detail-row"><span>Suma pozostałych zaliczek</span><strong>${formatSettlementCurrency(otherAdvanceSum)}</strong></div>
          <div class="settlement-detail-row"><span>Suma premii pracowników</span><strong>${formatSettlementCurrency(totalBonusSum)}</strong></div>
          <div class="settlement-detail-row"><span>Suma wszystkich zaliczek</span><strong>${formatSettlementCurrency(totalAdvanceSum)}</strong></div>
        </div>
        <div class="settlement-detail-formula">Zaliczki: ${formatSettlementCurrency(clientAdvanceSum)} + ${formatSettlementCurrency(otherAdvanceSum)} = <strong>${formatSettlementCurrency(totalAdvanceSum)}</strong>${totalBonusSum > 0 ? `<div style="margin-top: 0.45rem;">Premie pracowników: <strong>${formatSettlementCurrency(totalBonusSum)}</strong></div>` : ''}</div>
      </div>

      <div class="settlement-detail-card">
        <h3>7. Podatek i ZUS</h3>
        <div class="settlement-detail-stack">
          <div class="settlement-detail-row"><span>Podatek / ryczałt</span><strong>${((parseFloat(settlementConfig.taxRate) || 0) * 100).toFixed(2)}%</strong></div>
          <div class="settlement-detail-row"><span>Stały ZUS na osobę</span><strong>${formatSettlementCurrency(settlementConfig.zusFixedAmount)}</strong></div>
          <div class="settlement-detail-row"><span>Łączny podatek wspólny</span><strong>${formatSettlementCurrency(res.partnerSharedCompanyTaxTotal || 0)}</strong></div>
          <div class="settlement-detail-row"><span>Podatek wspólny na wspólnika</span><strong>${formatSettlementCurrency(res.partnerSharedCompanyTaxPerPerson || 0)}</strong></div>
        </div>
      </div>

      <div class="settlement-detail-card">
        <h3>8. Sprawdzenie</h3>
        <div class="settlement-detail-stack">
          <div class="settlement-detail-row"><span>Przychód brutto z odjętymi zaliczkami klienta</span><strong>${formatSettlementCurrency(finalRevenue)}</strong></div>
          <div class="settlement-detail-row"><span>Suma wszystkich wypłat przed podatkiem i ZUS</span><strong>${formatSettlementCurrency(details.totalGrossPayouts)}</strong></div>
          <div class="settlement-detail-row"><span>Suma przychodów brutto wspólników z pensjami pracowników</span><strong>${formatSettlementCurrency(totalPartnersGrossWithEmployees)}</strong></div>
          <div class="settlement-detail-row"><span>Porównanie</span><strong class="${partnersGrossMatch ? 'settlement-accent-positive' : 'settlement-accent-negative'}">${partnersGrossMatch ? 'Zgadza się' : 'Nie zgadza się'}</strong></div>
        </div>
        ${!partnersGrossMatch ? `<div class="settlement-detail-formula">${formatSettlementCurrency(totalPartnersGrossWithEmployees)} - ${formatSettlementCurrency(finalRevenue)} = <strong>${formatSettlementCurrency(partnersGrossDifference)}</strong></div>` : ''}
        ${!grossPayoutsMatch ? `<div class="settlement-detail-formula" style="margin-top: 0.75rem;">${grossPayoutsDifferenceText}</div>` : ''}
      </div>
    </div>

    <div class="settlement-detail-section">
      <h3>Źródła przychodu z arkuszy godzin</h3>
      <div class="settlement-detail-list">
        ${buildSettlementEntryListHtml(
          details.hourlyRevenueItems.map(item => ({
            ...item,
            meta: `${formatSettlementHours(item.hours)} łącznie • w tym pracownicy ${formatSettlementHours(item.employeeHours)}`,
            description: `${formatSettlementHours(item.hours)} × ${formatSettlementCurrency(item.clientRate).replace(' zł', ' zł/h')}`
          })),
          'revenue',
          'Brak przychodów z arkuszy godzin w tym miesiącu.'
        )}
      </div>
    </div>

    <div class="settlement-detail-section">
      <h3>Źródła przychodu z wykonanych prac</h3>
      <div class="settlement-detail-list">
        ${details.worksRevenueItems.length ? details.worksRevenueItems.map(item => `
          <div class="settlement-detail-item">
            <div class="settlement-detail-item-header">
              <div>
                <h4>${item.label}</h4>
                <div class="settlement-detail-meta">Wartość arkusza ${formatSettlementCurrency(item.revenue)} • koszt pracowników ${formatSettlementCurrency(item.employeeCost)}</div>
              </div>
              <strong>${formatSettlementCurrency(item.profit)}</strong>
            </div>
            <div class="settlement-detail-formula" style="margin-top: 0; margin-bottom: 0.85rem;">${item.entries.some(entry => entry.profitBaseRevenue !== entry.revenue) ? `${formatSettlementCurrency(item.profitBaseRevenue)} - ${formatSettlementCurrency(item.employeeCost)} = <strong>${formatSettlementCurrency(item.profit)}</strong>` : `${formatSettlementCurrency(item.revenue)} - ${formatSettlementCurrency(item.employeeCost)} = <strong>${formatSettlementCurrency(item.profit)}</strong>`}</div>
            <div class="settlement-detail-item" style="margin-top: 0; margin-bottom: 0.85rem;">
              <div class="settlement-detail-item-header">
                <div>
                  <h4>Podział zysku z arkusza</h4>
                  <div class="settlement-detail-meta">Pokazuje, jak zysk z tego arkusza dzieli się między aktywne osoby uprawnione do udziału.</div>
                </div>
                <strong>${formatSettlementCurrency(item.profit)}</strong>
              </div>
              <ul class="settlement-inline-list">${item.profitSplitItems.length
                ? item.profitSplitItems.map(split => `<li>${split.name} (${split.typeLabel}): ${formatSettlementCurrency(split.amount)}</li>`).join('')
                : '<li>Brak aktywnych osób uprawnionych do udziału w zysku.</li>'}</ul>
            </div>
            <div class="settlement-detail-list">
              ${buildSettlementEntryListHtml(item.entries.map(entry => ({
                ...entry,
                meta: entry.profitBaseRevenue !== entry.revenue
                  ? `${entry.profitBaseLabel}: ${formatSettlementCurrency(entry.profitBaseRevenue)}`
                  : ''
              })), 'revenue', 'Brak pozycji w tym arkuszu.')}
            </div>
          </div>
        `).join('') : '<p class="settlement-detail-empty">Brak przychodu z wykonanych prac w tym miesiącu.</p>'}
      </div>
    </div>

    <div class="settlement-detail-section">
      <h3>Koszty i zaliczki z miesiąca ${formatMonthLabel(details.selectedMonth)}</h3>
      <div class="settlement-details-grid">
        <div>
          <h4 style="margin-bottom: 0.75rem;">Koszty wspólne</h4>
          <div class="settlement-detail-list">
            ${buildSettlementEntryListHtml(
              details.selectedExpenses.filter(item => item.type === 'COST').map(item => ({
                label: `${item.date} • ${item.name || 'Koszt'}`,
                amount: item.amount,
                meta: `Zapłacił: ${getSettlementEntityName(state, item.paidById)}`
              })),
              'amount',
              'Brak kosztów wspólnych w wybranym miesiącu.'
            )}
          </div>
        </div>
        <div>
          <h4 style="margin-bottom: 0.75rem;">Zaliczki od klientów</h4>
          <div class="settlement-detail-list">
            ${buildSettlementEntryListHtml(
              details.clientAdvanceItems.map(item => ({
                label: `${item.date} • ${item.name || 'Zaliczka klienta'}`,
                amount: item.amount,
                meta: `Od: ${getSettlementEntityName(state, item.paidById)}${item.advanceForId ? ` • dla ${getSettlementEntityName(state, item.advanceForId)}` : ''}`
              })),
              'amount',
              'Brak zaliczek klienta w wybranym miesiącu.'
            )}
          </div>
        </div>
        <div>
          <h4 style="margin-bottom: 0.75rem;">Pozostałe zaliczki</h4>
          <div class="settlement-detail-list">
            ${buildSettlementEntryListHtml(
              details.otherAdvanceItems.map(item => ({
                label: `${item.date} • ${item.name || 'Zaliczka'}`,
                amount: item.amount,
                meta: `Zapłacił: ${getSettlementEntityName(state, item.paidById)}${item.advanceForId ? ` • odbiorca ${getSettlementEntityName(state, item.advanceForId)}` : ''}`
              })),
              'amount',
              'Brak zaliczek wewnętrznych w wybranym miesiącu.'
            )}
          </div>
        </div>
        <div>
          <h4 style="margin-bottom: 0.75rem;">Premie</h4>
          <div class="settlement-detail-list">
            ${buildSettlementEntryListHtml(
              details.bonusItems.map(item => ({
                label: `${item.date} • Premia`,
                amount: item.amount,
                meta: `Dla: ${getSettlementEntityName(state, item.advanceForId)} • płaci: Od Wszystkich Wspólników`
              })),
              'amount',
              'Brak premii w wybranym miesiącu.'
            )}
          </div>
        </div>
      </div>
    </div>

    ${buildSettlementProfitBreakdownSectionHtml(res, buildSettlementTaxDetailsSectionHtml(state, res, settlementConfig, details.selectedMonth))}

    <div class="settlement-detail-section">
      <h3>Rozpiska każdej osoby</h3>
      <div class="settlement-person-list">
        ${res.partners.map(person => buildSettlementPersonCardHtml(person, 'partner', state, details)).join('')}
        ${(res.separateCompanies || []).map(person => buildSettlementPersonCardHtml(person, 'separateCompany', state, details)).join('')}
        ${res.workingPartners.map(person => buildSettlementPersonCardHtml(person, 'workingPartner', state, details)).join('')}
        ${res.employees.map(person => buildSettlementPersonCardHtml(person, 'employee', state, details)).join('')}
      </div>
    </div>
  `;
}

function initSettlement() {
  const btnEditConfig = document.getElementById('btn-edit-config');
  const btnToggleDetails = document.getElementById('btn-toggle-settlement-details');
  const configPanel = document.getElementById('config-panel');
  const detailsPanel = document.getElementById('settlement-details-panel');
  const btnSaveConfig = document.getElementById('btn-save-config');
  const btnFetchZus = document.getElementById('btn-fetch-zus');
  const zusStatus = document.getElementById('config-zus-status');
  const zusThreshold = document.getElementById('config-zus-threshold');
  const zusWithSick = document.getElementById('config-zus-with-sick');

  if (btnToggleDetails && detailsPanel) {
    btnToggleDetails.addEventListener('click', () => {
      const shouldShow = detailsPanel.style.display === 'none';
      detailsPanel.style.display = shouldShow ? 'block' : 'none';
      btnToggleDetails.textContent = shouldShow ? 'Ukryj szczegóły' : 'Szczegóły';

      if (shouldShow) {
        detailsPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    });
  }

  btnEditConfig.addEventListener('click', () => {
    populateSettlementConfigPanel();
    
    if (configPanel.style.display === 'none') {
      configPanel.style.display = 'block';
      configPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
    } else {
      configPanel.style.display = 'none';
    }
  });

  btnSaveConfig.addEventListener('click', () => {
    const selectedMonth = getSelectedMonthKey();
    const monthTaxRaw = document.getElementById('config-month-tax').value.trim();
    const monthZusRaw = document.getElementById('config-month-zus').value.trim();
    const personContractCharges = {};

    document.querySelectorAll('#settlement-person-contract-grid .glass-panel[data-person-id]').forEach(card => {
      const personId = card.getAttribute('data-person-id');
      const taxInput = card.querySelector('.config-month-contract-tax');
      const zusInput = card.querySelector('.config-month-contract-zus');
      const taxRaw = taxInput ? taxInput.value.trim() : '';
      const zusRaw = zusInput ? zusInput.value.trim() : '';

      if (taxRaw === '' && zusRaw === '') return;

      personContractCharges[personId] = {
        contractTaxAmount: taxRaw === '' ? null : (parseFloat(taxRaw) || 0),
        contractZusAmount: zusRaw === '' ? null : (parseFloat(zusRaw) || 0)
      };
    });

    Store.updateConfig({
      taxRate: (parseFloat(document.getElementById('config-tax').value) || 0) / 100,
      zusFixedAmount: parseFloat(document.getElementById('config-zus').value) || 0
    });
    Store.updateSettlementMonthConfig({
      taxRate: monthTaxRaw === '' ? null : (parseFloat(monthTaxRaw) || 0) / 100,
      zusFixedAmount: monthZusRaw === '' ? null : (parseFloat(monthZusRaw) || 0)
    }, personContractCharges, selectedMonth);
    configPanel.style.display = 'none';
  });

  btnFetchZus.addEventListener('click', async () => {
    const originalText = btnFetchZus.textContent;
    btnFetchZus.disabled = true;
    btnFetchZus.textContent = 'Pobieranie...';
    zusStatus.textContent = 'Trwa pobieranie aktualnej stawki ZUS z internetu...';

    try {
      const result = await fetchCurrentZusRateForRyczalt({
        threshold: zusThreshold.value,
        includeSick: zusWithSick.checked
      });
      document.getElementById('config-zus').value = result.amount.toFixed(2);
      const sourceLabel = result.source.startsWith('fallback-')
        ? 'Nie udało się pobrać online, użyto danych aplikacji.'
        : 'Pobrano online.';
      zusStatus.textContent = `${sourceLabel} ${result.amount.toFixed(2)} zł (${result.thresholdLabel}, zdrowotna ${result.healthAmount.toFixed(2)} zł, społeczne ${result.socialAmount.toFixed(2)} zł${result.includeSick ? ', z chorobowym' : ', bez chorobowego'}) ${result.source.startsWith('fallback-') ? '' : `z ${result.source}.`}`.trim();
      window.lastZusDebug = result.debugInfo || [];
    } catch (error) {
      const debugInfo = Array.isArray(error.debugInfo) ? error.debugInfo : [];
      window.lastZusDebug = debugInfo;
      console.error('ZUS fetch debug:', debugInfo);
      zusStatus.textContent = debugInfo.length > 0
        ? `${error.message} Debug: ${debugInfo.join(' | ')}`
        : error.message;
      alert(error.message);
    } finally {
      btnFetchZus.disabled = false;
      btnFetchZus.textContent = originalText;
    }
  });
}

function renderSettlement() {
  const state = Store.getState();
  const res = Calculations.generateSettlement(state);
  const selectedMonth = Calculations.getSelectedMonth(state);
  const getDisplayedNetAmount = (entry) => (parseFloat(entry?.toPayout) || 0) - (parseFloat(entry?.taxAmount) || 0) - (parseFloat(entry?.zusAmount) || 0);
  populateSettlementConfigPanel();

  const totalUnreduced = res.commonRevenue;
  const advances = res.clientAdvances || 0;
  const finalRevenue = totalUnreduced - advances;
  const profitFromEmployees = res.employeeProfitShared || 0;
  document.getElementById('set-total-revenue').textContent = `${finalRevenue.toFixed(2)} zł`;
  
  const formatEl = document.getElementById('set-revenue-formula');
  if (formatEl) {
    if (advances > 0) {
      formatEl.textContent = `Cały przychód bez zaliczki (${totalUnreduced.toFixed(2)} zł) - zaliczka klienta (${advances.toFixed(2)} zł) = ${finalRevenue.toFixed(2)} zł`;
      formatEl.style.display = 'block';
    } else {
      formatEl.style.display = 'none';
    }
  }

  document.getElementById('set-profit').textContent = `${res.profitToSplit.toFixed(2)} zł`;
  document.getElementById('set-profit-formula').textContent = 
    `Formuła: zysk z pracowników (${profitFromEmployees.toFixed(2)} zł) + zysk z prac (${res.totalWorksProfit.toFixed(2)} zł) = ${res.profitToSplit.toFixed(2)} zł`;

  const partnersContainer = document.getElementById('set-partners-list');
  const workingPartnersContainer = document.getElementById('set-working-partners-list');
  const employeesContainer = document.getElementById('set-employees-list');
  const separateCompaniesContainer = document.getElementById('set-separate-companies-list');
  const leftColumn = document.getElementById('settlement-left-column');
  const rightColumn = document.getElementById('settlement-right-column');
  const partnersPanel = document.getElementById('settlement-panel-partners');
  const workingPartnersPanel = document.getElementById('settlement-panel-working-partners');
  const employeesPanel = document.getElementById('settlement-panel-employees');
  const separateCompaniesPanel = document.getElementById('settlement-panel-separate-companies');

  const setPanelVisibility = (panelEl, hasContent) => {
    if (!panelEl) return;
    panelEl.classList.toggle('is-hidden', !hasContent);
  };

  const setColumnVisibility = (columnEl, hasVisiblePanels) => {
    if (!columnEl) return;
    columnEl.classList.toggle('is-hidden', !hasVisiblePanels);
  };
  
  partnersContainer.innerHTML = '';
  if (workingPartnersContainer) workingPartnersContainer.innerHTML = '';
  employeesContainer.innerHTML = '';
  if (separateCompaniesContainer) separateCompaniesContainer.innerHTML = '';

  const renderPartnerBlockHtml = (p) => `
        <div style="margin-bottom: 1.5rem; padding-bottom: 1rem; border-bottom: 1px dashed var(--border-color);">
          <h4 class="settlement-person-name--large">${getSettlementPersonDisplayNameHtml(p.person, state)}</h4>
          <div style="display: flex; flex-direction: column; gap: 0.3rem; font-size: 0.9rem; margin-top: 0.5rem;">
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Godziny:</span>
              <span>${p.hours.toFixed(1)}h</span>
            </div>
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zarobek (Własne godziny):</span>
              <span>${p.salary.toFixed(2)} zł</span>
            </div>
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zarobek (Podział zysku z pracowników):</span>
              <span>${p.revenueShare.toFixed(2)} zł</span>
            </div>
            ${p.worksShare > 0 ? `
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zarobek (Wykonane Prace):</span>
              <span style="color: var(--primary);">${p.worksShare.toFixed(2)} zł</span>
            </div>
            ` : ''}
            <div style="display: flex; justify-content: space-between; font-weight: 600; margin-top: 0.2rem;">
              <span style="color: var(--text-secondary);">Przychód własny(Brutto):</span>
              <span>${(p.ownGrossAmount || 0).toFixed(2)} zł</span>
            </div>
            ${p.paidCosts > 0 ? `
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zwrot kosztów:</span>
              <span style="color: var(--success);">+${p.paidCosts.toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${p.paidAdvances > 0 ? `
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zwrot zaliczek:</span>
              <span style="color: var(--success);">+${p.paidAdvances.toFixed(2)} zł</span>
            </div>
            ` : ''}
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Udział w kosztach:</span>
              <span style="color: var(--danger);">- ${(p.costShareApplied || 0).toFixed(2)} zł</span>
            </div>
            ${p.advancesTaken > 0 ? `
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zaliczki:</span>
              <span style="color: var(--danger);">-${p.advancesTaken.toFixed(2)} zł</span>
            </div>
            ` : ''}
            <div style="display: flex; justify-content: space-between; font-weight: 600; margin-top: 0.5rem; padding-top: 0.5rem; border-top: 1px solid var(--border-color);">
              <span>Przychód (Brutto):</span>
              <span style="color: ${p.toPayout >= 0 ? 'var(--success)' : 'var(--danger)'};">${p.toPayout.toFixed(2)} zł</span>
            </div>
            ${p.employeeSalaries > 0 ? `
            <div style="display: flex; justify-content: space-between; margin-top: 0.35rem;">
              <span style="color: var(--text-secondary);">Pensje pracowników:</span>
              <span style="color: var(--warning);">${p.employeeSalaries.toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${(p.employeeAccountingRefund || 0) > 0 ? `
            <div style="display: flex; justify-content: space-between; margin-top: 0.35rem;">
              <span style="color: var(--text-secondary);">Zwrot Podatku i ZUS za pracowników:</span>
              <span style="color: var(--warning);">${(p.employeeAccountingRefund || 0).toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${(p.employeeReceivables || 0) > 0 ? `
            <div style="display: flex; justify-content: space-between; margin-top: 0.35rem;">
              <span style="color: var(--text-secondary);">Do odebrania od pracowników:</span>
              <span style="color: var(--danger);">-${p.employeeReceivables.toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${(p.employeeSalaries > 0 || (p.employeeReceivables || 0) > 0) ? `
            <div style="display: flex; justify-content: space-between; font-weight: 600; margin-top: 0.35rem;">
              <span>Przychód (Brutto) z Pensjami:</span>
              <span style="color: ${(p.grossWithEmployeeSalaries || 0) >= 0 ? 'var(--warning)' : 'var(--danger)'};">${(p.grossWithEmployeeSalaries || 0).toFixed(2)} zł</span>
            </div>
            ` : ''}
            <div style="display: flex; justify-content: space-between; background: rgba(239, 68, 68, 0.1); padding: 0.5rem; border-radius: 4px; margin-top: 0.5rem; border: 1px dashed var(--danger);">
              <span style="color: var(--text-secondary);">Koszty ZUS i Podatek:</span>
              <span style="color: var(--danger); font-weight: 500;">
                Podatek własny: ${(p.ownTaxAmount || 0).toFixed(2)} zł<br>
                Podatek wspólny: ${(p.sharedCompanyTaxAmount || 0).toFixed(2)} zł<br>
                ZUS: ${p.zusAmount.toFixed(2)} zł
              </span>
            </div>
            <div style="display: flex; justify-content: space-between; font-weight: 600; margin-top: 0.35rem;">
              <span>Przychód (Netto):</span>
              <span style="color: ${getDisplayedNetAmount(p) >= 0 ? 'var(--success)' : 'var(--danger)'};">${getDisplayedNetAmount(p).toFixed(2)} zł</span>
            </div>
          </div>
        </div>
      `;

  const renderSeparateCompanyBlockHtml = (company) => `
        <div style="margin-bottom: 1.5rem; padding-bottom: 1rem; border-bottom: 1px dashed var(--border-color);">
          <h4 class="settlement-person-name--large">${getSettlementPersonDisplayNameHtml(company.person, state)}</h4>
          <div style="display: flex; flex-direction: column; gap: 0.3rem; font-size: 0.9rem; margin-top: 0.5rem;">
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Godziny:</span>
              <span>${company.hours.toFixed(1)}h${company.effectiveRate > 0 ? ` (${company.effectiveRate.toFixed(2)} zł/h)` : ''}</span>
            </div>
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zarobek (Własne godziny):</span>
              <span>${company.salary.toFixed(2)} zł</span>
            </div>
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zarobek (Podział zysku z pracowników):</span>
              <span>${(company.revenueShare || 0).toFixed(2)} zł</span>
            </div>
            ${company.worksShare > 0 ? `
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zarobek (Wykonane Prace):</span>
              <span style="color: var(--primary);">${company.worksShare.toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${company.paidCosts > 0 ? `
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zwrot kosztów:</span>
              <span style="color: var(--success);">+${company.paidCosts.toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${company.paidAdvances > 0 ? `
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zwrot zaliczek:</span>
              <span style="color: var(--success);">+${company.paidAdvances.toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${(company.costShareApplied || 0) > 0 ? `
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Udział w kosztach:</span>
              <span style="color: var(--danger);">-${(company.costShareApplied || 0).toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${company.advancesTaken > 0 ? `
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zaliczki:</span>
              <span style="color: var(--danger);">-${company.advancesTaken.toFixed(2)} zł</span>
            </div>
            ` : ''}
            <div style="display: flex; justify-content: space-between; font-weight: 600; margin-top: 0.5rem; padding-top: 0.5rem; border-top: 1px solid var(--border-color);">
              <span>Przychód (Brutto):</span>
              <span style="color: ${company.toPayout >= 0 ? 'var(--success)' : 'var(--danger)'};">${company.toPayout.toFixed(2)} zł</span>
            </div>
            ${company.employeeSalaries > 0 ? `
            <div style="display: flex; justify-content: space-between; margin-top: 0.35rem;">
              <span style="color: var(--text-secondary);">Pensje pracowników:</span>
              <span style="color: var(--warning);">${company.employeeSalaries.toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${(company.employeeAccountingRefund || 0) > 0 ? `
            <div style="display: flex; justify-content: space-between; margin-top: 0.35rem;">
              <span style="color: var(--text-secondary);">Zwrot Podatku i ZUS za pracowników:</span>
              <span style="color: var(--warning);">${(company.employeeAccountingRefund || 0).toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${(company.employeeReceivables || 0) > 0 ? `
            <div style="display: flex; justify-content: space-between; margin-top: 0.35rem;">
              <span style="color: var(--text-secondary);">Do odebrania od pracowników:</span>
              <span style="color: var(--danger);">-${company.employeeReceivables.toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${(company.employeeSalaries > 0 || (company.employeeReceivables || 0) > 0 || (company.employeeAccountingRefund || 0) > 0) ? `
            <div style="display: flex; justify-content: space-between; font-weight: 600; margin-top: 0.35rem;">
              <span>Przychód (Brutto) z Pensjami:</span>
              <span style="color: ${(company.grossWithEmployeeSalaries || 0) >= 0 ? 'var(--warning)' : 'var(--danger)'};">${(company.grossWithEmployeeSalaries || 0).toFixed(2)} zł</span>
            </div>
            ` : ''}
          </div>
        </div>
      `;

  const renderWorkingPartnerBlockHtml = (wp) => `
        <div style="margin-bottom: 1.5rem; padding-bottom: 1rem; border-bottom: 1px dashed var(--border-color);">
          <h4 class="settlement-person-name--large">${getSettlementPersonDisplayNameHtml(wp.person, state)}</h4>
          <div style="display: flex; flex-direction: column; gap: 0.3rem; font-size: 0.9rem; margin-top: 0.5rem;">
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Godziny:</span>
              <span>${wp.hours.toFixed(1)}h (${wp.effectiveRate.toFixed(2)} zł/h)</span>
            </div>
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zarobek:</span>
              <span>${wp.salary.toFixed(2)} zł</span>
            </div>
            ${wp.worksShare > 0 ? `
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zarobek (Wykonane Prace):</span>
              <span style="color: var(--primary);">${wp.worksShare.toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${wp.paidCosts > 0 ? `
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zwrot kosztów:</span>
              <span style="color: var(--success);">+${wp.paidCosts.toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${wp.paidAdvances > 0 ? `
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zwrot zaliczek:</span>
              <span style="color: var(--success);">+${wp.paidAdvances.toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${(wp.costShareApplied || 0) > 0 ? `
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Udział w kosztach:</span>
              <span style="color: var(--danger);">-${(wp.costShareApplied || 0).toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${wp.advancesTaken > 0 ? `
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zaliczki pobrane:</span>
              <span style="color: var(--danger);">-${wp.advancesTaken.toFixed(2)} zł</span>
            </div>
            ` : ''}
            <div style="display: flex; justify-content: space-between; font-weight: 600; margin-top: 0.5rem; padding-top: 0.5rem; border-top: 1px solid var(--border-color);">
              <span>Przychód (Brutto):</span>
              <span style="color: ${wp.toPayout >= 0 ? 'var(--success)' : 'var(--danger)'};">${wp.toPayout.toFixed(2)} zł</span>
            </div>
            ${wp.employeeSalaries > 0 ? `
            <div style="display: flex; justify-content: space-between; margin-top: 0.35rem;">
              <span style="color: var(--text-secondary);">Pensje pracowników:</span>
              <span style="color: var(--warning);">${wp.employeeSalaries.toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${(wp.employeeAccountingRefund || 0) > 0 ? `
            <div style="display: flex; justify-content: space-between; margin-top: 0.35rem;">
              <span style="color: var(--text-secondary);">Zwrot Podatku i ZUS za pracowników:</span>
              <span style="color: var(--warning);">${(wp.employeeAccountingRefund || 0).toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${(wp.employeeReceivables || 0) > 0 ? `
            <div style="display: flex; justify-content: space-between; margin-top: 0.35rem;">
              <span style="color: var(--text-secondary);">Do odebrania od pracowników:</span>
              <span style="color: var(--danger);">-${wp.employeeReceivables.toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${(wp.ownTaxAmount || 0) > 0 ? `
            <div style="display: flex; justify-content: space-between; margin-top: 0.35rem;">
              <span style="color: var(--text-secondary);">Podatek własny:</span>
              <span style="color: var(--danger);">-${(wp.ownTaxAmount || 0).toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${(wp.contractTaxAmount || 0) > 0 ? `
            <div style="display: flex; justify-content: space-between; margin-top: 0.35rem;">
              <span style="color: var(--text-secondary);">Podatek UZ${wp.contractChargesPaidByEmployer ? ' (opłaca pracodawca)' : ''}:</span>
              <span style="color: ${wp.contractChargesPaidByEmployer ? 'var(--warning)' : 'var(--danger)'};">${wp.contractChargesPaidByEmployer ? '' : '-'}${(wp.contractTaxAmount || 0).toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${(wp.contractZusAmount || 0) > 0 ? `
            <div style="display: flex; justify-content: space-between; margin-top: 0.35rem;">
              <span style="color: var(--text-secondary);">ZUS UZ${wp.contractChargesPaidByEmployer ? ' (opłaca pracodawca)' : ''}:</span>
              <span style="color: ${wp.contractChargesPaidByEmployer ? 'var(--warning)' : 'var(--danger)'};">${wp.contractChargesPaidByEmployer ? '' : '-'}${(wp.contractZusAmount || 0).toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${(wp.employeeSalaries > 0 || (wp.employeeReceivables || 0) > 0) ? `
            <div style="display: flex; justify-content: space-between; font-weight: 600; margin-top: 0.35rem;">
              <span>Przychód (Brutto) z Pensjami:</span>
              <span style="color: ${(wp.grossWithEmployeeSalaries || 0) >= 0 ? 'var(--warning)' : 'var(--danger)'};">${(wp.grossWithEmployeeSalaries || 0).toFixed(2)} zł</span>
            </div>
            ` : ''}
            <div style="display: flex; justify-content: space-between; font-weight: 600; margin-top: 0.35rem;">
              <span>Przychód (Netto):</span>
              <span style="color: ${getDisplayedNetAmount(wp) >= 0 ? 'var(--success)' : 'var(--danger)'};">${getDisplayedNetAmount(wp).toFixed(2)} zł</span>
            </div>
          </div>
        </div>
      `;

  const renderEmployeeBlockHtml = (e) => {
    const employeeGeneratedProfit = getEmployeeGeneratedProfitDisplay(e);
    return `
        <div style="margin-bottom: 1.5rem; padding-bottom: 1rem; border-bottom: 1px dashed var(--border-color);">
          <h4 class="settlement-person-name--large">${getSettlementPersonDisplayNameHtml(e.person, state)}</h4>
          <div style="display: flex; flex-direction: column; gap: 0.3rem; font-size: 0.9rem; margin-top: 0.5rem;">
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Godziny:</span>
              <span>${e.hours.toFixed(1)}h (${e.effectiveRate.toFixed(2)} zł/h)</span>
            </div>
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zarobek:</span>
              <span>${e.salary.toFixed(2)} zł</span>
            </div>
            ${(e.bonusAmount || 0) > 0 ? `
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Premia:</span>
              <span style="color: #fbbf24;">+${(e.bonusAmount || 0).toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${e.paidCosts > 0 ? `
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zwrot kosztów:</span>
              <span style="color: var(--success);">+${e.paidCosts.toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${e.paidAdvances > 0 ? `
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zwrot zaliczek:</span>
              <span style="color: var(--success);">+${e.paidAdvances.toFixed(2)} zł</span>
            </div>
            ` : ''}
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zaliczki pobrane:</span>
              <span style="color: var(--danger);">-${e.advancesTaken.toFixed(2)} zł</span>
            </div>
            ${(e.contractTaxAmount || 0) > 0 ? `
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Podatek UZ${e.contractChargesPaidByEmployer ? ' (opłaca pracodawca)' : ''}:</span>
              <span style="color: ${e.contractChargesPaidByEmployer ? 'var(--warning)' : 'var(--danger)'};">${e.contractChargesPaidByEmployer ? '' : '-'}${(e.contractTaxAmount || 0).toFixed(2)} zł</span>
            </div>
            ` : ''}
            ${(e.contractZusAmount || 0) > 0 ? `
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">ZUS UZ${e.contractChargesPaidByEmployer ? ' (opłaca pracodawca)' : ''}:</span>
              <span style="color: ${e.contractChargesPaidByEmployer ? 'var(--warning)' : 'var(--danger)'};">${e.contractChargesPaidByEmployer ? '' : '-'}${(e.contractZusAmount || 0).toFixed(2)} zł</span>
            </div>
            ` : ''}
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">${employeeGeneratedProfit.label}</span>
              <span style="color: #2563eb;">${employeeGeneratedProfit.valueText}</span>
            </div>
            <div style="display: flex; justify-content: space-between; font-weight: 600; margin-top: 0.5rem; padding-top: 0.5rem; border-top: 1px solid var(--border-color);">
              <span>Do Wypłaty:</span>
              <span style="color: ${e.toPayout >= 0 ? 'var(--success)' : 'var(--danger)'};">${e.toPayout.toFixed(2)} zł</span>
            </div>
          </div>
        </div>
    `;
  };

  if (res.partners.length > 0) {
    partnersContainer.innerHTML = res.partners.map(renderPartnerBlockHtml).join('');
  }

  if (workingPartnersContainer && res.workingPartners.length > 0) {
    workingPartnersContainer.innerHTML = res.workingPartners.map(renderWorkingPartnerBlockHtml).join('');
  }

  if (employeesContainer && res.employees.length > 0) {
    employeesContainer.innerHTML = res.employees.map(renderEmployeeBlockHtml).join('');
  }

  if (separateCompaniesContainer && res.separateCompanies.length > 0) {
    separateCompaniesContainer.innerHTML = res.separateCompanies.map(renderSeparateCompanyBlockHtml).join('');
  }

  setPanelVisibility(partnersPanel, res.partners.length > 0);
  setPanelVisibility(workingPartnersPanel, res.workingPartners.length > 0);
  setPanelVisibility(employeesPanel, res.employees.length > 0);
  setPanelVisibility(separateCompaniesPanel, res.separateCompanies.length > 0);
  setColumnVisibility(leftColumn, res.partners.length > 0 || res.separateCompanies.length > 0);
  setColumnVisibility(rightColumn, res.workingPartners.length > 0 || res.employees.length > 0);

  renderSettlementDetails(state, res);
}

// ==========================================
// INVOICES VIEW
// ==========================================
function copyTextToClipboard(value) {
  if (!value) return Promise.resolve();
  if (navigator.clipboard && navigator.clipboard.writeText) {
    return navigator.clipboard.writeText(value);
  }

  const temp = document.createElement('textarea');
  temp.value = value;
  document.body.appendChild(temp);
  temp.select();
  document.execCommand('copy');
  temp.remove();
  return Promise.resolve();
}

function getInvoiceExtraEligibleIssuers(state) {
  return Calculations.getInvoiceEligibleIssuers(state, getSelectedMonthKey())
    .filter(person => person.type === 'PARTNER' || person.type === 'WORKING_PARTNER');
}

function getInvoiceExtraClientDisplayName(invoice, state) {
  if (!invoice) return '';
  if (invoice.clientId) {
    const client = (state?.clients || []).find(item => item.id === invoice.clientId);
    if (client?.name) return client.name;
  }
  return (invoice.clientName || '').toString().trim();
}

function resetInvoiceExtraForm() {
  const form = document.getElementById('invoice-extra-form');
  const idInput = document.getElementById('invoice-extra-id');
  const clientSelect = document.getElementById('invoice-extra-client-select');
  const clientNameInput = document.getElementById('invoice-extra-client-name');
  const issuerSelect = document.getElementById('invoice-extra-issuer');
  const amountInput = document.getElementById('invoice-extra-amount');
  const saveButton = document.getElementById('btn-save-invoice-extra');
  const cancelButton = document.getElementById('btn-cancel-invoice-extra');

  if (form) form.reset();
  if (idInput) idInput.value = '';
  if (clientSelect) clientSelect.value = '';
  if (clientNameInput) clientNameInput.value = '';
  if (issuerSelect) issuerSelect.value = '';
  if (amountInput) amountInput.value = '';
  if (saveButton) saveButton.textContent = 'Dodaj';
  if (cancelButton) cancelButton.style.display = 'none';
}

function setInvoiceExtraPanelCollapsed(isCollapsed) {
  const content = document.getElementById('invoice-extra-panel-content');
  const icon = document.getElementById('invoice-extra-panel-icon');
  const toggleButton = document.getElementById('btn-toggle-invoice-extra-panel');
  if (!content || !icon || !toggleButton) return;

  content.style.display = isCollapsed ? 'none' : 'block';
  icon.setAttribute('data-lucide', isCollapsed ? 'chevron-right' : 'chevron-down');
  toggleButton.dataset.collapsed = isCollapsed ? 'true' : 'false';
  lucide.createIcons();
}

function normalizeInvoicePercentageAllocationsForCard(issuerIds, percentageAllocations = {}) {
  if (!Array.isArray(issuerIds) || issuerIds.length === 0) return {};

  const parsedAllocations = {};
  let total = 0;
  issuerIds.forEach(issuerId => {
    const value = Math.max(0, Math.min(100, Math.round(parseFloat(percentageAllocations?.[issuerId]) || 0)));
    parsedAllocations[issuerId] = value;
    total += value;
  });

  if (!(total > 0)) {
    const equalShare = 100 / issuerIds.length;
    const normalized = {};
    let assigned = 0;
    issuerIds.forEach((issuerId, index) => {
      if (index === issuerIds.length - 1) {
        normalized[issuerId] = Math.max(0, 100 - assigned);
        return;
      }
      normalized[issuerId] = Math.round(equalShare);
      assigned += normalized[issuerId];
    });
    return normalized;
  }

  const normalized = {};
  let assigned = 0;
  issuerIds.forEach((issuerId, index) => {
    if (index === issuerIds.length - 1) {
      normalized[issuerId] = Math.max(0, 100 - assigned);
      return;
    }
    normalized[issuerId] = Math.round((parsedAllocations[issuerId] / total) * 100);
    assigned += normalized[issuerId];
  });

  return normalized;
}

function spreadInvoicePercentageDelta(percentages, issuerIds, amount, subtract = true) {
  let remaining = Math.max(0, amount);

  issuerIds.forEach(issuerId => {
    if (!(remaining > 0)) return;
    const currentValue = percentages[issuerId] || 0;
    const capacity = subtract ? currentValue : (100 - currentValue);
    const applied = Math.max(0, Math.min(capacity, remaining));
    percentages[issuerId] = Math.round(currentValue + (subtract ? -applied : applied));
    remaining -= applied;
  });

  return amount - Math.max(0, remaining);
}

function finalizeInvoicePercentageAllocations(issuerIds, percentageAllocations, priorityIssuerIds = [], fallbackIssuerId = '') {
  const result = Object.fromEntries(issuerIds.map(issuerId => [issuerId, Math.max(0, Math.min(100, Math.round(percentageAllocations[issuerId] || 0)))]));
  let remainder = 100 - issuerIds.reduce((sum, issuerId) => sum + (result[issuerId] || 0), 0);

  const orderedIssuerIds = [
    ...priorityIssuerIds.filter(issuerId => issuerIds.includes(issuerId)),
    ...issuerIds.filter(issuerId => !priorityIssuerIds.includes(issuerId) && issuerId !== fallbackIssuerId),
    ...(fallbackIssuerId && issuerIds.includes(fallbackIssuerId) ? [fallbackIssuerId] : [])
  ];

  for (const issuerId of orderedIssuerIds) {
    if (remainder === 0) break;
    if (remainder > 0) {
      const capacity = 100 - (result[issuerId] || 0);
      const applied = Math.min(capacity, remainder);
      result[issuerId] += applied;
      remainder -= applied;
    } else {
      const capacity = result[issuerId] || 0;
      const applied = Math.min(capacity, Math.abs(remainder));
      result[issuerId] -= applied;
      remainder += applied;
    }
  }

  return result;
}

function rebalanceInvoicePercentageAllocations(issuerIds, percentageAllocations, editedIssuerId, nextValue, lastEditedIssuerId = '', touchedIssuerIds = []) {
  if (!Array.isArray(issuerIds) || issuerIds.length === 0) {
    return { percentageAllocations: {}, percentageTouchedIssuerIds: [], lastEditedPercentageIssuerId: '' };
  }

  if (issuerIds.length === 1) {
    return {
      percentageAllocations: { [issuerIds[0]]: 100 },
      percentageTouchedIssuerIds: [issuerIds[0]],
      lastEditedPercentageIssuerId: issuerIds[0]
    };
  }

  const normalized = normalizeInvoicePercentageAllocationsForCard(issuerIds, percentageAllocations);
  const touchedOrder = (touchedIssuerIds || []).filter(issuerId => issuerIds.includes(issuerId));
  const touchedSet = new Set(touchedOrder);
  const clampedTargetValue = Math.max(0, Math.min(100, parseFloat(nextValue) || 0));
  const currentTargetValue = normalized[editedIssuerId] || 0;
  const requestedDelta = clampedTargetValue - currentTargetValue;
  const otherIssuerIds = issuerIds.filter(issuerId => issuerId !== editedIssuerId);

  if (Math.abs(requestedDelta) < 0.0001) {
    const nextTouchedOrder = [...touchedOrder.filter(issuerId => issuerId !== editedIssuerId), editedIssuerId];
    const protectedIssuerId = lastEditedIssuerId && lastEditedIssuerId !== editedIssuerId && otherIssuerIds.includes(lastEditedIssuerId)
      ? lastEditedIssuerId
      : '';
    const priorityIssuerIds = [
      ...touchedOrder.filter(issuerId => issuerId !== editedIssuerId && issuerId !== protectedIssuerId && otherIssuerIds.includes(issuerId)),
      ...otherIssuerIds.filter(issuerId => !touchedOrder.includes(issuerId) && issuerId !== protectedIssuerId),
      editedIssuerId
    ];
    return {
      percentageAllocations: finalizeInvoicePercentageAllocations(issuerIds, normalized, priorityIssuerIds, protectedIssuerId),
      percentageTouchedIssuerIds: nextTouchedOrder,
      lastEditedPercentageIssuerId: editedIssuerId
    };
  }

  let remainingDelta = Math.abs(requestedDelta);
  const protectedIssuerId = lastEditedIssuerId && lastEditedIssuerId !== editedIssuerId && otherIssuerIds.includes(lastEditedIssuerId)
    ? lastEditedIssuerId
    : '';
  const queuedEditedIssuerIds = touchedOrder.filter(issuerId => issuerId !== editedIssuerId && issuerId !== protectedIssuerId && otherIssuerIds.includes(issuerId));
  const untouchedIssuerIds = otherIssuerIds.filter(issuerId => !touchedSet.has(issuerId));

  remainingDelta -= spreadInvoicePercentageDelta(normalized, queuedEditedIssuerIds, remainingDelta, requestedDelta > 0);
  if (remainingDelta > 0.0001) {
    remainingDelta -= spreadInvoicePercentageDelta(normalized, untouchedIssuerIds, remainingDelta, requestedDelta > 0);
  }
  if (remainingDelta > 0.0001 && protectedIssuerId) {
    remainingDelta -= spreadInvoicePercentageDelta(normalized, [protectedIssuerId], remainingDelta, requestedDelta > 0);
  }

  const actualDelta = Math.sign(requestedDelta) * (Math.abs(requestedDelta) - Math.max(0, remainingDelta));
  normalized[editedIssuerId] = Math.round(currentTargetValue + actualDelta);
  const nextTouchedOrder = [...touchedOrder.filter(issuerId => issuerId !== editedIssuerId), editedIssuerId];
  const priorityIssuerIds = [
    ...queuedEditedIssuerIds,
    ...untouchedIssuerIds,
    editedIssuerId
  ];

  return {
    percentageAllocations: finalizeInvoicePercentageAllocations(issuerIds, normalized, priorityIssuerIds, protectedIssuerId),
    percentageTouchedIssuerIds: nextTouchedOrder,
    lastEditedPercentageIssuerId: editedIssuerId
  };
}

function getInvoiceClientExistingConfig(clientId) {
  const selectedMonth = getSelectedMonthKey();
  const invoiceSettings = Store.getMonthSettings(selectedMonth)?.invoices || { clients: {} };
  return invoiceSettings.clients?.[clientId] || {};
}

function updateInvoicePercentageControls(card, percentageAllocations = {}) {
  if (!card) return;
  const netRevenue = parseFloat(card.getAttribute('data-net-revenue')) || 0;

  card.querySelectorAll('.invoice-percentage-slider').forEach(input => {
    const issuerId = input.getAttribute('data-issuer-id');
    const percentValue = Math.max(0, Math.min(100, Math.round(parseFloat(percentageAllocations?.[issuerId]) || 0)));
    input.value = percentValue;

    const wrapper = input.closest('.invoice-percentage-controls');
    if (!wrapper) return;

    const valueEl = wrapper.querySelector('.invoice-percentage-value');
    const amountEl = wrapper.querySelector('.invoice-percentage-amount');
    if (valueEl) valueEl.textContent = `${percentValue}%`;
    if (amountEl) amountEl.textContent = Calculations.formatInvoiceCurrency(netRevenue * (percentValue / 100));
  });
}

function saveInvoiceClientConfigFromCard(card, overrides = {}) {
  if (!card) return;
  const clientId = card.getAttribute('data-client-id');
  if (!clientId) return;
  const existingConfig = getInvoiceClientExistingConfig(clientId);

  const mode = card.querySelector('.invoice-client-mode')?.value || 'SETTLEMENT_REVENUE';
  const issuerIds = Array.from(card.querySelectorAll('.invoice-issuer-toggle:checked')).map(input => input.value);
  const deductClientAdvances = card.querySelector('.invoice-deduct-advances-toggle')?.checked !== false;
  const includeClientCosts = mode === 'SETTLEMENT_REVENUE'
    && card.querySelector('.invoice-include-costs-toggle')?.checked !== false;
  const randomizeEqualSplitInvoices = mode === 'EQUAL_SPLIT'
    && card.querySelector('.invoice-randomize-equal-split-toggle')?.checked === true;
  const equalSplitVarianceAmount = mode === 'EQUAL_SPLIT'
    ? Math.max(0, parseFloat(card.querySelector('.invoice-equal-split-variance-amount')?.value) || 10)
    : 10;
  const manualAmounts = {};
  const separateCompanyWithEmployees = {};
  card.querySelectorAll('.invoice-manual-amount').forEach(input => {
    const amount = parseFloat(input.value);
    if (Number.isFinite(amount)) {
      manualAmounts[input.getAttribute('data-issuer-id')] = amount;
    }
  });
  card.querySelectorAll('.invoice-company-with-employees-toggle').forEach(input => {
    separateCompanyWithEmployees[input.getAttribute('data-issuer-id')] = input.checked === true;
  });
  const percentageEligibleIssuerIds = issuerIds.filter(issuerId => separateCompanyWithEmployees[issuerId] !== true);
  const rawPercentageAllocations = { ...(existingConfig.percentageAllocations || {}) };
  card.querySelectorAll('.invoice-percentage-slider').forEach(input => {
    rawPercentageAllocations[input.getAttribute('data-issuer-id')] = parseFloat(input.value) || 0;
  });
  issuerIds.forEach(issuerId => {
    if (!percentageEligibleIssuerIds.includes(issuerId)) {
      rawPercentageAllocations[issuerId] = 0;
    }
  });
  const normalizedEligiblePercentageAllocations = normalizeInvoicePercentageAllocationsForCard(percentageEligibleIssuerIds, rawPercentageAllocations);
  const percentageAllocations = Object.fromEntries(issuerIds.map(issuerId => [issuerId, normalizedEligiblePercentageAllocations[issuerId] || 0]));
  const percentageTouchedIssuerIds = (overrides.percentageTouchedIssuerIds || existingConfig.percentageTouchedIssuerIds || [])
    .filter(issuerId => percentageEligibleIssuerIds.includes(issuerId));
  const lastEditedPercentageIssuerId = percentageEligibleIssuerIds.includes(overrides.lastEditedPercentageIssuerId || existingConfig.lastEditedPercentageIssuerId)
    ? (overrides.lastEditedPercentageIssuerId || existingConfig.lastEditedPercentageIssuerId)
    : '';
  const notes = card.querySelector('.invoice-client-notes')?.value.trim() || '';

  Store.updateInvoiceClientConfig(clientId, {
    mode,
    issuerIds,
    deductClientAdvances,
    includeClientCosts,
    randomizeEqualSplitInvoices,
    equalSplitVarianceAmount,
    percentageAllocations,
    percentageTouchedIssuerIds,
    lastEditedPercentageIssuerId,
    manualAmounts,
    separateCompanyWithEmployees,
    notes,
    ...overrides,
    percentageAllocations: overrides.percentageAllocations || percentageAllocations
  });
  renderInvoices();
}

function initInvoices() {
  const issueDateInput = document.getElementById('invoices-issue-date');
  const emailIntroInput = document.getElementById('invoices-email-intro');
  const copyButton = document.getElementById('btn-copy-invoices-email');
  const extraForm = document.getElementById('invoice-extra-form');
  const extraClientSelect = document.getElementById('invoice-extra-client-select');
  const extraClientNameInput = document.getElementById('invoice-extra-client-name');
  const extraCancelButton = document.getElementById('btn-cancel-invoice-extra');
  const extraPanelToggleButton = document.getElementById('btn-toggle-invoice-extra-panel');

  if (issueDateInput) {
    issueDateInput.addEventListener('change', () => {
      Store.updateInvoiceMonthConfig({ issueDate: issueDateInput.value, emailIntro: emailIntroInput ? emailIntroInput.value.trim() : '' });
    });
  }

  if (extraPanelToggleButton) {
    extraPanelToggleButton.addEventListener('click', () => {
      const isCollapsed = extraPanelToggleButton.dataset.collapsed === 'true';
      setInvoiceExtraPanelCollapsed(!isCollapsed);
    });
  }

  if (emailIntroInput) {
    emailIntroInput.addEventListener('change', () => {
      Store.updateInvoiceMonthConfig({ issueDate: issueDateInput ? issueDateInput.value : '', emailIntro: emailIntroInput.value.trim() });
    });
  }

  if (copyButton) {
    copyButton.addEventListener('click', async () => {
      const emailText = document.getElementById('invoices-email-text')?.value || '';
      if (!emailText) return;
      try {
        await copyTextToClipboard(emailText);
      } catch {
        alert('Nie udało się skopiować treści maila.');
      }
    });
  }

  if (extraClientSelect && extraClientNameInput) {
    extraClientSelect.addEventListener('change', () => {
      if (extraClientSelect.value) {
        extraClientNameInput.value = '';
      }
    });
  }

  if (extraCancelButton) {
    extraCancelButton.addEventListener('click', () => {
      resetInvoiceExtraForm();
    });
  }

  if (extraForm) {
    extraForm.addEventListener('submit', (e) => {
      e.preventDefault();

      const extraId = document.getElementById('invoice-extra-id')?.value || '';
      const selectedClientId = extraClientSelect?.value || '';
      const manualClientName = extraClientNameInput?.value.trim() || '';
      const issuerId = document.getElementById('invoice-extra-issuer')?.value || '';
      const amount = parseFloat(document.getElementById('invoice-extra-amount')?.value);
      const state = Store.getState();
      const selectedClient = selectedClientId ? state.clients.find(client => client.id === selectedClientId) : null;
      const clientName = (selectedClient?.name || manualClientName || '').trim();

      if (!clientName) {
        alert('Wybierz klienta z listy lub wpisz jego nazwę ręcznie.');
        return;
      }

      if (!issuerId) {
        alert('Wybierz wspólnika wystawiającego fakturę.');
        return;
      }

      if (!Number.isFinite(amount) || amount <= 0) {
        alert('Podaj kwotę faktury większą od zera.');
        return;
      }

      const payload = {
        clientId: selectedClient?.id || '',
        clientName,
        issuerId,
        amount
      };

      if (extraId) {
        Store.updateInvoiceExtraInvoice(extraId, payload);
      } else {
        Store.addInvoiceExtraInvoice(payload);
      }

      resetInvoiceExtraForm();
    });
  }
}

function renderInvoices() {
  const subtitle = document.getElementById('invoices-subtitle');
  const totalRevenueEl = document.getElementById('invoices-total-revenue');
  const totalIssuedEl = document.getElementById('invoices-total-issued');
  const totalDifferenceEl = document.getElementById('invoices-total-difference');
  const issueDateInput = document.getElementById('invoices-issue-date');
  const emailIntroInput = document.getElementById('invoices-email-intro');
  const configList = document.getElementById('invoices-config-list');
  const issuerSummary = document.getElementById('invoices-issuer-summary');
  const clientSummary = document.getElementById('invoices-client-summary');
  const emailText = document.getElementById('invoices-email-text');
  const extraClientSelect = document.getElementById('invoice-extra-client-select');
  const extraIssuerSelect = document.getElementById('invoice-extra-issuer');
  const extraList = document.getElementById('invoice-extra-list');
  const extraPanelToggleButton = document.getElementById('btn-toggle-invoice-extra-panel');
  if (!configList || !issuerSummary || !clientSummary || !emailText || !extraClientSelect || !extraIssuerSelect || !extraList || !extraPanelToggleButton) return;

  const state = Store.getState();
  const selectedMonth = getSelectedMonthKey();
  const invoiceData = Calculations.calculateInvoices(state, selectedMonth);
  const monthSettings = Store.getMonthSettings(selectedMonth);
  const invoiceSettings = monthSettings?.invoices || { issueDate: '', emailIntro: '', clients: {} };
  const extraEligibleIssuers = getInvoiceExtraEligibleIssuers(state);

  extraClientSelect.innerHTML = '<option value="">-- Wybierz klienta --</option>'
    + (state.clients || []).map(client => `<option value="${client.id}">${client.name}</option>`).join('');
  extraIssuerSelect.innerHTML = '<option value="">-- Wybierz wspólnika --</option>'
    + extraEligibleIssuers.map(person => `<option value="${person.id}">${getPersonDisplayName(person)}</option>`).join('');

  if (subtitle) {
    subtitle.textContent = `Podział faktur i mail do księgowej dla ${formatMonthLabel(selectedMonth)}`;
  }

  if (totalRevenueEl) totalRevenueEl.textContent = Calculations.formatInvoiceCurrency(invoiceData.totalRevenue);
  if (totalIssuedEl) totalIssuedEl.textContent = Calculations.formatInvoiceCurrency(invoiceData.totalInvoices);
  if (totalDifferenceEl) {
    totalDifferenceEl.textContent = Calculations.formatInvoiceCurrency(invoiceData.difference);
    totalDifferenceEl.style.color = Math.abs(invoiceData.difference) < 0.005 ? 'var(--success)' : 'var(--danger)';
  }

  if (issueDateInput) issueDateInput.value = invoiceData.issueDate;
  if (emailIntroInput) emailIntroInput.value = invoiceSettings.emailIntro || '';
  emailText.value = invoiceData.emailText || '';

  const extraInvoices = invoiceData.extraInvoices || [];
  setInvoiceExtraPanelCollapsed(extraInvoices.length === 0);
  if (extraInvoices.length === 0) {
    extraList.innerHTML = '<tr><td colspan="4" style="text-align:center;color:var(--text-muted)">Brak dodatkowych faktur poza rozliczeniem.</td></tr>';
  } else {
    extraList.innerHTML = extraInvoices.map(invoice => `
      <tr>
        <td>${getInvoiceExtraClientDisplayName(invoice, state)}</td>
        <td>${invoice.issuerName}</td>
        <td>${Calculations.formatInvoiceCurrency(invoice.amount)}</td>
        <td>
          <div style="display:flex; gap:0.5rem;">
            <button class="btn btn-secondary btn-icon btn-edit-invoice-extra" data-id="${invoice.id}">
              <i data-lucide="edit-2" style="width:16px;height:16px"></i>
            </button>
            <button class="btn btn-danger btn-icon btn-delete-invoice-extra" data-id="${invoice.id}">
              <i data-lucide="trash-2" style="width:16px;height:16px"></i>
            </button>
          </div>
        </td>
      </tr>
    `).join('');
  }

  if (invoiceData.clientInvoices.length === 0) {
    configList.innerHTML = '<p style="color: var(--text-muted);">Brak przychodów klientów w wybranym miesiącu. Najpierw dodaj arkusze godzin lub prace.</p>';
    issuerSummary.innerHTML = invoiceData.issuerSummaries.length > 0
      ? invoiceData.issuerSummaries.map(summary => `
        <div class="settlement-detail-row" style="padding: 0.45rem 0; border-bottom: 1px solid var(--border-color);">
          <div>
            <span>${summary.issuerName}</span>
            ${summary.extraInvoicesAmount > 0 ? `<div style="font-size:0.78rem; color: var(--text-secondary); margin-top:0.15rem;">W tym poza rozliczeniem: ${Calculations.formatInvoiceCurrency(summary.extraInvoicesAmount)}</div>` : ''}
          </div>
          <div style="text-align:right;">
            <strong>${Calculations.formatInvoiceCurrency(summary.totalAmount)}</strong>
            ${summary.issuerType !== 'SEPARATE_COMPANY' ? `<div style="font-size:0.78rem; color: var(--text-secondary); margin-top:0.15rem;">Podatek: ${Calculations.formatInvoiceCurrency(summary.taxAmount || 0)}</div>` : ''}
            ${summary.issuerType !== 'SEPARATE_COMPANY' ? `<div style="font-size:0.78rem; color: var(--text-secondary); margin-top:0.15rem;">Suma wszystkich faktur: ${Calculations.formatInvoiceCurrency(summary.yearToDateTotal || 0)}</div>` : ''}
          </div>
        </div>
      `).join('')
      : '<p style="color: var(--text-muted);">Brak danych do podsumowania.</p>';
    clientSummary.innerHTML = '<p style="color: var(--text-muted);">Brak danych do podsumowania.</p>';
    document.querySelectorAll('.btn-edit-invoice-extra').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const id = e.currentTarget.getAttribute('data-id');
        const invoice = (invoiceSettings.extraInvoices || []).find(item => item.id === id);
        if (!invoice) return;

        setInvoiceExtraPanelCollapsed(false);
        document.getElementById('invoice-extra-id').value = invoice.id;
        document.getElementById('invoice-extra-client-select').value = invoice.clientId || '';
        document.getElementById('invoice-extra-client-name').value = invoice.clientId ? '' : (invoice.clientName || '');
        document.getElementById('invoice-extra-issuer').value = invoice.issuerId || '';
        document.getElementById('invoice-extra-amount').value = invoice.amount || '';
        document.getElementById('btn-save-invoice-extra').textContent = 'Zapisz';
        document.getElementById('btn-cancel-invoice-extra').style.display = 'inline-flex';
        document.getElementById('invoice-extra-form').scrollIntoView({ behavior: 'smooth', block: 'center' });
      });
    });

    document.querySelectorAll('.btn-delete-invoice-extra').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const id = e.currentTarget.getAttribute('data-id');
        if (confirm('Usunąć tę dodatkową fakturę?')) {
          Store.deleteInvoiceExtraInvoice(id);
          resetInvoiceExtraForm();
        }
      });
    });

    lucide.createIcons();
    return;
  }

  configList.innerHTML = invoiceData.clientInvoices.map(clientInvoice => {
    const clientConfig = invoiceSettings.clients?.[clientInvoice.clientId] || {};
    const selectedIssuerIds = Array.isArray(clientConfig.issuerIds) && clientConfig.issuerIds.length > 0
      ? clientConfig.issuerIds
      : clientInvoice.allocations.map(allocation => allocation.issuerId);
    const isSettlementRevenueMode = clientInvoice.mode === 'SETTLEMENT_REVENUE';
    const isEqualSplitMode = clientInvoice.mode === 'EQUAL_SPLIT';
    const isPercentageMode = clientInvoice.mode === 'PERCENTAGE_SPLIT';

    return `
      <div class="glass-panel invoice-client-card" data-client-id="${clientInvoice.clientId}" data-net-revenue="${clientInvoice.netRevenue}" style="padding: 1rem; border-radius: 14px;">
        <div style="display:flex; justify-content:space-between; gap:1rem; align-items:flex-start; flex-wrap:wrap; margin-bottom: 0.75rem;">
          <div>
            <h4 style="margin:0; color: var(--text-primary);">${clientInvoice.clientName}</h4>
            <div style="color: var(--text-secondary); font-size: 0.82rem; margin-top: 0.25rem;">
              ${clientInvoice.fullCompanyName ? `Pełna nazwa: ${clientInvoice.fullCompanyName}<br>` : ''}
              ${clientInvoice.nip ? `NIP: ${clientInvoice.nip}` : 'Brak danych rejestrowych klienta'}
            </div>
          </div>
          <div style="text-align:right; min-width: 180px;">
            <div style="color: var(--text-secondary); font-size: 0.8rem;">Przychód klienta z arkuszy</div>
            <div style="font-size:1.1rem; font-weight:700; color: var(--success);">${Calculations.formatInvoiceCurrency(clientInvoice.netRevenue)}</div>
            ${clientInvoice.deductClientAdvances && (clientInvoice.clientAdvances || 0) > 0 ? `<div style="font-size:0.8rem; color: var(--warning); margin-top: 0.25rem;">Arkusze: ${Calculations.formatInvoiceCurrency(clientInvoice.totalRevenue)}<br>Zaliczki: -${Calculations.formatInvoiceCurrency(clientInvoice.clientAdvances || 0)}</div>` : ''}
          </div>
        </div>
        <div style="display:grid; grid-template-columns: minmax(220px, 260px) 1fr; gap: 1rem; align-items:start; margin-bottom: 0.85rem;">
          <div class="form-group" style="margin-bottom: 0;">
            <label>Tryb podziału</label>
            <select class="invoice-client-mode">
              <option value="SETTLEMENT_REVENUE" ${clientInvoice.mode === 'SETTLEMENT_REVENUE' ? 'selected' : ''}>Według przychodów z rozliczenia</option>
              <option value="EQUAL_SPLIT" ${clientInvoice.mode === 'EQUAL_SPLIT' ? 'selected' : ''}>Po równo między zaznaczonych</option>
              <option value="PERCENTAGE_SPLIT" ${clientInvoice.mode === 'PERCENTAGE_SPLIT' ? 'selected' : ''}>Procentowo między zaznaczonych</option>
              <option value="BALANCE_INVOICE_SUMS" ${clientInvoice.mode === 'BALANCE_INVOICE_SUMS' ? 'selected' : ''}>Wyrównywanie Sumy Faktur</option>
              <option value="MANUAL" ${clientInvoice.mode === 'MANUAL' ? 'selected' : ''}>Ręczne kwoty</option>
              <option value="OWN_REVENUE" ${clientInvoice.mode === 'OWN_REVENUE' ? 'selected' : ''}>Według przychodów z arkuszy</option>
            </select>
          </div>
          <div class="form-group" style="margin-bottom: 0;">
            <label>Uwagi dla klienta / księgowej</label>
            <input type="text" class="invoice-client-notes" value="${(clientConfig.notes || '').replace(/"/g, '&quot;')}" placeholder="Np. faktury na spółkę tylko od wybranych wspólników">
          </div>
        </div>
        <div style="margin-bottom: 0.85rem; display:flex; justify-content:space-between; gap:1rem; align-items:center; flex-wrap:wrap;">
          <label style="display:flex; align-items:center; gap:0.6rem; cursor:pointer; color: var(--text-secondary); margin: 0;">
            <input type="checkbox" class="invoice-deduct-advances-toggle" ${clientInvoice.deductClientAdvances ? 'checked' : ''} style="width:auto;">
            Odlicz zaliczki klienta od Przychodu klienta z arkuszy
          </label>
          ${(clientInvoice.clientAdvances || 0) > 0 ? `<div style="font-size:0.82rem; color: var(--text-secondary);">Zaliczki klienta w miesiącu: <strong style="color: var(--warning);">${Calculations.formatInvoiceCurrency(clientInvoice.clientAdvances || 0)}</strong></div>` : ''}
        </div>
        <div class="invoice-include-costs-row" style="margin-bottom: 0.85rem; display:${isSettlementRevenueMode ? 'flex' : 'none'}; justify-content:space-between; gap:1rem; align-items:center; flex-wrap:wrap;">
          <label style="display:flex; align-items:center; gap:0.6rem; cursor:pointer; color: var(--text-secondary); margin: 0;">
            <input type="checkbox" class="invoice-include-costs-toggle" ${clientInvoice.includeClientCosts !== false ? 'checked' : ''} style="width:auto;">
            Czy liczyć koszty w fakturach dla tego klienta
          </label>
          ${clientInvoice.includeClientCosts !== false && (clientInvoice.clientCostShareRatio || 0) > 0 ? `<div style="font-size:0.82rem; color: var(--text-secondary);">Udział kosztów w fakturach: <strong style="color: var(--warning);">${((clientInvoice.clientCostShareRatio || 0) * 100).toFixed(2)}%</strong></div>` : ''}
        </div>
        <div class="invoice-randomize-equal-split-row" style="margin-bottom: 0.85rem; display:${isEqualSplitMode ? 'flex' : 'none'}; justify-content:space-between; gap:1rem; align-items:center; flex-wrap:wrap;">
          <label style="display:flex; align-items:center; gap:0.6rem; cursor:pointer; color: var(--text-secondary); margin: 0;">
            <input type="checkbox" class="invoice-randomize-equal-split-toggle" ${clientInvoice.randomizeEqualSplitInvoices === true ? 'checked' : ''} style="width:auto;">
            Zróżnicuj faktury o:
          </label>
          <div style="display:flex; align-items:center; gap:0.5rem;">
            <input type="number" step="0.01" min="0" class="invoice-equal-split-variance-amount" value="${Number.isFinite(parseFloat(clientInvoice.equalSplitVarianceAmount)) ? parseFloat(clientInvoice.equalSplitVarianceAmount) : 10}" style="width:120px;">
            <span style="color: var(--text-secondary); font-size:0.82rem;">zł</span>
          </div>
        </div>
        <div class="invoice-issuer-grid" style="display:grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 0.75rem;">
          ${invoiceData.issuers.map(issuer => {
            const currentAllocation = clientInvoice.allocations.find(allocation => allocation.issuerId === issuer.id);
            const manualAmount = clientConfig.manualAmounts?.[issuer.id];
            const checked = selectedIssuerIds.includes(issuer.id) || !!currentAllocation;
            const companyWithEmployees = clientConfig.separateCompanyWithEmployees?.[issuer.id] === true;
            const percentageValue = Number.isFinite(parseFloat(clientInvoice.percentageAllocations?.[issuer.id]))
              ? parseFloat(clientInvoice.percentageAllocations[issuer.id])
              : 0;
            const settlementInfo = issuer.type === 'SEPARATE_COMPANY'
              ? (clientInvoice.settlementAmountsByIssuer?.[issuer.id] || null)
              : null;
            return `
              <div class="invoice-issuer-option" style="display:flex; flex-direction:column; gap:0.55rem; padding:0.85rem; border:1px solid var(--border-color); border-radius:12px; background:rgba(255,255,255,0.03);">
                <span style="display:flex; align-items:center; justify-content:space-between; gap:0.75rem;">
                  <label style="display:flex; align-items:center; gap:0.6rem; cursor:pointer; margin:0;">
                    <input type="checkbox" class="invoice-issuer-toggle" value="${issuer.id}" ${checked ? 'checked' : ''} style="width:auto;">
                    <span style="font-weight:600; color: var(--text-primary);">${getPersonDisplayName(issuer)}</span>
                  </label>
                  <span class="badge ${issuer.type === 'SEPARATE_COMPANY' ? 'badge-working-partner' : 'badge-partner'}">${issuer.type === 'SEPARATE_COMPANY' ? 'Osobna Firma' : 'Wspólnik'}</span>
                </span>
                <span style="font-size:0.8rem; color: var(--text-secondary);">Aktualnie wyliczone: ${Calculations.formatInvoiceCurrency(currentAllocation?.amount || 0)}</span>
                ${isPercentageMode && checked && !(issuer.type === 'SEPARATE_COMPANY' && companyWithEmployees) ? `<div class="invoice-percentage-controls" style="display:flex; flex-direction:column; gap:0.35rem;"><div style="display:flex; justify-content:space-between; font-size:0.8rem; color: var(--text-secondary);"><span>Udział procentowy</span><strong class="invoice-percentage-value">${percentageValue}%</strong></div><input type="range" min="0" max="100" step="1" class="invoice-percentage-slider" data-issuer-id="${issuer.id}" value="${percentageValue}" style="width:100%; accent-color: var(--accent-primary);"><div style="font-size:0.78rem; color: var(--text-secondary);">Kwota z procentu: <strong class="invoice-percentage-amount" style="color: var(--warning);">${Calculations.formatInvoiceCurrency(clientInvoice.netRevenue * (percentageValue / 100))}</strong></div></div>` : ''}
                ${issuer.type === 'SEPARATE_COMPANY' ? `<label style="display:flex; align-items:center; gap:0.5rem; color: var(--text-secondary); font-size:0.8rem; margin:0;"><input type="checkbox" class="invoice-company-with-employees-toggle" data-issuer-id="${issuer.id}" ${companyWithEmployees ? 'checked' : ''} style="width:auto;">Tylko Faktura za firmę z rozliczenia</label>` : ''}
                ${issuer.type === 'SEPARATE_COMPANY' && companyWithEmployees && settlementInfo ? `<div style="font-size:0.78rem; color: var(--text-secondary);">${settlementInfo.invoiceAmountLabel}: <strong style="color: var(--warning);">${Calculations.formatInvoiceCurrency(settlementInfo.invoiceAmount || 0)}</strong></div>` : ''}
                <input type="number" step="0.01" class="invoice-manual-amount" data-issuer-id="${issuer.id}" value="${Number.isFinite(parseFloat(manualAmount)) ? parseFloat(manualAmount) : ''}" placeholder="Ręczna kwota" ${clientInvoice.mode === 'MANUAL' && !(issuer.type === 'SEPARATE_COMPANY' && companyWithEmployees) ? '' : 'disabled'} style="display:${clientInvoice.mode === 'MANUAL' && !(issuer.type === 'SEPARATE_COMPANY' && companyWithEmployees) ? 'block' : 'none'};">
              </div>
            `;
          }).join('')}
        </div>
      </div>
    `;
  }).join('');

  issuerSummary.innerHTML = invoiceData.issuerSummaries.length > 0
    ? invoiceData.issuerSummaries.map(summary => `
      <div class="settlement-detail-row" style="padding: 0.45rem 0; border-bottom: 1px solid var(--border-color);">
        <div>
          <span>${summary.issuerName}</span>
          ${summary.extraInvoicesAmount > 0 ? `<div style="font-size:0.78rem; color: var(--text-secondary); margin-top:0.15rem;">W tym poza rozliczeniem: ${Calculations.formatInvoiceCurrency(summary.extraInvoicesAmount)}</div>` : ''}
        </div>
        <div style="text-align:right;">
          <strong>${Calculations.formatInvoiceCurrency(summary.totalAmount)}</strong>
          ${summary.issuerType !== 'SEPARATE_COMPANY' ? `<div style="font-size:0.78rem; color: var(--text-secondary); margin-top:0.15rem;">Podatek: ${Calculations.formatInvoiceCurrency(summary.taxAmount || 0)}</div>` : ''}
          ${summary.issuerType !== 'SEPARATE_COMPANY' ? `<div style="font-size:0.78rem; color: var(--text-secondary); margin-top:0.15rem;">Suma wszystkich faktur: ${Calculations.formatInvoiceCurrency(summary.yearToDateTotal || 0)}</div>` : ''}
        </div>
      </div>
    `).join('')
    : '<p style="color: var(--text-muted);">Brak wystawców z przypisanymi fakturami.</p>';

  clientSummary.innerHTML = invoiceData.clientInvoices.map(clientInvoice => `
    <div class="settlement-detail-item" style="margin-bottom: 0.75rem;">
      <div class="settlement-detail-item-header">
        <div>
          <h4>${clientInvoice.clientName}</h4>
          <div class="settlement-detail-meta">${clientInvoice.modeLabel}${clientInvoice.deductClientAdvances && (clientInvoice.clientAdvances || 0) > 0 ? ` • zaliczki odliczone: ${Calculations.formatInvoiceCurrency(clientInvoice.clientAdvances || 0)}` : ''}</div>
        </div>
        <strong>${Calculations.formatInvoiceCurrency(clientInvoice.allocatedTotal)}</strong>
      </div>
      <ul class="settlement-inline-list">
        ${clientInvoice.allocations.map(allocation => `<li>${allocation.issuerName}: ${Calculations.formatInvoiceCurrency(allocation.amount)}</li>`).join('')}
      </ul>
    </div>
  `).join('');

  document.querySelectorAll('.invoice-client-card').forEach(card => {
    card.querySelector('.invoice-client-mode')?.addEventListener('change', () => {
      const isSettlementRevenueMode = card.querySelector('.invoice-client-mode')?.value === 'SETTLEMENT_REVENUE';
      const isEqualSplitMode = card.querySelector('.invoice-client-mode')?.value === 'EQUAL_SPLIT';
      card.querySelectorAll('.invoice-manual-amount').forEach(input => {
        const isManual = card.querySelector('.invoice-client-mode')?.value === 'MANUAL';
        input.disabled = !isManual;
        input.style.display = isManual ? 'block' : 'none';
      });
      const includeCostsRow = card.querySelector('.invoice-include-costs-row');
      if (includeCostsRow) includeCostsRow.style.display = isSettlementRevenueMode ? 'flex' : 'none';
      const includeCostsToggle = card.querySelector('.invoice-include-costs-toggle');
      if (includeCostsToggle && !isSettlementRevenueMode) {
        includeCostsToggle.checked = false;
      }
      const randomizeEqualSplitRow = card.querySelector('.invoice-randomize-equal-split-row');
      if (randomizeEqualSplitRow) randomizeEqualSplitRow.style.display = isEqualSplitMode ? 'flex' : 'none';
      const randomizeEqualSplitToggle = card.querySelector('.invoice-randomize-equal-split-toggle');
      if (randomizeEqualSplitToggle && !isEqualSplitMode) {
        randomizeEqualSplitToggle.checked = false;
      }
      saveInvoiceClientConfigFromCard(card);
    });

    card.querySelectorAll('.invoice-issuer-toggle').forEach(input => {
      input.addEventListener('change', () => saveInvoiceClientConfigFromCard(card));
    });

    card.querySelector('.invoice-deduct-advances-toggle')?.addEventListener('change', () => saveInvoiceClientConfigFromCard(card));
    card.querySelector('.invoice-include-costs-toggle')?.addEventListener('change', () => saveInvoiceClientConfigFromCard(card));
    card.querySelector('.invoice-randomize-equal-split-toggle')?.addEventListener('change', () => saveInvoiceClientConfigFromCard(card));
    card.querySelector('.invoice-randomize-equal-split-toggle')?.addEventListener('input', () => saveInvoiceClientConfigFromCard(card));
    card.querySelector('.invoice-equal-split-variance-amount')?.addEventListener('change', () => saveInvoiceClientConfigFromCard(card));
    card.querySelector('.invoice-equal-split-variance-amount')?.addEventListener('input', () => saveInvoiceClientConfigFromCard(card));

    card.querySelectorAll('.invoice-company-with-employees-toggle').forEach(input => {
      input.addEventListener('change', () => saveInvoiceClientConfigFromCard(card));
    });

    card.querySelectorAll('.invoice-manual-amount').forEach(input => {
      input.addEventListener('change', () => saveInvoiceClientConfigFromCard(card));
    });

    card.querySelectorAll('.invoice-percentage-slider').forEach(input => {
      const rebalance = (persist = false) => {
        const clientId = card.getAttribute('data-client-id');
        const issuerIds = Array.from(card.querySelectorAll('.invoice-issuer-toggle:checked')).map(toggle => toggle.value);
        const existingConfig = getInvoiceClientExistingConfig(clientId);
        const currentPercentageAllocations = { ...(existingConfig.percentageAllocations || {}) };
        card.querySelectorAll('.invoice-percentage-slider').forEach(slider => {
          currentPercentageAllocations[slider.getAttribute('data-issuer-id')] = parseFloat(slider.value) || 0;
        });

        const result = rebalanceInvoicePercentageAllocations(
          issuerIds,
          currentPercentageAllocations,
          input.getAttribute('data-issuer-id'),
          input.value,
          existingConfig.lastEditedPercentageIssuerId,
          existingConfig.percentageTouchedIssuerIds || []
        );

        card._pendingInvoicePercentageResult = result;
        updateInvoicePercentageControls(card, result.percentageAllocations);

        if (persist) {
          saveInvoiceClientConfigFromCard(card, result);
        }
      };

      ['pointerdown', 'mousedown', 'touchstart', 'click'].forEach(eventName => {
        input.addEventListener(eventName, (event) => event.stopPropagation());
      });
      input.addEventListener('input', () => rebalance(false));
      input.addEventListener('change', () => rebalance(true));
    });

    card.querySelector('.invoice-client-notes')?.addEventListener('change', () => saveInvoiceClientConfigFromCard(card));
  });

  document.querySelectorAll('.btn-edit-invoice-extra').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = e.currentTarget.getAttribute('data-id');
      const invoice = (invoiceSettings.extraInvoices || []).find(item => item.id === id);
      if (!invoice) return;

      setInvoiceExtraPanelCollapsed(false);
      document.getElementById('invoice-extra-id').value = invoice.id;
      document.getElementById('invoice-extra-client-select').value = invoice.clientId || '';
      document.getElementById('invoice-extra-client-name').value = invoice.clientId ? '' : (invoice.clientName || '');
      document.getElementById('invoice-extra-issuer').value = invoice.issuerId || '';
      document.getElementById('invoice-extra-amount').value = invoice.amount || '';
      document.getElementById('btn-save-invoice-extra').textContent = 'Zapisz';
      document.getElementById('btn-cancel-invoice-extra').style.display = 'inline-flex';
      document.getElementById('invoice-extra-form').scrollIntoView({ behavior: 'smooth', block: 'center' });
    });
  });

  document.querySelectorAll('.btn-delete-invoice-extra').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = e.currentTarget.getAttribute('data-id');
      if (confirm('Usunąć tę dodatkową fakturę?')) {
        Store.deleteInvoiceExtraInvoice(id);
        resetInvoiceExtraForm();
      }
    });
  });

  lucide.createIcons();
}

function renderInvoiceEqualization() {
  const subtitle = document.getElementById('invoice-equalization-subtitle');
  const list = document.getElementById('invoice-equalization-list');
  const unresolved = document.getElementById('invoice-equalization-unresolved');
  if (!list || !unresolved) return;

  const state = Store.getState();
  const selectedMonth = getSelectedMonthKey();
  const settlement = Calculations.generateSettlement(state);
  const data = settlement.invoiceReconciliation || Calculations.calculateInvoiceReconciliation(state, selectedMonth);
  const settlementIssuerById = Object.fromEntries(
    [...(settlement.partners || []), ...(settlement.separateCompanies || []), ...(settlement.workingPartners || [])]
      .filter(entry => entry?.person?.id)
      .map(entry => [entry.person.id, entry])
  );

  if (subtitle) {
    subtitle.textContent = `Wyrównanie po fakturach dla ${formatMonthLabel(selectedMonth)}`;
  }

  if (!data?.issuers?.length) {
    list.innerHTML = '<p style="color: var(--text-muted);">Brak wystawców z wpływami z faktur w wybranym miesiącu.</p>';
    unresolved.innerHTML = '<p style="color: var(--text-muted);">Brak nierozliczonych kwot.</p>';
    return;
  }

  const buildPaymentsHtml = (items, emptyMessage) => {
    if (!items || items.length === 0) {
      return `<p style="color: var(--text-muted);">${emptyMessage}</p>`;
    }

    return `
      <ul class="settlement-inline-list">
        ${items.map(item => {
          const rawAmount = parseFloat(item.amount) || 0;
          const isReceivableFromEmployee = item.type === 'Pensja pracownika' && rawAmount < 0;
          const displayAmount = Calculations.formatInvoiceCurrency(Math.abs(rawAmount));
          const label = isReceivableFromEmployee
            ? `Do odebrania od pracownika: ${item.recipientName}`
            : `${item.recipientName}:`;
          const amountHtml = isReceivableFromEmployee
            ? `<strong style="color: var(--warning);">${displayAmount}</strong>`
            : `<strong>${displayAmount}</strong>`;
          return `<li>${label} ${amountHtml} <span style="color: var(--text-secondary);">(${item.type})</span></li>`;
        }).join('')}
      </ul>
    `;
  };

  const buildBalanceFlowHtml = (entry, settlementEntry) => {
    const revenueInflowEvents = [
      { label: 'Wpływ z faktur', amount: entry.receivedAmount || 0, direction: 'in' },
      ...((entry.incomingTransferPayments || []).map(item => ({
        label: `Wpływ od ${item.payerName}`,
        amount: item.amount || 0,
        direction: 'in',
        type: item.type,
        category: item.category
      })).filter(item => item.category === 'revenue-equalization'))
    ].filter(item => Math.abs(item.amount || 0) >= 0.005);

    const revenueOutflowEvents = (entry.transferPayments || []).map(item => ({
      label: `${item.type}: ${item.recipientName}`,
      amount: item.amount || 0,
      direction: 'out',
      type: item.type,
      category: item.category
    })).filter(item => item.category === 'revenue-equalization' && Math.abs(item.amount || 0) >= 0.005);

    const grossOutflowEvents = [
      ...((entry.salaryPayments || []).map(item => ({
        label: (parseFloat(item.amount) || 0) < 0
          ? `Do odebrania od pracownika: ${item.recipientName}`
          : `${item.type}: ${item.recipientName}`,
        amount: Math.abs(parseFloat(item.amount) || 0),
        direction: (parseFloat(item.amount) || 0) < 0 ? 'in' : 'out',
        amountColor: (parseFloat(item.amount) || 0) < 0 ? 'var(--warning)' : '',
        showSign: (parseFloat(item.amount) || 0) >= 0
      }))),
      ...((entry.officePayments || []).map(item => ({
        label: item.type || 'Płatność do urzędu',
        amount: item.amount || 0,
        direction: 'out',
        category: item.category
      })).filter(item => item.category === 'employee-office'))
    ].filter(item => Math.abs(item.amount || 0) >= 0.005);

    const postTaxEvents = [
      ...((entry.officePayments || []).map(item => ({
        label: item.type || 'Płatność do urzędu',
        amount: item.amount || 0,
        direction: 'out',
        category: item.category
      })).filter(item => item.category === 'invoice-tax' || item.category === 'own-zus')),
      ...((entry.incomingTransferPayments || []).map(item => ({
        label: `Zwrot od ${item.payerName}`,
        amount: item.amount || 0,
        direction: 'in',
        category: item.category
      })).filter(item => item.category === 'tax-reimbursement')),
      ...((entry.transferPayments || []).map(item => ({
        label: `${item.type}: ${item.recipientName}`,
        amount: item.amount || 0,
        direction: 'out',
        category: item.category
      })).filter(item => item.category === 'tax-reimbursement'))
    ].filter(item => Math.abs(item.amount || 0) >= 0.005);

    const events = [...revenueInflowEvents, ...revenueOutflowEvents, ...grossOutflowEvents, ...postTaxEvents];

    if (events.length === 0) {
      return '<p style="color: var(--text-muted);">Brak przepływów dla tego wystawcy.</p>';
    }

    const grossAmount = Number.isFinite(parseFloat(entry.grossTargetAmount))
      ? parseFloat(entry.grossTargetAmount)
      : (Number.isFinite(parseFloat(settlementEntry?.toPayout)) ? parseFloat(settlementEntry.toPayout) : (parseFloat(entry.targetAmount) || 0));
    const grossWithEmployeeSalariesAmount = Number.isFinite(parseFloat(entry.revenueTargetAmount))
      ? parseFloat(entry.revenueTargetAmount)
      : (Number.isFinite(parseFloat(settlementEntry?.grossWithEmployeeSalaries)) ? parseFloat(settlementEntry.grossWithEmployeeSalaries) : grossAmount);
    const postTaxAmount = Number.isFinite(parseFloat(entry.postTaxTargetAmount))
      ? parseFloat(entry.postTaxTargetAmount)
      : grossAmount;
    const showGrossWithEmployeeSalaries = Math.abs(grossWithEmployeeSalariesAmount - grossAmount) >= 0.005;
    const showPostTaxComparison = postTaxEvents.length > 0 || Math.abs(postTaxAmount - grossAmount) >= 0.005;
    let runningTotal = 0;
    const comparisonMatchColor = '#f59e0b';
    const comparisonMismatchColor = '#7f1d1d';

    const getComparisonColor = (actualAmount, comparisonAmount) => (
      Math.abs((actualAmount || 0) - (comparisonAmount || 0)) < 0.005
        ? comparisonMatchColor
        : comparisonMismatchColor
    );

    const renderEventRow = (event, nextTotal) => {
      const formattedAmount = Calculations.formatInvoiceCurrency(Math.abs(event.amount || 0));
      const signedAmount = event.showSign === false
        ? formattedAmount
        : `${event.direction === 'in' ? '+' : '-'}${formattedAmount}`;
      return `
        <div style="display:grid; grid-template-columns: minmax(180px, 1fr) minmax(140px, auto) minmax(180px, auto); gap: 0.75rem; align-items:center; padding: 0.65rem 0; border-bottom: 1px solid var(--border-color); background: transparent;">
          <div>
            <div style="font-weight: 600; color: var(--text-primary);">${event.label}</div>
          </div>
          <div style="text-align:right; font-weight: 700; color: ${event.amountColor || (event.direction === 'in' ? 'var(--success)' : 'var(--danger)')};">${event.amountPrefix || ''}${signedAmount}</div>
          <div style="text-align:right; color: var(--text-primary); font-size: 0.95rem; font-weight: 700;">
            ${Calculations.formatInvoiceCurrency(nextTotal)}
          </div>
        </div>
      `;
    };

    const renderComparisonRow = (label, amount, comparisonColor, isMatch) => `
      <div style="display:grid; grid-template-columns: minmax(180px, 1fr) minmax(140px, auto) minmax(180px, auto); gap: 0.75rem; align-items:center; padding: 0.65rem 0; border-bottom: 1px solid var(--border-color); background: rgba(255,255,255,0.02);">
        <div>
          <div style="font-weight: 600; color: var(--text-primary);">${label}</div>
          <div style="font-size: 0.8rem; color: var(--text-secondary);">z rozliczenia</div>
        </div>
        <div style="text-align:right; color: var(--text-secondary); font-size: 0.82rem;">porównanie</div>
        <div style="text-align:right; color: ${comparisonColor}; font-size: 0.95rem; font-weight: 700;">
          <span style="color: ${isMatch ? 'var(--success)' : 'var(--danger)'}; display: inline-block; font-size: 1.5rem; font-weight: 800; line-height: 1; margin-right: 1rem; vertical-align: middle;">${isMatch ? '✓' : '❌'}</span>${Calculations.formatInvoiceCurrency(amount)}
        </div>
      </div>
    `;

    const renderSectionWithComparison = (sectionEvents, comparisonLabel = '', comparisonAmount = null) => {
      let html = '';

      sectionEvents.forEach((event, index) => {
        const nextTotal = runningTotal + (event.direction === 'in' ? event.amount : -event.amount);
        html += renderEventRow(event, nextTotal);
        runningTotal = nextTotal;
      });

      if (comparisonAmount !== null) {
        html += renderComparisonRow(
          comparisonLabel,
          comparisonAmount,
          getComparisonColor(runningTotal, comparisonAmount),
          Math.abs((runningTotal || 0) - (comparisonAmount || 0)) < 0.005
        );
      }

      return html;
    };

    return `
      <div style="display: flex; flex-direction: column; gap: 0;">
        ${renderSectionWithComparison(
          [...revenueInflowEvents, ...revenueOutflowEvents],
          'Przychód (Brutto) z Pensjami',
          showGrossWithEmployeeSalaries ? grossWithEmployeeSalariesAmount : null
        )}
        ${renderSectionWithComparison(grossOutflowEvents, 'Przychód (Brutto)', grossAmount)}
        ${showPostTaxComparison ? renderSectionWithComparison(postTaxEvents, 'Przychód (Netto)', postTaxAmount) : ''}
      </div>
    `;
  };

  const buildIssuerSummaryCardHtml = (entry) => {
    const counterpartyNetMap = new Map();

    (entry.transferPayments || []).forEach(payment => {
      if (!payment.recipientId || payment.recipientId === 'office') return;
      const amount = parseFloat(payment.amount) || 0;
      if (Math.abs(amount) < 0.005) return;

      const current = counterpartyNetMap.get(payment.recipientId) || {
        counterpartyId: payment.recipientId,
        counterpartyName: payment.recipientName,
        netAmount: 0,
        breakdownAmounts: []
      };
      current.netAmount -= amount;
      current.breakdownAmounts.push(-amount);
      counterpartyNetMap.set(payment.recipientId, current);
    });

    (entry.incomingTransferPayments || []).forEach(payment => {
      if (!payment.payerId || payment.payerId === 'office') return;
      const amount = parseFloat(payment.amount) || 0;
      if (Math.abs(amount) < 0.005) return;

      const current = counterpartyNetMap.get(payment.payerId) || {
        counterpartyId: payment.payerId,
        counterpartyName: payment.payerName,
        netAmount: 0,
        breakdownAmounts: []
      };
      current.netAmount += amount;
      current.breakdownAmounts.push(amount);
      counterpartyNetMap.set(payment.payerId, current);
    });

    const counterpartyBalances = [...counterpartyNetMap.values()]
      .filter(item => Math.abs(item.netAmount) >= 0.005)
      .sort((a, b) => Math.abs(b.netAmount) - Math.abs(a.netAmount) || a.counterpartyName.localeCompare(b.counterpartyName, 'pl-PL'));

    const salaryPayments = (entry.salaryPayments || [])
      .map(payment => ({
        ...payment,
        amount: parseFloat(payment.amount) || 0
      }))
      .filter(payment => payment.amount > 0.005)
      .sort((a, b) => b.amount - a.amount || (a.recipientName || '').localeCompare(b.recipientName || '', 'pl-PL'));

    const salaryTotal = salaryPayments.reduce((sum, payment) => sum + payment.amount, 0);
    const summaryTotal = counterpartyBalances.reduce((sum, item) => sum + (item.netAmount || 0), 0) - salaryTotal;

    return `
      <div class="glass-panel" style="padding: 1rem; border-radius: 14px;">
        <div class="settlement-detail-item-header">
          <div>
            <h4>${entry.issuerName}</h4>
            <div class="settlement-detail-meta">${entry.issuerType === 'WORKING_PARTNER' ? 'Wspólnik pracujący' : 'Wspólnik wystawiający faktury'}</div>
          </div>
          <strong style="font-size: 1.05rem; text-align: right; color: ${summaryTotal > 0 ? 'var(--success)' : (summaryTotal < 0 ? 'var(--danger)' : '#111827')};">${Calculations.formatInvoiceCurrency(summaryTotal)}</strong>
        </div>
        <div class="settlement-detail-stack" style="margin-top: 0.75rem;">
          <div>
            <div class="settlement-detail-meta" style="margin-bottom: 0.45rem;">Rozliczenia netto z innymi wspólnikami i firmami</div>
            <ul class="settlement-inline-list">${counterpartyBalances.length
              ? counterpartyBalances.map(item => `<li>${item.netAmount >= 0 ? `Od <strong style="font-weight: 700; color: inherit;">${item.counterpartyName}</strong> dostaje <strong style="color: var(--success);">${Calculations.formatInvoiceCurrency(Math.abs(item.netAmount))}</strong>` : `Dla <strong style="font-weight: 700; color: inherit;">${item.counterpartyName}</strong> wypłaca <strong style="color: var(--danger);">${Calculations.formatInvoiceCurrency(Math.abs(item.netAmount))}</strong>`}${item.breakdownAmounts && item.breakdownAmounts.length > 1 ? `<div style="font-size: 0.82rem; color: var(--text-secondary); margin-top: 0.15rem; font-weight: 400;">(${item.breakdownAmounts.map(value => `${value >= 0 ? '+' : '-'}${Calculations.formatInvoiceCurrency(Math.abs(value))}`).join(' ')})</div>` : ''}</li>`).join('')
              : '<li>Brak rozliczeń z innymi wspólnikami.</li>'}</ul>
          </div>
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Wypłacane pensje pracownikom</span><strong style="font-weight: 700;">${Calculations.formatInvoiceCurrency(salaryTotal)}</strong></div>
          <ul class="settlement-inline-list">${salaryPayments.length
            ? salaryPayments.map(payment => `<li>${payment.recipientName}: <strong>${Calculations.formatInvoiceCurrency(payment.amount)}</strong></li>`).join('')
            : '<li>Brak wypłat pensji pracownikom.</li>'}</ul>
        </div>
      </div>
    `;
  };

  const summaryIssuerEntries = (data.issuers || []).filter(entry => entry.issuerType === 'PARTNER' || entry.issuerType === 'WORKING_PARTNER');
  const summaryPanelsHtml = summaryIssuerEntries.length > 0
    ? `
      <div class="settlement-detail-section" style="margin-bottom: 1rem;">
        <h3>Skrót zwrotów dla wspólników</h3>
        <div class="settlement-details-grid" style="grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); margin-top: 0.75rem;">
          ${summaryIssuerEntries.map(buildIssuerSummaryCardHtml).join('')}
        </div>
      </div>
    `
    : '';

  list.innerHTML = `${summaryPanelsHtml}${data.issuers.map(entry => `
    <div class="glass-panel" style="padding: 1rem; border-radius: 14px;">
      <div class="settlement-detail-item-header">
        <div>
          <h4 style="font-size: 2rem; line-height: 1.1;">${entry.issuerName}</h4>
          <div class="settlement-detail-meta">${entry.issuerType === 'SEPARATE_COMPANY' ? 'Osobna firma' : 'Wspólnik wystawiający faktury'}</div>
        </div>
        <div style="display:flex; flex-direction:column; align-items:flex-end; gap:0.2rem;">
          <strong style="color: #111827; font-size: 1.05rem;">${Calculations.formatInvoiceCurrency(entry.receivedAmount)}</strong>
          <strong style="color: var(--success); font-size: 1.05rem;">${Calculations.formatInvoiceCurrency(entry.retainedAmount || 0)}</strong>
        </div>
      </div>
      <div style="margin-top: 1rem; margin-bottom: 1rem; background: transparent;">
        ${buildBalanceFlowHtml(entry, settlementIssuerById[entry.issuerId])}
      </div>
      <div class="settlement-details-grid" style="margin-top: 0; margin-bottom: 0;">
        <div class="settlement-detail-item">
          <div class="settlement-detail-item-header">
            <div>
              <h4>Pensje pracowników</h4>
            </div>
            <strong>${Calculations.formatInvoiceCurrency((entry.salaryPayments || []).reduce((sum, item) => sum + (item.amount || 0), 0))}</strong>
          </div>
          ${buildPaymentsHtml(entry.salaryPayments || [], 'Brak pensji do wypłaty z tego konta.')}
        </div>
        <div class="settlement-detail-item">
          <div class="settlement-detail-item-header">
            <div>
              <h4>Opłaca w urzędzie</h4>
            </div>
            <strong>${Calculations.formatInvoiceCurrency((entry.officePayments || []).reduce((sum, item) => sum + (item.amount || 0), 0))}</strong>
          </div>
          ${buildPaymentsHtml(entry.officePayments || [], 'Brak płatności do urzędu z tego konta.')}
        </div>
        <div class="settlement-detail-item">
          <div class="settlement-detail-item-header">
            <div>
              <h4>Dostaje od wspólników</h4>
            </div>
            <strong>${Calculations.formatInvoiceCurrency((entry.incomingTransferPayments || []).reduce((sum, item) => sum + (item.amount || 0), 0))}</strong>
          </div>
          ${buildPaymentsHtml((entry.incomingTransferPayments || []).map(item => ({ ...item, recipientName: item.payerName })), 'Brak wpływów od innych wspólników.')}
        </div>
        <div class="settlement-detail-item">
          <div class="settlement-detail-item-header">
            <div>
              <h4>Przekazuje dalej z faktur</h4>
            </div>
            <strong>${Calculations.formatInvoiceCurrency((entry.transferPayments || []).reduce((sum, item) => sum + (item.amount || 0), 0))}</strong>
          </div>
          ${buildPaymentsHtml(entry.transferPayments || [], 'Brak przelewów wyrównujących z tego konta.')}
        </div>
      </div>
    </div>
  `).join('')}`;

  unresolved.innerHTML = data.unresolvedRecipients && data.unresolvedRecipients.length > 0
    ? `<ul class="settlement-inline-list">${data.unresolvedRecipients.map(item => `<li>${item.personName}: ${Calculations.formatInvoiceCurrency(item.remainingNeed || 0)}</li>`).join('')}</ul>`
    : '<p style="color: var(--success);">Brak nierozliczonych kwot.</p>';
}

function showCustomTimePicker(input) {
  // Remove existing
  const old = document.querySelector('.time-picker-popup');
  if (old) old.remove();

  const rect = input.getBoundingClientRect();
  const popup = document.createElement('div');
  popup.className = 'time-picker-popup animate-fade-in';
  popup.style.top = (rect.bottom + window.scrollY + 5) + 'px';
  popup.style.left = (rect.left + window.scrollX) + 'px';

  const [currentH, currentM] = (input.value || "07:00").split(':');

  // Hours Column
  const colH = document.createElement('div');
  colH.className = 'time-picker-col';
  colH.innerHTML = '<div class="time-picker-header">Godz.</div>';
  for (let h = 0; h < 24; h++) {
    const hh = h.toString().padStart(2, '0');
    const item = document.createElement('div');
    item.className = 'time-picker-item' + (hh === currentH ? ' active' : '');
    item.textContent = hh;
    item.onclick = () => {
      const m = input.value.split(':')[1] || '00';
      updateTime(hh, m);
    };
    colH.appendChild(item);
  }

  // Minutes Column
  const colM = document.createElement('div');
  colM.className = 'time-picker-col';
  colM.innerHTML = '<div class="time-picker-header">Min.</div>';
  ['00', '15', '30', '45'].forEach(mm => {
    const item = document.createElement('div');
    item.className = 'time-picker-item' + (mm === currentM ? ' active' : '');
    item.textContent = mm;
    item.onclick = () => {
      const h = input.value.split(':')[0] || '07';
      updateTime(h, mm);
    };
    colM.appendChild(item);
  });

  popup.appendChild(colH);
  popup.appendChild(colM);
  document.body.appendChild(popup);

  // Scroll active items into view
  [colH, colM].forEach(col => {
    const active = col.querySelector('.active');
    if (active) {
      active.scrollIntoView({ block: 'center', behavior: 'instant' });
    }
  });

  function updateTime(h, m) {
    const newVal = `${h}:${m}`;
    input.value = newVal;
    
    // Save to store
    const d = input.getAttribute('data-day');
    const type = input.getAttribute('data-type');
    const s = Store.getMonthlySheet(window.currentSheetId);
    if (!s.days) s.days = {};
    if (!s.days[d]) s.days[d] = { hours: {} };
    
    if (type === 'start') s.days[d].globalStart = newVal;
    if (type === 'end') s.days[d].globalEnd = newVal;
    
    // Recalculate
    if (s.days[d].isWholeTeamChecked && s.days[d].globalStart && s.days[d].globalEnd) {
      const calcH = Calculations.calculateHours(s.days[d].globalStart, s.days[d].globalEnd);
      activePersonsReloader(s, d, calcH);
    } else {
      Store.updateMonthlySheet(s.id, { days: s.days });
    }
    
    // Close and small refresh if recalculating
    popup.remove();
  }

  // Helper for reloading
  function activePersonsReloader(s, d, calcH) {
    const state = Store.getState();
    const activePersons = getVisiblePersonsForSheet(state, s);

    if (calcH >= 0) {
      activePersons.forEach(p => {
        if (!s.days[d].manual || !s.days[d].manual[p.id]) {
           syncMonthlySheetDayPersonHours(s, d, p.id);
        }
      });
    }
    Store.updateMonthlySheet(s.id, { days: s.days });
    renderSheetDetail(window.currentSheetId);
  }

  // Close on outside click
  const closeHandler = (e) => {
    if (!popup.contains(e.target) && e.target !== input) {
      popup.remove();
      document.removeEventListener('click', closeHandler);
    }
  };
  setTimeout(() => document.addEventListener('click', closeHandler), 10);
}

function showWorksTimePicker(input) {
  // Remove existing
  const old = document.querySelector('.time-picker-popup');
  if (old) old.remove();

  const rect = input.getBoundingClientRect();
  const popup = document.createElement('div');
  popup.className = 'time-picker-popup animate-fade-in';
  popup.style.top = (rect.bottom + window.scrollY + 5) + 'px';
  popup.style.left = (rect.left + window.scrollX) + 'px';

  const type = input.getAttribute('data-type');
  const defaultVal = (type === 'start' ? '07:00' : '17:00');
  const [currentH, currentM] = (input.value || defaultVal).split(':');

  // Hours Column
  const colH = document.createElement('div');
  colH.className = 'time-picker-col';
  colH.innerHTML = '<div class="time-picker-header">Godz.</div>';
  for (let h = 0; h < 24; h++) {
    const hh = h.toString().padStart(2, '0');
    const item = document.createElement('div');
    item.className = 'time-picker-item' + (hh === currentH ? ' active' : '');
    item.textContent = hh;
    item.onclick = () => {
      const m = input.value.split(':')[1] || '00';
      updateTime(hh, m);
    };
    colH.appendChild(item);
  }

  // Minutes Column
  const colM = document.createElement('div');
  colM.className = 'time-picker-col';
  colM.innerHTML = '<div class="time-picker-header">Min.</div>';
  ['00', '15', '30', '45'].forEach(mm => {
    const item = document.createElement('div');
    item.className = 'time-picker-item' + (mm === currentM ? ' active' : '');
    item.textContent = mm;
    item.onclick = () => {
      const h = input.value.split(':')[0] || (type === 'start' ? '07' : '17');
      updateTime(h, mm);
    };
    colM.appendChild(item);
  });

  popup.appendChild(colH);
  popup.appendChild(colM);
  document.body.appendChild(popup);

  // Scroll active items into view
  [colH, colM].forEach(col => {
    const active = col.querySelector('.active');
    if (active) {
      active.scrollIntoView({ block: 'center', behavior: 'instant' });
    }
  });

  function updateTime(h, m) {
    const newVal = `${h}:${m}`;
    input.value = newVal;
    
    // Save to store
    const eid = input.getAttribute('data-eid');
    const type = input.getAttribute('data-type');
    const sheetId = window.currentWorksSheetId;
    const s = Store.getWorksSheet(sheetId);
    
    if (s && s.entries) {
      const entryIdx = s.entries.findIndex(x => x.id === eid);
      if (entryIdx !== -1) {
        const e = s.entries[entryIdx];
        if (type === 'start') {
          e.startTime = newVal;
          if (!e.endTime) e.endTime = '17:00';
        }
        if (type === 'end') {
          e.endTime = newVal;
          if (!e.startTime) e.startTime = '07:00';
        }
        
        // Recalculate hours for all non-manual participants
        if (e.startTime && e.endTime) {
          const calcH = Calculations.calculateHours(e.startTime, e.endTime);
          const state = Store.getState();
          if (s.activePersons) {
            if (!e.hours) e.hours = {};
            if (!e.manual) e.manual = {};
            s.activePersons.forEach(pid => {
              if (!e.manual[pid]) {
                e.hours[pid] = calcH;
              }
            });
          }
        }
        
        Store.updateWorksSheet(sheetId, { entries: s.entries });
        renderWorksSheetDetail(sheetId);
      }
    }
    
    popup.remove();
  }

  // Close on outside click
  const closeHandler = (e) => {
    if (!popup.contains(e.target) && e.target !== input) {
      popup.remove();
      document.removeEventListener('click', closeHandler);
    }
  };
  setTimeout(() => document.addEventListener('click', closeHandler), 10);
}

// ==========================================
// WORKS VIEW (Wykonane Prace)
// ==========================================
function getActiveWorksPersons(month = getSelectedMonthKey()) {
  return Store.getState().persons.filter(person => Store.isPersonActiveInMonth(person.id, month));
}

function getDefaultWorksSheetActivePersonIds(month = getSelectedMonthKey()) {
  return getActiveWorksPersons(month).map(person => person.id);
}

function getDefaultWorksSheetPartnerProfitOverrideIds(month = getSelectedMonthKey()) {
  return getActiveWorksPersons(month)
    .filter(person => person.type === 'PARTNER')
    .map(person => person.id);
}

function getWorksPersonBadgeMeta(person) {
  if (person?.type === 'PARTNER') {
    return { badgeClass: 'badge-partner', badgeLabel: 'WSPÓLNIK' };
  }
  if (person?.type === 'SEPARATE_COMPANY') {
    return { badgeClass: 'badge-partner', badgeLabel: 'OSOBNA FIRMA' };
  }
  if (person?.type === 'WORKING_PARTNER') {
    return { badgeClass: 'badge-working-partner', badgeLabel: 'WSPÓLN. PRAC.' };
  }
  return { badgeClass: 'badge-employee', badgeLabel: 'PRACOWNIK' };
}

function getWorksSheetPersonConfig(sheet = null, month = getSelectedMonthKey()) {
  return {
    activePersons: Array.isArray(sheet?.activePersons)
      ? [...new Set(sheet.activePersons)]
      : getDefaultWorksSheetActivePersonIds(month),
    partnerProfitOverrides: Array.isArray(sheet?.partnerProfitOverrides)
      ? [...new Set(sheet.partnerProfitOverrides)]
      : getDefaultWorksSheetPartnerProfitOverrideIds(month)
  };
}

function renderWorksSheetActivePersonsList(container, sheet = null) {
  if (!container) return;

  const month = getSelectedMonthKey();
  const allPersons = getActiveWorksPersons(month);
  const { activePersons, partnerProfitOverrides } = getWorksSheetPersonConfig(sheet, month);
  const activePersonsSet = new Set(activePersons);
  const partnerProfitOverridesSet = new Set(partnerProfitOverrides);

  if (allPersons.length === 0) {
    container.innerHTML = '<p style="color: var(--text-muted);">Brak aktywnych osób w tym miesiącu.</p>';
    return;
  }

  container.innerHTML = allPersons.map(person => {
    const { badgeClass, badgeLabel } = getWorksPersonBadgeMeta(person);
    const isParticipant = activePersonsSet.has(person.id);
    const canUseProfitOverride = person.type === 'PARTNER';
    const isProfitOverride = canUseProfitOverride && partnerProfitOverridesSet.has(person.id);

    return `
      <div class="active-person-card">
        <div class="active-person-card-header">
          <label class="active-person-toggle">
            <input type="checkbox" class="cb-participant" data-pid="${person.id}" ${isParticipant ? 'checked' : ''} style="width: 19px; height: 19px; margin: 0;">
            <span class="active-person-name">${getPersonDisplayName(person)}</span>
          </label>
          <span class="badge ${badgeClass}" style="padding: 0.3rem 0.7rem; font-size: 0.72rem; letter-spacing: 0.02em;">${badgeLabel}</span>
        </div>
        ${canUseProfitOverride ? `
          <div class="active-person-override">
            <label class="active-person-override-label">
              <input type="checkbox" class="cb-profit-override" data-pid="${person.id}" ${isProfitOverride ? 'checked' : ''} style="width: 16px; height: 16px; margin: 0;">
              Otrzymuje udział w zysku pomimo braku uczestnictwa
            </label>
          </div>
        ` : ''}
      </div>
    `;
  }).join('');
}

function collectWorksSheetPersonConfig(container) {
  const scope = container || document;
  const activePersons = [];
  const partnerProfitOverrides = [];

  scope.querySelectorAll('.cb-participant:checked').forEach(checkbox => {
    activePersons.push(checkbox.getAttribute('data-pid'));
  });

  scope.querySelectorAll('.cb-profit-override:checked').forEach(checkbox => {
    partnerProfitOverrides.push(checkbox.getAttribute('data-pid'));
  });

  return { activePersons, partnerProfitOverrides };
}

function populateWorksSheetClientSelect(selectedClient = '') {
  const clientSelect = document.getElementById('works-sheet-client-select');
  if (!clientSelect) return;

  const activeClients = getActiveClients();
  clientSelect.innerHTML = '<option value="">-- Wybierz klienta --</option>';

  activeClients.forEach(client => {
    const selected = client.name === selectedClient ? 'selected' : '';
    clientSelect.innerHTML += `<option value="${client.name}" ${selected}>${client.name}</option>`;
  });

  if (selectedClient && !activeClients.some(client => client.name === selectedClient)) {
    const currentClient = Store.getState().clients.find(client => client.name === selectedClient);
    if (currentClient) {
      clientSelect.innerHTML += `<option value="${currentClient.name}" selected>${currentClient.name} - nieaktywny</option>`;
    }
  }
}

function populateWorksSheetPricesGrid(sheet = null) {
  const grid = document.getElementById('works-sheet-prices-grid');
  if (!grid) return;

  grid.innerHTML = '';
  Store.getState().worksCatalog.forEach(work => {
    const priceVal = (sheet?.customWorkPrices && sheet.customWorkPrices[work.id]) ? sheet.customWorkPrices[work.id] : '';
    grid.innerHTML += `
      <div class="form-group">
        <label>${work.name} (${work.unit})</label>
        <input type="number" step="0.01" data-work-id="${work.id}" class="sheet-work-price-input" placeholder="Domyślnie ${parseFloat(work.defaultPrice).toFixed(2)}" value="${priceVal}">
      </div>
    `;
  });
}

function openWorksSheetMetaForm(sheet = null) {
  const listContainer = document.getElementById('works-sheet-list');
  const metaForm = document.getElementById('works-meta-form');
  const detailContainer = document.getElementById('works-detail-container');
  const catalogContainer = document.getElementById('works-catalog-container');
  const activePersonsForm = document.getElementById('works-active-persons-form');
  const activePersonsList = document.getElementById('works-meta-active-persons-list');
  const isEditing = !!sheet;

  document.getElementById('works-sheet-id').value = isEditing ? sheet.id : '';
  document.getElementById('works-sheet-site').value = isEditing ? (sheet.site || '') : '';
  document.getElementById('works-meta-title').textContent = isEditing ? 'Edytuj Arkusz Wykonanych Prac' : 'Nowy Arkusz Wykonanych Prac';

  populateWorksSheetClientSelect(isEditing ? (sheet.client || '') : '');
  renderWorksSheetActivePersonsList(activePersonsList, sheet);
  populateWorksSheetPricesGrid(sheet);

  listContainer.style.display = 'none';
  detailContainer.style.display = 'none';
  catalogContainer.style.display = 'none';
  if (activePersonsForm) activePersonsForm.style.display = 'none';
  metaForm.style.display = 'block';
}

function initWorksTracker() {
  const listContainer = document.getElementById('works-sheet-list');
  const metaForm = document.getElementById('works-meta-form');
  const detailContainer = document.getElementById('works-detail-container');
  const catalogContainer = document.getElementById('works-catalog-container');
  
  document.getElementById('btn-works-catalog').addEventListener('click', () => {
    if (catalogContainer.style.display === 'none') {
      catalogContainer.style.display = 'block';
      listContainer.style.display = 'none';
      metaForm.style.display = 'none';
      detailContainer.style.display = 'none';
      renderWorksCatalog();
    } else {
      catalogContainer.style.display = 'none';
      if (window.currentWorksSheetId) {
        detailContainer.style.display = 'block';
      } else {
        listContainer.style.display = 'block';
      }
    }
  });
  document.getElementById('btn-close-works-catalog').addEventListener('click', () => {
    catalogContainer.style.display = 'none';
    if (window.currentWorksSheetId) {
      detailContainer.style.display = 'block';
    } else {
      listContainer.style.display = 'block';
    }
  });

  document.getElementById('btn-create-works-sheet').addEventListener('click', () => {
    openWorksSheetMetaForm(null);
  });

  document.getElementById('btn-cancel-works-sheet').addEventListener('click', () => {
    metaForm.style.display = 'none';
    if (window.currentWorksSheetId) {
      detailContainer.style.display = 'block';
    } else {
      listContainer.style.display = 'block';
    }
  });

  document.getElementById('btn-back-to-works-sheets').addEventListener('click', () => {
    window.currentWorksSheetId = null;
    detailContainer.style.display = 'none';
    catalogContainer.style.display = 'none';
    metaForm.style.display = 'none';
    listContainer.style.display = 'block';
    renderWorksSheets();
  });

  document.getElementById('btn-edit-works-meta').addEventListener('click', () => {
    if (window.currentWorksSheetId) {
      const s = Store.getWorksSheet(window.currentWorksSheetId);
      if (s) {
        openWorksSheetMetaForm(s);
      }
    }
  });

  // Active Persons logic
  const activePersonsForm = document.getElementById('works-active-persons-form');
  const activePersonsList = document.getElementById('active-persons-list');

  document.getElementById('btn-active-persons').addEventListener('click', () => {
    if (window.currentWorksSheetId) {
      const s = Store.getWorksSheet(window.currentWorksSheetId);
      if (s) {
        renderWorksSheetActivePersonsList(activePersonsList, s);

        document.getElementById('works-sheet-list').style.display = 'none';
        document.getElementById('works-detail-container').style.display = 'none';
        document.getElementById('works-catalog-container').style.display = 'none';
        document.getElementById('works-meta-form').style.display = 'none';
        activePersonsForm.style.display = 'block';
        activePersonsForm.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    }
  });

  document.getElementById('btn-close-active-persons').addEventListener('click', () => {
    activePersonsForm.style.display = 'none';
    if (window.currentWorksSheetId) {
      document.getElementById('works-detail-container').style.display = 'block';
    } else {
      document.getElementById('works-sheet-list').style.display = 'block';
    }
  });

  activePersonsForm.addEventListener('submit', (e) => {
    e.preventDefault();
    if (window.currentWorksSheetId) {
      const { activePersons, partnerProfitOverrides } = collectWorksSheetPersonConfig(activePersonsForm);

      Store.updateWorksSheet(window.currentWorksSheetId, { activePersons, partnerProfitOverrides });
      
      activePersonsForm.style.display = 'none';
      renderWorksSheetDetail(window.currentWorksSheetId);
    }
  });

  metaForm.addEventListener('submit', (e) => {
    e.preventDefault();
    const id = document.getElementById('works-sheet-id').value;
    const month = getSelectedMonthKey();
    const client = document.getElementById('works-sheet-client-select').value;
    const site = document.getElementById('works-sheet-site').value;
    
    if (!client) {
      alert('Wybierz klienta!');
      return;
    }

    const worksMetaActivePersonsList = document.getElementById('works-meta-active-persons-list');
    const { activePersons, partnerProfitOverrides } = collectWorksSheetPersonConfig(worksMetaActivePersonsList);
    const customWorkPrices = {};
    document.querySelectorAll('.sheet-work-price-input').forEach(input => {
      if (input.value !== '') {
        customWorkPrices[input.getAttribute('data-work-id')] = parseFloat(input.value);
      }
    });

    if (id) {
      Store.updateWorksSheet(id, { client, site, customWorkPrices, activePersons, partnerProfitOverrides });
      window.currentWorksSheetId = id;
    } else {
      Store.addWorksSheet({ month, client, site, customWorkPrices, activePersons, partnerProfitOverrides });
      const newState = Store.getState();
      window.currentWorksSheetId = newState.worksSheets[newState.worksSheets.length - 1].id;
    }

    metaForm.style.display = 'none';
    renderWorksSheets();
    renderWorksSheetDetail(window.currentWorksSheetId);
  });

  // Entry form logic
  const entryForm = document.getElementById('works-entry-form');
  const workSelect = document.getElementById('works-entry-work');
  const unitInput = document.getElementById('works-entry-unit');
  const priceInput = document.getElementById('works-entry-price');
  
  const btnCalc = document.getElementById('btn-calc-qty');
  const calcPopup = document.getElementById('works-calc-popup');
  const btnCloseCalc = document.getElementById('btn-close-calc');
  const btnApplyCalc = document.getElementById('btn-apply-calc');
  const calcLen = document.getElementById('calc-length');
  const calcWid = document.getElementById('calc-width');
  const calcHgt = document.getElementById('calc-height');
  const calcHgtGroup = document.getElementById('calc-height-group');
  const calcTitle = document.getElementById('calc-title');
  const calcWidthLabel = document.getElementById('calc-width-label');
  const calcResult = document.getElementById('calc-result');

  workSelect.addEventListener('change', (e) => {
    const work = Store.getState().worksCatalog.find(w => w.id === e.target.value);
    if (work) {
      unitInput.value = work.unit;

      let price = work.defaultPrice;
      if (window.currentWorksSheetId) {
        const sheet = Store.getWorksSheet(window.currentWorksSheetId);
        if (sheet) {
          const clientObj = Store.getState().clients.find(c => c.name === sheet.client);
          if (clientObj && clientObj.customWorkPrices && typeof clientObj.customWorkPrices[work.id] !== 'undefined') {
            price = clientObj.customWorkPrices[work.id];
          }
          if (sheet.customWorkPrices && typeof sheet.customWorkPrices[work.id] !== 'undefined') {
            price = sheet.customWorkPrices[work.id];
          }
        }
      }
      priceInput.value = price;

      if (work.unit === 'm3' || work.unit === 'm²' || work.unit === 'm³' || work.unit === 'm2') {
        btnCalc.style.display = 'block';
      } else {
        btnCalc.style.display = 'none';
        calcPopup.style.display = 'none';
      }
    } else {
      unitInput.value = '';
      priceInput.value = '';
      btnCalc.style.display = 'none';
      calcPopup.style.display = 'none';
    }
  });

  btnCalc.addEventListener('click', () => {
    calcPopup.style.display = 'block';
    const workId = workSelect.value;
    const work = Store.getState().worksCatalog.find(w => w.id === workId);
    if (work && (work.unit === 'm²' || work.unit === 'm2')) {
      calcTitle.textContent = 'Kalkulator m²';
      if (calcWidthLabel) calcWidthLabel.textContent = 'Szerokość / Wysokość (m)';
      calcHgtGroup.style.display = 'none';
    } else {
      calcTitle.textContent = 'Kalkulator m³';
      if (calcWidthLabel) calcWidthLabel.textContent = 'Szerokość (m)';
      calcHgtGroup.style.display = 'block';
    }
    calcResult.textContent = '0.00';
    calcLen.value = '';
    calcWid.value = '';
    calcHgt.value = '';
    setTimeout(() => calcLen.focus(), 50);
  });

  btnCloseCalc.addEventListener('click', () => {
    calcPopup.style.display = 'none';
  });

  const updateCalc = () => {
    const l = parseFloat(calcLen.value) || 0;
    const w = parseFloat(calcWid.value) || 0;
    let res = 0;
    if (calcTitle.textContent.includes('m³')) {
      const h = parseFloat(calcHgt.value) || 0;
      res = l * w * h;
    } else {
      res = l * w;
    }
    calcResult.textContent = res.toFixed(2);
  };

  [calcLen, calcWid, calcHgt].forEach(inp => inp.addEventListener('input', updateCalc));

  btnApplyCalc.addEventListener('click', () => {
    document.getElementById('works-entry-qty').value = calcResult.textContent;
    calcPopup.style.display = 'none';
  });

  entryForm.addEventListener('submit', (e) => {
    e.preventDefault();
    if (!window.currentWorksSheetId) return;
    
    const editId = document.getElementById('works-entry-edit-id').value;
    const workId = workSelect.value;
    const work = Store.getState().worksCatalog.find(w => w.id === workId) || { name: 'Inne', unit: '-' };
    const date = document.getElementById('works-entry-date').value;
    const qty = parseFloat(document.getElementById('works-entry-qty').value);
    const price = parseFloat(document.getElementById('works-entry-price').value);
    
    const sheet = Store.getWorksSheet(window.currentWorksSheetId);
    if (sheet) {
      if (!sheet.entries) sheet.entries = [];
      
      if (editId) {
        const entryIndex = sheet.entries.findIndex(e => e.id === editId);
        if (entryIndex !== -1) {
          sheet.entries[entryIndex] = {
            ...sheet.entries[entryIndex],
            workId,
            name: work.name,
            unit: work.unit,
            date: date,
            quantity: qty,
            price: price
          };
        }
      } else {
        sheet.entries.push({
          id: crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36),
          workId,
          name: work.name,
          unit: work.unit,
          date: date,
          quantity: qty,
          price: price
        });
      }
      Store.updateWorksSheet(sheet.id, { entries: sheet.entries });
      
      // Reset
      document.getElementById('works-entry-edit-id').value = '';
      document.getElementById('works-entry-date').value = getDefaultDateForSelectedMonth();
      document.getElementById('works-entry-qty').value = '';
      const formTitle = document.getElementById('works-entry-form-title');
      if (formTitle) formTitle.textContent = 'Dodaj nową pozycję';
      document.getElementById('btn-submit-works-entry').textContent = 'Dodaj';
      document.getElementById('btn-cancel-works-entry').style.display = 'none';
      renderWorksSheetDetail(sheet.id);
    }
  });

  const btnCancelWorksEntry = document.getElementById('btn-cancel-works-entry');
  if (btnCancelWorksEntry) {
    btnCancelWorksEntry.addEventListener('click', () => {
      document.getElementById('works-entry-edit-id').value = '';
      document.getElementById('works-entry-date').value = getDefaultDateForSelectedMonth();
      document.getElementById('works-entry-qty').value = '';

      const formTitle = document.getElementById('works-entry-form-title');
      if (formTitle) formTitle.textContent = 'Dodaj nową pozycję';
      document.getElementById('btn-submit-works-entry').textContent = 'Dodaj';
      document.getElementById('btn-cancel-works-entry').style.display = 'none';

      const workId = workSelect.value;
      const work = Store.getState().worksCatalog.find(w => w.id === workId);
      if (work) {
        document.getElementById('works-entry-unit').value = work.unit;
        document.getElementById('works-entry-price').value = work.defaultPrice;
      } else {
        document.getElementById('works-entry-unit').value = '';
        document.getElementById('works-entry-price').value = '';
      }
    });
  }

  // Catalog Form Logic
  const catalogForm = document.getElementById('work-catalog-form');

  function resetCatalogForm() {
    document.getElementById('work-catalog-id').value = '';
    document.getElementById('work-catalog-name').value = '';
    document.getElementById('work-catalog-price').value = '';
    const formTitle = document.getElementById('work-catalog-form-title');
    if (formTitle) formTitle.textContent = 'Dodaj nową pozycję do katalogu';
    const btnSubmit = document.getElementById('btn-submit-works-catalog');
    if (btnSubmit) btnSubmit.textContent = 'Dodaj';
    const btnCancel = document.getElementById('btn-cancel-works-catalog');
    if (btnCancel) btnCancel.style.display = 'none';
    const btnRestore = document.getElementById('btn-restore-works-catalog');
    if (btnRestore) btnRestore.style.display = 'none';
    if (catalogForm) catalogForm.reset();
  }

  function checkCatalogFormAltered() {
    const id = document.getElementById('work-catalog-id').value;
    if (!id) {
        document.getElementById('btn-restore-works-catalog').style.display = 'none';
        return;
    }
    const w = Store.getState().worksCatalog.find(x => x.id === id);
    if (!w || !w.coreId) {
        document.getElementById('btn-restore-works-catalog').style.display = 'none';
        return;
    }

    const def = Store.getDefaultCatalogItem(w.coreId);
    if (!def) return;

    const currentName = document.getElementById('work-catalog-name').value.trim();
    const currentUnit = document.getElementById('work-catalog-unit').value;
    const currentPrice = parseFloat(document.getElementById('work-catalog-price').value);

    const isAltered = currentName !== def.name || currentUnit !== def.unit || currentPrice !== parseFloat(def.defaultPrice);
    document.getElementById('btn-restore-works-catalog').style.display = isAltered ? 'block' : 'none';
  }

  document.getElementById('work-catalog-name').addEventListener('input', checkCatalogFormAltered);
  document.getElementById('work-catalog-unit').addEventListener('change', checkCatalogFormAltered);
  document.getElementById('work-catalog-price').addEventListener('input', checkCatalogFormAltered);

  document.getElementById('btn-cancel-works-catalog').addEventListener('click', () => {
    resetCatalogForm();
  });

  document.getElementById('btn-restore-works-catalog').addEventListener('click', () => {
    const id = document.getElementById('work-catalog-id').value;
    if (!id) return;
    const w = Store.getState().worksCatalog.find(x => x.id === id);
    if (w && w.coreId) {
      const def = Store.getDefaultCatalogItem(w.coreId);
      if (def) {
        document.getElementById('work-catalog-name').value = def.name;
        document.getElementById('work-catalog-unit').value = def.unit;
        document.getElementById('work-catalog-price').value = def.defaultPrice;
        document.getElementById('btn-restore-works-catalog').style.display = 'none';
      }
    }
  });

  catalogForm.addEventListener('submit', (e) => {
    e.preventDefault();
    const id = document.getElementById('work-catalog-id').value;
    const name = document.getElementById('work-catalog-name').value;
    const unit = document.getElementById('work-catalog-unit').value;
    const price = parseFloat(document.getElementById('work-catalog-price').value);

    if (id) {
      Store.updateWorkInCatalog(id, { name, unit, defaultPrice: price });
    } else {
      Store.addWorkToCatalog({ name, unit, defaultPrice: price });
    }

    resetCatalogForm();
    renderWorksCatalog();
    if (window.currentWorksSheetId) {
      renderWorksSheetDetail(window.currentWorksSheetId);
    }
  });

  // Navigation fix for works view
  const views = document.querySelectorAll('.view');
  document.querySelectorAll('.nav-item').forEach(link => {
    link.addEventListener('click', (e) => {
       const targetId = link.getAttribute('data-target');
       if (targetId === 'works-view') {
         if (window.currentWorksSheetId) {
           renderWorksSheetDetail(window.currentWorksSheetId);
         } else {
           renderWorksSheets();
         }
       }
    });
  });
}

function renderWorksSheets() {
  const state = Store.getState();
  const selectedMonth = getSelectedMonthKey();
  const tbody = document.getElementById('works-sheets-body');
  
  if (window.currentWorksSheetId && document.getElementById('works-detail-container').style.display === 'block') {
    return;
  }
  
  tbody.innerHTML = '';
  document.getElementById('works-sheet-list').style.display = 'block';
  document.getElementById('works-detail-container').style.display = 'none';
  document.getElementById('works-meta-form').style.display = 'none';
  document.getElementById('works-catalog-container').style.display = 'none';
  document.getElementById('works-active-persons-form').style.display = 'none';

  const sheetsForMonth = (state.worksSheets || []).filter(s => s.month === selectedMonth);
  if (sheetsForMonth.length === 0) {
    tbody.innerHTML = `<tr><td colspan="5" style="text-align:center;color:var(--text-muted)">Brak arkuszy prac w tym miesiącu. Kliknij "Nowy Arkusz".</td></tr>`;
    return;
  }

  sheetsForMonth.forEach(s => {
    const entries = s.entries || [];
    const totalItems = entries.length;
    const totalValue = entries.reduce((sum, e) => sum + (parseFloat(e.quantity) * parseFloat(e.price)), 0);

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${s.client || '-'}</td>
      <td>${s.site || '-'}</td>
      <td>${totalItems} pozycji</td>
      <td style="color: var(--success); font-weight: bold;">${totalValue.toFixed(2)} zł</td>
      <td>
        <button class="btn btn-primary btn-sm btn-open-works-sheet" data-id="${s.id}" style="margin-right:0.5rem;">Otwórz</button>
        <button class="btn btn-secondary btn-icon btn-edit-works-sheet" data-id="${s.id}" style="margin-right:0.5rem;">
          <i data-lucide="edit-2" style="width:16px;height:16px"></i>
        </button>
        <button class="btn btn-danger btn-icon btn-delete-works-sheet" data-id="${s.id}">
          <i data-lucide="trash-2" style="width:16px;height:16px"></i>
        </button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  document.querySelectorAll('.btn-open-works-sheet').forEach(btn => {
    btn.addEventListener('click', (e) => {
      window.currentWorksSheetId = e.currentTarget.getAttribute('data-id');

      // Reset form on open
      document.getElementById('works-entry-edit-id').value = '';
      document.getElementById('works-entry-date').value = getDefaultDateForSelectedMonth();
      document.getElementById('works-entry-qty').value = '';
      const formTitle = document.getElementById('works-entry-form-title');
      if (formTitle) formTitle.textContent = 'Dodaj nową pozycję';
      const btnSubmit = document.getElementById('btn-submit-works-entry');
      if (btnSubmit) btnSubmit.textContent = 'Dodaj';
      const btnCancel = document.getElementById('btn-cancel-works-entry');
      if (btnCancel) btnCancel.style.display = 'none';

      renderWorksSheetDetail(window.currentWorksSheetId);
    });
  });

  document.querySelectorAll('.btn-edit-works-sheet').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = e.currentTarget.getAttribute('data-id');
      const s = state.worksSheets.find(x => x.id === id);
      if (s) {
        openWorksSheetMetaForm(s);
      }
    });
  });

  document.querySelectorAll('.btn-delete-works-sheet').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = e.currentTarget.getAttribute('data-id');
      if (confirm('Usunąć ten arkusz? Przychody z wykonanych prac zostaną odjęte od podziału.')) {
        Store.deleteWorksSheet(id);
        renderWorksSheets();
      }
    });
  });
  
  lucide.createIcons();
}

function renderWorksSheetDetail(sheetId) {
  const sheet = Store.getWorksSheet(sheetId);
  if (!sheet) {
    renderWorksSheets();
    return;
  }
  document.getElementById('works-sheet-list').style.display = 'none';
  document.getElementById('works-meta-form').style.display = 'none';
  document.getElementById('works-catalog-container').style.display = 'none';
  document.getElementById('works-active-persons-form').style.display = 'none';
  document.getElementById('works-detail-container').style.display = 'block';
  document.getElementById('detail-works-title').textContent = `${sheet.month} - ${sheet.client}`;
  document.getElementById('detail-works-site').textContent = sheet.site || 'Brak wpisanej budowy';

  const workSelect = document.getElementById('works-entry-work');
  workSelect.innerHTML = '<option value="">-- Wybierz z katalogu --</option>';
  Store.getState().worksCatalog.forEach(w => {
    workSelect.innerHTML += `<option value="${w.id}">${w.name} (${w.unit})</option>`;
  });
  
  const tbody = document.getElementById('works-entries-body');
  tbody.innerHTML = '';
  let sum = 0;
  
  let entries = sheet.entries || [];
  entries.sort((a, b) => new Date(a.date || '1970-01-01') - new Date(b.date || '1970-01-01'));

  if (entries.length === 0) {
    tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;color:var(--text-muted)">Brak wprowadzonych pozycji.</td></tr>';
  } else {
    const state = Store.getState();
    entries.forEach((e, idx) => {
      const metrics = Calculations.getWorksEntryMetrics(e, sheet, state);
      const isRoboczogodziny = metrics.isRoboczogodziny;
      const totalLoggedHours = metrics.totalLoggedHours;
      const val = metrics.revenue;
      sum += val;

      const profitToSplit = metrics.profit;
      let wartoscHtml = `<div style="font-weight: bold; font-size: 1.05rem;">${val.toFixed(2)} zł</div>`;
      if (sheet.activePersons && sheet.activePersons.length > 0) {
         if (metrics.profitBaseRevenue !== metrics.revenue) {
           wartoscHtml += `<div style="font-size: 0.75rem; color: var(--primary); margin-top: 4px;">Roboczogodziny pracowników: ${metrics.profitBaseRevenue.toFixed(2)} zł</div>`;
         }
         wartoscHtml += `<div style="font-size: 0.75rem; color: var(--danger); margin-top: 4px;">Pensje: -${metrics.employeeCost.toFixed(2)} zł</div>`;
         wartoscHtml += `<div style="font-size: 0.75rem; color: var(--success); margin-top: 2px; font-weight: 500;">Dla wspóln.: ${profitToSplit.toFixed(2)} zł</div>`;
      }

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${e.date || '-'}</td>
        <td style="font-weight: 500">${e.name}</td>
        <td>${isRoboczogodziny ? (totalLoggedHours + ' h') : (e.quantity + ' ' + e.unit)}</td>
        <td>${isRoboczogodziny ? (Calculations.getSheetClientRate(sheet, state).toFixed(2) + ' zł/h') : (parseFloat(e.price).toFixed(2) + ' zł')}</td>
        <td>${wartoscHtml}</td>
        <td style="vertical-align: middle;">
          <div style="display:flex; gap:0.5rem; justify-content: flex-start;">
            <button class="btn btn-secondary btn-icon btn-edit-works-entry" data-idx="${idx}">
              <i data-lucide="edit-2" style="width:16px;height:16px"></i>
            </button>
            <button class="btn btn-danger btn-icon btn-delete-works-entry" data-idx="${idx}">
              <i data-lucide="trash-2" style="width:16px;height:16px"></i>
            </button>
          </div>
        </td>
      `;
      tbody.appendChild(tr);

      if (sheet.activePersons && sheet.activePersons.length > 0) {
        const subTr = document.createElement('tr');
        subTr.className = 'works-hours-row';

        const startTime = e.startTime || '';
        const endTime = e.endTime || '';

        let subHtml = `<td colspan="6" style="padding: 0; background: var(--bg-secondary); border-bottom: 1px solid var(--border-color);">
          <div class="hide-scrollbar" style="display: flex; overflow-x: auto; overflow-y: hidden; -webkit-overflow-scrolling: touch;">
            <!-- CZAS Column -->
            <div style="flex: 0 0 130px; border-right: 1px solid var(--border-color); display: flex; flex-direction: column; min-width: 130px;">
              <div style="padding: 0.2rem 0.4rem; text-align: center; border-bottom: 1px solid var(--border-color); font-size: 0.6rem; font-weight: 700; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 0.05em;">CZAS</div>
              <div style="padding: 0.25rem 0.4rem; flex: 1; display: flex; align-items: center; justify-content: center; gap: 2px;">
                <input type="text" class="works-time-in" data-eid="${e.id}" data-type="start" readonly value="${startTime}" placeholder="--:--" style="width: 44px; padding: 1px; background: transparent; text-align: center; cursor: pointer; border: 1px solid transparent; border-radius: 4px; font-size: 0.85rem; font-weight: 500; color: #60a5fa;">
                <span style="opacity: 0.4; color: var(--text-secondary); font-size: 0.75rem;">-</span>
                <input type="text" class="works-time-in" data-eid="${e.id}" data-type="end" readonly value="${endTime}" placeholder="--:--" style="width: 44px; padding: 1px; background: transparent; text-align: center; cursor: pointer; border: 1px solid transparent; border-radius: 4px; font-size: 0.85rem; font-weight: 500; color: #60a5fa;">
              </div>
            </div>

            <!-- Persons Columns -->
            ${sheet.activePersons.map(pid => {
               const person = state.persons.find(p => p.id === pid);
               if (!person) return '';
               const hrVal = (e.hours && typeof e.hours[pid] !== 'undefined') ? e.hours[pid] : '';
               const isManual = (e.manual && e.manual[pid]);
               return `
                 <div class="works-hours-cell" style="flex: 0 0 110px; border-right: 1px solid var(--border-color); display: flex; flex-direction: column; min-width: 110px; background: ${isManual ? 'rgba(96, 165, 250, 0.03)' : 'transparent'};">
                    <div style="padding: 0.2rem 0.4rem; text-align: center; border-bottom: 1px solid var(--border-color); font-size: 0.6rem; font-weight: 700; color: var(--text-secondary); letter-spacing: 0.05em; overflow: hidden; text-overflow: ellipsis; white-space: normal; line-height: 1.05;" title="${getPersonDisplayName(person)}">${getPersonCompactHeaderHtml(person)}</div>
                   <div style="padding: 0.25rem 0.4rem; flex: 1; display: flex; align-items: center; justify-content: center; position: relative;">
                      <input type="number" step="0.25" min="0" class="works-sub-hour-in" data-eid="${e.id}" data-pid="${pid}" value="${hrVal}" placeholder="0" style="width: 55px; padding: 1px; font-size: 1.05rem; font-weight: 600; text-align: center; border: none; background: transparent; color: ${isManual ? '#60a5fa' : 'var(--text-primary)'};">
                      <div class="works-cell-actions" style="position: absolute; right: 2px; top: 50%; transform: translateY(-50%); display: flex; flex-direction: column; gap: 2px;">
                        <button class="btn-works-zero" data-eid="${e.id}" data-pid="${pid}" title="Ustaw 0" style="background: none; border: none; padding: 1px; cursor: pointer; color: var(--danger); opacity: 0.6;">
                          <i data-lucide="x" style="width: 12px; height: 12px;"></i>
                        </button>
                        <button class="btn-works-reset" data-eid="${e.id}" data-pid="${pid}" title="Przywróć automat" style="background: none; border: none; padding: 1px; cursor: pointer; color: var(--primary); opacity: 0.6;">
                          <i data-lucide="rotate-ccw" style="width: 12px; height: 12px;"></i>
                        </button>
                      </div>
                   </div>
                 </div>
               `;
            }).join('')}
          </div>
        </td>`;
        subTr.innerHTML = subHtml;
        tbody.appendChild(subTr);
      }
    });
  }

  document.querySelectorAll('.works-sub-hour-in').forEach(inp => {
    inp.addEventListener('change', (ev) => {
      const eid = ev.target.getAttribute('data-eid');
      const pid = ev.target.getAttribute('data-pid');
      const paramVal = ev.target.value;
      
      const sheetToUpdate = Store.getWorksSheet(sheetId);
      if (sheetToUpdate) {
          const entryIdx = sheetToUpdate.entries.findIndex(x => x.id === eid);
          if (entryIdx !== -1) {
             const e = sheetToUpdate.entries[entryIdx];
             if (!e.hours) e.hours = {};
             if (!e.manual) e.manual = {};
             
             if (paramVal !== '') {
               e.hours[pid] = parseFloat(paramVal);
               e.manual[pid] = true;
             } else {
               delete e.hours[pid];
               delete e.manual[pid];
             }
             
             Store.updateWorksSheet(sheetId, { entries: sheetToUpdate.entries });
             renderWorksSheetDetail(sheetId);
          }
      }
    });
  });

  document.querySelectorAll('.works-time-in').forEach(inp => {
    inp.addEventListener('click', (e) => {
      showWorksTimePicker(e.target);
    });
  });

  document.querySelectorAll('.btn-works-zero').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const target = e.currentTarget;
      const eid = target.getAttribute('data-eid');
      const pid = target.getAttribute('data-pid');
      const s = Store.getWorksSheet(sheetId);
      if (s) {
        const entryIdx = s.entries.findIndex(x => x.id === eid);
        if (entryIdx !== -1) {
          const entry = s.entries[entryIdx];
          if (!entry.hours) entry.hours = {};
          if (!entry.manual) entry.manual = {};
          entry.hours[pid] = 0;
          entry.manual[pid] = true;
          Store.updateWorksSheet(sheetId, { entries: s.entries });
          renderWorksSheetDetail(sheetId);
        }
      }
    });
  });

  document.querySelectorAll('.btn-works-reset').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const target = e.currentTarget;
      const eid = target.getAttribute('data-eid');
      const pid = target.getAttribute('data-pid');
      const s = Store.getWorksSheet(sheetId);
      if (s) {
        const entryIdx = s.entries.findIndex(x => x.id === eid);
        if (entryIdx !== -1) {
          const entry = s.entries[entryIdx];
          if (entry.manual) delete entry.manual[pid];
          
          if (entry.startTime && entry.endTime) {
            const calcH = Calculations.calculateHours(entry.startTime, entry.endTime);
            if (!entry.hours) entry.hours = {};
            entry.hours[pid] = calcH;
          } else {
            if (entry.hours) delete entry.hours[pid];
          }
          
          Store.updateWorksSheet(sheetId, { entries: s.entries });
          renderWorksSheetDetail(sheetId);
        }
      }
    });
  });

  document.getElementById('works-sheet-total-value').textContent = `${sum.toFixed(2)} zł`;

  document.querySelectorAll('.btn-edit-works-entry').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const idx = e.currentTarget.getAttribute('data-idx');
      const entry = sheet.entries[idx];
      if (entry) {
        document.getElementById('works-entry-edit-id').value = entry.id;
        document.getElementById('works-entry-date').value = entry.date || getDefaultDateForSelectedMonth();
        document.getElementById('works-entry-work').value = entry.workId;
        document.getElementById('works-entry-qty').value = entry.quantity;
        document.getElementById('works-entry-unit').value = entry.unit;
        document.getElementById('works-entry-price').value = entry.price;

        const formTitle = document.getElementById('works-entry-form-title');
        if (formTitle) formTitle.textContent = 'Edytujesz pozycję: ' + entry.name;
        document.getElementById('btn-submit-works-entry').textContent = 'Zapisz';
        document.getElementById('btn-cancel-works-entry').style.display = 'block';

        document.getElementById('works-entry-form').scrollIntoView({ behavior: 'smooth', block: 'center' });

        // Ensure calc button visibility matches
        const btnCalc = document.getElementById('btn-calc-qty');
        if (entry.unit === 'm3' || entry.unit === 'm²' || entry.unit === 'm³' || entry.unit === 'm2') {
          btnCalc.style.display = 'block';
        } else {
          btnCalc.style.display = 'none';
        }
      }
    });
  });

  document.querySelectorAll('.btn-delete-works-entry').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const idx = e.currentTarget.getAttribute('data-idx');
      sheet.entries.splice(idx, 1);
      Store.updateWorksSheet(sheet.id, { entries: sheet.entries });
      renderWorksSheetDetail(sheet.id);
    });
  });

  lucide.createIcons();
}

function renderWorksCatalog() {
  const state = Store.getState();
  const tbody = document.getElementById('works-catalog-body');
  tbody.innerHTML = '';
  
  const allDefaults = Store.getDefaultCatalogItems ? Store.getDefaultCatalogItems() : [];
  const currentCoreIds = new Set(state.worksCatalog.map(w => w.coreId).filter(Boolean));

  if ((!state.worksCatalog || state.worksCatalog.length === 0) && allDefaults.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;color:var(--text-muted)">Brak prac w katalogu.</td></tr>';
    return;
  }
  
  // Render current works
  if (state.worksCatalog) {
    state.worksCatalog.forEach(w => {
      let isAltered = false;
      if (w.coreId) {
        const def = Store.getDefaultCatalogItem(w.coreId);
        if (def && (w.name !== def.name || w.unit !== def.unit || parseFloat(w.defaultPrice) !== parseFloat(def.defaultPrice))) {
          isAltered = true;
        }
      }

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${w.name}</td>
        <td>${w.unit}</td>
        <td>${parseFloat(w.defaultPrice).toFixed(2)} zł</td>
        <td>
          <div style="display:flex; gap:0.5rem;">
            <button class="btn btn-secondary btn-icon btn-edit-catalog" data-id="${w.id}">
              <i data-lucide="edit-2" style="width:16px;height:16px"></i>
            </button>
            ${isAltered ? `
            <button class="btn btn-secondary btn-icon btn-restore-catalog" data-id="${w.id}" title="Przywróć domyślne">
              <i data-lucide="rotate-ccw" style="width:16px;height:16px"></i>
            </button>` : ''}
            <button class="btn btn-danger btn-icon btn-delete-catalog" data-id="${w.id}">
              <i data-lucide="trash-2" style="width:16px;height:16px"></i>
            </button>
          </div>
        </td>
      `;
      tbody.appendChild(tr);
    });
  }

  // Render deleted defaults
  allDefaults.forEach(def => {
    if (!currentCoreIds.has(def.coreId)) {
      const tr = document.createElement('tr');
      tr.style.opacity = '0.5';
      tr.innerHTML = `
        <td style="text-decoration: line-through;">${def.name}</td>
        <td style="text-decoration: line-through;">${def.unit}</td>
        <td style="text-decoration: line-through;">${parseFloat(def.defaultPrice).toFixed(2)} zł</td>
        <td>
          <div style="display:flex; gap:0.5rem;">
            <button class="btn btn-secondary btn-icon btn-restore-deleted-catalog" data-coreid="${def.coreId}" title="Przywróć pozycję z domyślnych">
              <i data-lucide="rotate-ccw" style="width:16px;height:16px"></i>
            </button>
          </div>
        </td>
      `;
      tbody.appendChild(tr);
    }
  });

  if (tbody.innerHTML === '') {
    tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;color:var(--text-muted)">Brak prac w katalogu.</td></tr>';
  }
  
  document.querySelectorAll('.btn-edit-catalog').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = e.currentTarget.getAttribute('data-id');
      const w = Store.getState().worksCatalog.find(x => x.id === id);
      if (w) {
        document.getElementById('work-catalog-id').value = w.id;
        document.getElementById('work-catalog-name').value = w.name;
        document.getElementById('work-catalog-unit').value = w.unit;
        document.getElementById('work-catalog-price').value = w.defaultPrice;

        const formTitle = document.getElementById('work-catalog-form-title');
        if (formTitle) formTitle.textContent = 'Edytujesz pozycję: ' + w.name;
        document.getElementById('btn-submit-works-catalog').textContent = 'Zapisz';
        document.getElementById('btn-cancel-works-catalog').style.display = 'block';

        if (typeof checkCatalogFormAltered === 'function') {
          checkCatalogFormAltered();
        } else {
          let isAltered = false;
          if (w.coreId) {
            const def = Store.getDefaultCatalogItem(w.coreId);
            if (def && (w.name !== def.name || w.unit !== def.unit || parseFloat(w.defaultPrice) !== parseFloat(def.defaultPrice))) {
              isAltered = true;
            }
          }
          document.getElementById('btn-restore-works-catalog').style.display = isAltered ? 'block' : 'none';
        }

        document.getElementById('work-catalog-form').scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
    });
  });

  document.querySelectorAll('.btn-restore-catalog').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = e.currentTarget.getAttribute('data-id');
      if (confirm('Przywrócić tę pozycję do wartości domyślnych z katalogu?')) {
        Store.restoreWorkInCatalogToDefault(id);
        if (typeof resetCatalogForm === 'function') resetCatalogForm();
        renderWorksCatalog(); // Trigger UI rebuild directly
      }
    });
  });

  document.querySelectorAll('.btn-restore-deleted-catalog').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const coreId = e.currentTarget.getAttribute('data-coreid');
      if (confirm('Przywrócić tę usuniętą domyślną pozycję?')) {
        Store.restoreDeletedDefaultToCatalog(coreId);
        if (typeof resetCatalogForm === 'function') resetCatalogForm();
        renderWorksCatalog();
      }
    });
  });

  document.querySelectorAll('.btn-delete-catalog').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = e.currentTarget.getAttribute('data-id');
      if (confirm('Usunąć tę pracę z katalogu?')) {
        Store.deleteWorkFromCatalog(id);
        renderWorksCatalog();
      }
    });
  });
  
  lucide.createIcons();
}

// ==========================================
// SETTINGS & SCALING
// ==========================================

const DEFAULT_APPEARANCE_SETTINGS = { theme: 'system', scaleLarge: 100, scaleVertical: 100 };
const DEFAULT_MOBILE_APPEARANCE_SETTINGS = { theme: 'system', scaleLarge: 60, scaleVertical: 80 };

function initFirebaseAuth() {
  const auth = firebase.auth();
  const db = firebase.database();
  const overlay = document.getElementById('firebase-auth-overlay');
  const appRoot = document.getElementById('app-root');
  
  const emailInput = document.getElementById('firebase-email');
  const passwordInput = document.getElementById('firebase-password');
  const btnLoginEmail = document.getElementById('btn-firebase-email-login');
  const btnLoginGoogle = document.getElementById('btn-firebase-google');
  const errorMsg = document.getElementById('firebase-auth-error');
  
  auth.onAuthStateChanged((user) => {
    if (user) {
      if (overlay) overlay.style.display = 'none';
      if (appRoot) appRoot.style.display = 'flex';
      
      const statusText = document.getElementById('sidebar-auth-status-text');
      const btnSidebarLogout = document.getElementById('btn-sidebar-firebase-logout');
      if (statusText) {
         statusText.textContent = user.email || 'Zalogowano';
         statusText.style.color = 'var(--success)';
      }
      if (btnSidebarLogout) {
         btnSidebarLogout.textContent = 'Wyloguj';
         btnSidebarLogout.style.display = 'block';
         btnSidebarLogout.onclick = () => auth.signOut();
      }
      
      db.ref('shared_data').on('value', (snapshot) => {
        const data = snapshot.val();
        if (data) {
          window.isImportingFromFirebase = true;
          const oldSheet = window.currentSheetId;
          const oldWorks = window.currentWorksSheetId;
          Store.importState(data);
          window.currentSheetId = oldSheet;
          window.currentWorksSheetId = oldWorks;
          renderAll();
          window.isImportingFromFirebase = false;
          
          if (user.email) {
             boostPersonByEmail(user.email);
          }
        }
      });
      
    } else {
      if (overlay) overlay.style.display = 'flex';
      if (appRoot) appRoot.style.display = 'none';
      db.ref('shared_data').off('value');
    }
  });

  if (btnLoginEmail) {
    btnLoginEmail.addEventListener('click', (e) => {
      e.preventDefault();
      const email = emailInput?.value || '';
      const password = passwordInput?.value || '';
      if (!email || !password) {
        if (errorMsg) { 
          errorMsg.textContent = 'Wprowadź e-mail i hasło.';
          errorMsg.style.display = 'block';
        }
        return;
      }
      auth.signInWithEmailAndPassword(email, password).catch(err => {
        if (errorMsg) {
          errorMsg.textContent = 'Błąd logowania: ' + err.message;
          errorMsg.style.display = 'block';
        }
      });
    });
  }

  if (btnLoginGoogle) {
    btnLoginGoogle.addEventListener('click', (e) => {
      e.preventDefault();
      const provider = new firebase.auth.GoogleAuthProvider();
      auth.signInWithPopup(provider).catch(err => {
        if (errorMsg) {
          errorMsg.textContent = 'Błąd Google: ' + err.message;
          errorMsg.style.display = 'block';
        }
      });
    });
  }
}

function setSettingsDataStatus(message = '', type = '') {
  const statusEl = document.getElementById('settings-data-status');
  if (!statusEl) return;
  statusEl.textContent = message;
  statusEl.classList.remove('is-success', 'is-error');
  if (type === 'success') statusEl.classList.add('is-success');
  if (type === 'error') statusEl.classList.add('is-error');
}

function resetAppearanceSettings(preset = 'default') {
  const nextSettings = preset === 'mobile'
    ? DEFAULT_MOBILE_APPEARANCE_SETTINGS
    : DEFAULT_APPEARANCE_SETTINGS;
  if (Store.resetAppearanceSettings) {
    Store.resetAppearanceSettings(nextSettings);
    return;
  }
  Store.updateAppearanceSettings(nextSettings);
}

function buildExportPayload() {
  return {
    app: 'work-tracker-html',
    version: 1,
    exportedAt: new Date().toISOString(),
    data: Store.getExportData()
  };
}

function boostPersonByEmail(email) {
  if (!email) return;
  const state = Store.getState();
  if (!state || !state.persons || state.persons.length === 0) return;
  
  const searchEmail = email.trim().toLowerCase();
  
  const matchIndex = state.persons.findIndex(p => {
    const googleEmail = p.email_google_account ? p.email_google_account.trim().toLowerCase() : '';
    const normalEmail = p.email ? p.email.trim().toLowerCase() : '';
    return googleEmail === searchEmail || normalEmail === searchEmail;
  });
  
  if (matchIndex > 0) {
    const ids = state.persons.map(p => p.id);
    const [matchedId] = ids.splice(matchIndex, 1);
    ids.unshift(matchedId);
    Store.reorderPersons(ids);
  }
}

function buildDataExportFileName() {
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  return `monitor-pracy-backup-${timestamp}.json`;
}

function exportDatabaseToFile() {
  const payload = buildExportPayload();

  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = buildDataExportFileName();
  document.body.appendChild(link);
  link.click();
  link.remove();

  setTimeout(() => URL.revokeObjectURL(url), 1000);
  setSettingsDataStatus('Eksport bazy danych zakończony. Plik JSON został pobrany.', 'success');
}

function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error('Nie udało się odczytać wybranego pliku.'));
    reader.readAsText(file, 'utf-8');
  });
}

async function importDatabaseFromFile(file) {
  if (!file) return;

  const content = await readFileAsText(file);
  let parsed;

  try {
    parsed = JSON.parse(content);
  } catch {
    throw new Error('Wybrany plik nie jest poprawnym plikiem JSON.');
  }

  window.currentSheetId = null;
  window.currentWorksSheetId = null;
  Store.importState(parsed);
  setSettingsDataStatus(`Zaimportowano bazę danych z pliku: ${file.name}`, 'success');
}

function initSettings() {
  const themeRadios = document.querySelectorAll('input[name="theme-select"]');
  const scaleLargeInput = document.getElementById('scale-large');
  const scaleVerticalInput = document.getElementById('scale-vertical');
  const btnReset = document.getElementById('btn-reset-settings');
  const btnResetMobile = document.getElementById('btn-reset-mobile-settings');
  const btnExportData = document.getElementById('btn-export-data');
  const btnImportData = document.getElementById('btn-import-data');
  const importDataFile = document.getElementById('import-data-file');

  if (!themeRadios.length || !scaleLargeInput || !scaleVerticalInput) return;

  themeRadios.forEach(radio => {
    radio.addEventListener('change', (e) => {
      Store.updateAppearanceSettings({ theme: e.target.value });
    });
  });

  scaleLargeInput.addEventListener('input', (e) => {
    const val = parseInt(e.target.value);
    document.getElementById('scale-large-text').textContent = `${val}%`;
    Store.updateAppearanceSettings({ scaleLarge: val });
  });

  scaleVerticalInput.addEventListener('input', (e) => {
    const val = parseInt(e.target.value);
    document.getElementById('scale-vertical-text').textContent = `${val}%`;
    Store.updateAppearanceSettings({ scaleVertical: val });
  });

  if (btnReset) {
    btnReset.addEventListener('click', () => {
      if (confirm('Czy na pewno chcesz przywrócić domyślne ustawienia wyglądu?')) {
        resetAppearanceSettings('default');
      }
    });
  }

  if (btnResetMobile) {
    btnResetMobile.addEventListener('click', () => {
      if (confirm('Czy na pewno chcesz przywrócić mobilne ustawienia wyglądu?')) {
        resetAppearanceSettings('mobile');
      }
    });
  }

  if (btnExportData) {
    btnExportData.addEventListener('click', () => {
      try {
        exportDatabaseToFile();
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Nie udało się wyeksportować bazy danych.';
        setSettingsDataStatus(message, 'error');
        alert(message);
      }
    });
  }

  if (btnImportData && importDataFile) {
    btnImportData.addEventListener('click', () => {
      setSettingsDataStatus('');
      importDataFile.click();
    });

    importDataFile.addEventListener('change', async (e) => {
      const file = e.target.files && e.target.files[0];
      if (!file) return;

      if (!confirm(`Zaimportować bazę danych z pliku "${file.name}"? Obecne dane aplikacji zostaną nadpisane i wysłane do chmury Firebase.`)) {
        importDataFile.value = '';
        return;
      }

      try {
        setSettingsDataStatus('Trwa import bazy danych...', '');
        await importDatabaseFromFile(file);
        
        if (typeof firebase !== 'undefined' && firebase.auth && firebase.auth().currentUser && !window.isImportingFromFirebase) {
           firebase.database().ref('shared_data').set(Store.getState());
        }

      } catch (error) {
        const message = error instanceof Error ? error.message : 'Nie udało się zaimportować bazy danych.';
        setSettingsDataStatus(message, 'error');
        alert(message);
      } finally {
        importDataFile.value = '';
      }
    });
  }

  initInteractiveScaling();
}

function applySettings() {
  const settings = Store.getSettings ? Store.getSettings() : { ...DEFAULT_APPEARANCE_SETTINGS, ...(Store.getState().settings || {}) };
  
  // Apply Theme
  const root = document.documentElement;
  if (settings.theme === 'system') {
    root.removeAttribute('data-theme');
  } else {
    root.setAttribute('data-theme', settings.theme);
  }

  // Update Settings UI if active
  const themeRadio = document.querySelector(`input[name="theme-select"][value="${settings.theme}"]`);
  if (themeRadio) themeRadio.checked = true;

  const scaleLargeInput = document.getElementById('scale-large');
  if (scaleLargeInput) {
    scaleLargeInput.value = settings.scaleLarge;
    document.getElementById('scale-large-text').textContent = `${settings.scaleLarge}%`;
  }

  const scaleVerticalInput = document.getElementById('scale-vertical');
  if (scaleVerticalInput) {
    scaleVerticalInput.value = settings.scaleVertical;
    document.getElementById('scale-vertical-text').textContent = `${settings.scaleVertical}%`;
  }

  // Apply Scale
  const isVertical = window.innerHeight > window.innerWidth;
  const currentScale = isVertical ? settings.scaleVertical : settings.scaleLarge;
  root.style.setProperty('--app-scale', currentScale / 100);
}

let scaleOverlayTimeout = null;
function showScaleOverlay(scalePercent) {
  const overlay = document.getElementById('scale-info-overlay');
  if (!overlay) return;

  overlay.textContent = `${Math.round(scalePercent)}%`;
  overlay.classList.add('visible');

  clearTimeout(scaleOverlayTimeout);
  scaleOverlayTimeout = setTimeout(() => {
    overlay.classList.remove('visible');
  }, 1000);
}

function initInteractiveScaling() {
  // 1. Ctrl + Mouse Wheel
  window.addEventListener('wheel', (e) => {
    if (e.ctrlKey) {
      e.preventDefault();
      const settings = Store.getSettings ? Store.getSettings() : Store.getState().settings;
      const isVertical = window.innerHeight > window.innerWidth;
      const currentScale = isVertical ? settings.scaleVertical : settings.scaleLarge;
      
      let nextScale = currentScale - (e.deltaY > 0 ? 5 : -5);
      nextScale = Math.max(20, Math.min(300, nextScale));

      if (nextScale !== currentScale) {
        const update = isVertical ? { scaleVertical: nextScale } : { scaleLarge: nextScale };
        Store.updateAppearanceSettings(update);
        showScaleOverlay(nextScale);
      }
    }
  }, { passive: false });

  // 2. Pinch to Zoom
  let initialPinchDistance = null;
  let initialScale = null;

  window.addEventListener('touchstart', (e) => {
    if (e.touches.length === 2) {
      const settings = Store.getSettings ? Store.getSettings() : Store.getState().settings;
      const isVertical = window.innerHeight > window.innerWidth;
      initialScale = isVertical ? settings.scaleVertical : settings.scaleLarge;
      initialPinchDistance = Math.hypot(
        e.touches[0].pageX - e.touches[1].pageX,
        e.touches[0].pageY - e.touches[1].pageY
      );
    }
  }, { passive: true });

  window.addEventListener('touchmove', (e) => {
    if (e.touches.length === 2 && initialPinchDistance !== null) {
      e.preventDefault();
      const currentDistance = Math.hypot(
        e.touches[0].pageX - e.touches[1].pageX,
        e.touches[0].pageY - e.touches[1].pageY
      );
      
      const factor = currentDistance / initialPinchDistance;
      let nextScale = initialScale * factor;
      nextScale = Math.max(20, Math.min(300, nextScale));

      const isVertical = window.innerHeight > window.innerWidth;
      const settings = Store.getSettings ? Store.getSettings() : Store.getState().settings;
      const currentAppScale = isVertical ? settings.scaleVertical : settings.scaleLarge;

      if (Math.abs(nextScale - currentAppScale) > 1) {
        const update = isVertical ? { scaleVertical: Math.round(nextScale) } : { scaleLarge: Math.round(nextScale) };
        Store.updateAppearanceSettings(update);
        showScaleOverlay(nextScale);
      }
    }
  }, { passive: false });

  window.addEventListener('touchend', () => {
    initialPinchDistance = null;
  });

  // Handle Resize/Orientation change
  window.addEventListener('resize', () => {
    applySettings();
  });
}
