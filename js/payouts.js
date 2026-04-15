function getPayoutsStateBundle() {
  const exportState = Store.getExportData ? Store.getExportData() : { version: 'v3', common: {}, months: {} };
  return {
    ...exportState,
    selectedMonth: getSelectedMonthKey()
  };
}

function formatPayoutDateLabel(value = '') {
  if (!value) return '—';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleDateString('pl-PL', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  });
}

function getPayoutEmployeeRecordState(item = {}) {
  if ((item.remainingAmount || 0) <= 0.005) return 'settled';
  if ((item.settledAmount || 0) > 0.005) return 'partial';
  return 'pending';
}

function buildPayoutEmployeeCardHtml(item = {}, data = {}) {
  const personId = item?.person?.id || '';
  const settlementStarted = (item.settledAmount || 0) > 0.005;
  const stateKey = getPayoutEmployeeRecordState(item);
  const statusMeta = stateKey === 'settled'
    ? { label: 'Rozliczono', className: 'badge-working-partner' }
    : (stateKey === 'partial'
        ? { label: 'Rozliczono częściowo', className: 'badge-partner' }
        : { label: 'Do rozliczenia', className: 'badge-employee' });
  const firstDayOfMonth = data?.payoutMonth ? `${data.payoutMonth}-01` : '';
  const lastDayOfMonth = (() => {
    if (!data?.payoutMonth) return '';
    const [year, month] = data.payoutMonth.split('-').map(Number);
    return `${data.payoutMonth}-${String(new Date(year, month, 0).getDate()).padStart(2, '0')}`;
  })();
  const totalAdvanceInfo = (item.totalAdvanceAmountToDate || 0) > (item.availableAdvanceAmount || 0) + 0.005
    ? `Do dnia wypłaty jest więcej zaliczek (${formatSettlementCompactCurrency(item.totalAdvanceAmountToDate)}), ale do tego rozliczenia mieszczą się ${formatSettlementCompactCurrency(item.availableAdvanceAmount)}.`
    : `Do rozliczenia można potrącić ${formatSettlementCompactCurrency(item.availableAdvanceAmount)}.`;

  return `
    <article class="glass-panel payout-employee-card" data-person-id="${escapeReportHtml(personId)}" style="padding: 1.25rem; margin-bottom: 1rem;">
      <div class="payout-card-header">
        <div>
          <h3 style="margin-bottom: 0.2rem;">${escapeReportHtml(getPersonDisplayName(item.person))}</h3>
          <p class="payout-card-subtitle">Pensja za ${escapeReportHtml(formatMonthLabel(item.previousMonth))} • Data wypłaty ${escapeReportHtml(formatPayoutDateLabel(item.payoutDate))}</p>
        </div>
        <span class="badge ${statusMeta.className}">${statusMeta.label}</span>
      </div>

      <div class="payout-card-summary-grid">
        <div class="glass-panel payout-summary-box">
          <div class="payout-summary-label">Pensja z poprzedniego miesiąca</div>
          <div class="payout-summary-value">${escapeReportHtml(formatSettlementCompactCurrency(item.baseAmount || 0))}</div>
        </div>
        <div class="glass-panel payout-summary-box">
          <div class="payout-summary-label">Nierozliczone wypłaty</div>
          <div class="payout-summary-value payout-summary-value--warning">${escapeReportHtml(formatSettlementCompactCurrency(item.carryoverAmount || 0))}</div>
        </div>
        <div class="glass-panel payout-summary-box">
          <div class="payout-summary-label">Pozostało do rozliczenia</div>
          <div class="payout-summary-value payout-summary-value--success">${escapeReportHtml(formatSettlementCompactCurrency(item.remainingAmount || 0))}</div>
        </div>
      </div>

      <div class="payout-card-body">
        <div class="payout-card-config">
          <label class="choice-option payout-choice-option">
            <input type="checkbox" class="payout-include-carryover-toggle" ${item.payoutRecord?.includeCarryover !== false ? 'checked' : ''} ${settlementStarted ? 'disabled' : ''}>
            <span class="choice-option-text">Dolicz nierozliczone wypłaty z poprzednich miesięcy</span>
          </label>

          <div class="form-group" style="margin-bottom: 0;">
            <label>Potrącanie zaliczek</label>
            <select class="payout-deduct-advances-mode" ${settlementStarted ? 'disabled' : ''}>
              <option value="default-day" ${item.deductAdvancesMode === 'default-day' ? 'selected' : ''}>Do dnia wypłaty</option>
              <option value="custom-day" ${item.deductAdvancesMode === 'custom-day' ? 'selected' : ''}>Do wybranego dnia</option>
              <option value="none" ${item.deductAdvancesMode === 'none' ? 'selected' : ''}>Bez potrącania zaliczek</option>
            </select>
          </div>

          <div class="form-group payout-custom-date-group" style="margin-bottom: 0; display: ${item.deductAdvancesMode === 'custom-day' ? 'block' : 'none'};">
            <label>Wybrany dzień potrącenia</label>
            <input type="date" class="payout-custom-date-input" value="${escapeReportHtml(item.deductionDate || item.payoutDate || '')}" min="${firstDayOfMonth}" max="${lastDayOfMonth}" ${settlementStarted ? 'disabled' : ''}>
          </div>

          <div class="glass-panel payout-advance-box">
            <div class="payout-summary-label">Zaliczki do rozliczenia</div>
            <div class="payout-summary-value ${item.availableAdvanceAmount > 0 ? 'payout-summary-value--danger' : ''}">${escapeReportHtml(formatSettlementCompactCurrency(item.availableAdvanceAmount || 0))}</div>
            <div class="payout-card-note">${escapeReportHtml(totalAdvanceInfo)}</div>
          </div>
        </div>

        <div class="payout-card-actions-panel">
          <div class="payout-card-note">Rozliczono wcześniej gotówką: <strong>${escapeReportHtml(formatSettlementCompactCurrency(item.settledCashAmount || 0))}</strong></div>
          <div class="payout-card-note">Rozliczono wcześniej zaliczkami: <strong>${escapeReportHtml(formatSettlementCompactCurrency(item.settledAdvanceAmount || 0))}</strong></div>
          ${item.payoutRecord?.lastSettledAt ? `<div class="payout-card-note">Ostatnie rozliczenie: <strong>${escapeReportHtml(formatHistoryTimestamp(item.payoutRecord.lastSettledAt))}</strong></div>` : ''}

          <div class="form-group" style="margin-bottom: 0;">
            <label>Kwota wypłacona teraz (gotówka / przelew)</label>
            <input type="number" step="0.01" min="0" class="payout-cash-amount" value="${(item.payoutNowAmount || 0) > 0 ? (item.payoutNowAmount || 0).toFixed(2) : ''}" ${item.remainingAmount <= 0.005 ? 'disabled' : ''}>
          </div>

          <div class="payout-card-buttons">
            <button type="button" class="btn btn-secondary btn-payout-settle-partial" ${item.remainingAmount <= 0.005 ? 'disabled' : ''}>Rozlicz częściowo</button>
            <button type="button" class="btn btn-primary btn-payout-settle-full" ${item.remainingAmount <= 0.005 ? 'disabled' : ''}>Wypłacono całość</button>
          </div>
        </div>
      </div>
    </article>
  `;
}

function settlePayoutForEmployee(personId = '', mode = 'partial') {
  const currentMonth = getSelectedMonthKey();
  const payoutsData = Calculations.calculatePayoutsData(getPayoutsStateBundle(), currentMonth);
  const item = (payoutsData.employees || []).find(entry => entry?.person?.id === personId);
  const card = document.querySelector(`.payout-employee-card[data-person-id="${personId}"]`);
  if (!item || !card) return;

  const advanceExpenseIds = (item.availableAdvanceExpenses || []).map(expense => expense.id);
  const advanceAmount = Math.max(0, parseFloat(item.availableAdvanceAmount) || 0);
  const maxCashAmount = Math.max(0, (item.remainingAmount || 0) - advanceAmount);
  const cashInput = card.querySelector('.payout-cash-amount');
  const inputCashAmount = Math.max(0, parseFloat(cashInput?.value) || 0);
  const cashAmount = mode === 'full' ? maxCashAmount : inputCashAmount;

  if ((item.remainingAmount || 0) <= 0.005) {
    alert('Ta wypłata jest już rozliczona.');
    return;
  }

  if (cashAmount > (maxCashAmount + 0.005)) {
    alert(`Kwota wypłacana teraz nie może przekroczyć ${formatSettlementCompactCurrency(maxCashAmount)}.`);
    return;
  }

  if (cashAmount <= 0 && advanceAmount <= 0) {
    alert('Brak kwoty do rozliczenia. Wpisz kwotę częściowej wypłaty lub włącz potrącanie zaliczek.');
    return;
  }

  const success = Store.settleEmployeePayout?.(personId, {
    includeCarryover: item.payoutRecord?.includeCarryover !== false,
    deductAdvancesMode: item.deductAdvancesMode,
    customDeductionDate: item.payoutRecord?.customDeductionDate || '',
    sourceMonth: item.previousMonth,
    baseAmountSnapshot: item.baseAmount,
    carryoverAmountSnapshot: item.carryoverAmount,
    plannedAmountSnapshot: item.plannedAmount,
    advanceDeductionAmountSnapshot: advanceAmount,
    advanceExpenseIds,
    cashAmount
  }, currentMonth);

  if (!success) {
    alert('Nie udało się rozliczyć wypłaty.');
  }
}

function initPayouts() {
  const defaultDayInput = document.getElementById('payouts-default-day');
  const list = document.getElementById('payouts-list');
  if (!defaultDayInput || !list || defaultDayInput.dataset.bound === 'true') return;

  defaultDayInput.dataset.bound = 'true';

  defaultDayInput.addEventListener('change', () => {
    const nextValue = Math.max(1, Math.min(31, parseInt(defaultDayInput.value, 10) || 15));
    defaultDayInput.value = `${nextValue}`;
    Store.updatePayoutMonthConfig?.({ defaultDay: nextValue });
  });

  list.addEventListener('change', (event) => {
    const card = event.target.closest('.payout-employee-card');
    const personId = card?.getAttribute('data-person-id') || '';
    if (!personId) return;

    if (event.target.classList.contains('payout-include-carryover-toggle')) {
      Store.updatePayoutEmployeeConfig?.(personId, { includeCarryover: event.target.checked }, getSelectedMonthKey());
      return;
    }

    if (event.target.classList.contains('payout-deduct-advances-mode')) {
      const updates = { deductAdvancesMode: event.target.value };
      if (event.target.value === 'custom-day') {
        updates.customDeductionDate = card.querySelector('.payout-custom-date-input')?.value || '';
      }
      Store.updatePayoutEmployeeConfig?.(personId, updates, getSelectedMonthKey());
      return;
    }

    if (event.target.classList.contains('payout-custom-date-input')) {
      Store.updatePayoutEmployeeConfig?.(personId, { customDeductionDate: event.target.value || '' }, getSelectedMonthKey());
    }
  });

  list.addEventListener('click', (event) => {
    const partialButton = event.target.closest('.btn-payout-settle-partial');
    if (partialButton) {
      const personId = partialButton.closest('.payout-employee-card')?.getAttribute('data-person-id') || '';
      settlePayoutForEmployee(personId, 'partial');
      return;
    }

    const fullButton = event.target.closest('.btn-payout-settle-full');
    if (fullButton) {
      const personId = fullButton.closest('.payout-employee-card')?.getAttribute('data-person-id') || '';
      settlePayoutForEmployee(personId, 'full');
    }
  });
}

function renderPayouts() {
  const subtitle = document.getElementById('payouts-subtitle');
  const defaultDayInput = document.getElementById('payouts-default-day');
  const sourceMonthLabel = document.getElementById('payouts-source-month-label');
  const defaultDateLabel = document.getElementById('payouts-default-date-label');
  const totalBaseEl = document.getElementById('payouts-total-base');
  const totalCarryoverEl = document.getElementById('payouts-total-carryover');
  const totalRemainingEl = document.getElementById('payouts-total-remaining');
  const list = document.getElementById('payouts-list');
  if (!defaultDayInput || !sourceMonthLabel || !defaultDateLabel || !totalBaseEl || !totalCarryoverEl || !totalRemainingEl || !list) return;

  const data = Calculations.calculatePayoutsData(getPayoutsStateBundle(), getSelectedMonthKey());

  if (subtitle) {
    subtitle.textContent = `Wypłaty w ${formatMonthLabel(data.payoutMonth)} dotyczą pensji z ${formatMonthLabel(data.previousMonth)} i mogą uwzględniać zaległości oraz zaliczki do wskazanego dnia.`;
  }

  defaultDayInput.value = `${data.defaultDay}`;
  sourceMonthLabel.textContent = formatMonthLabel(data.previousMonth);
  defaultDateLabel.textContent = formatPayoutDateLabel(data.payoutDate);
  totalBaseEl.textContent = formatSettlementCompactCurrency(data.totalBaseAmount || 0);
  totalCarryoverEl.textContent = formatSettlementCompactCurrency(data.totalCarryoverAmount || 0);
  totalRemainingEl.textContent = formatSettlementCompactCurrency(data.totalRemainingAmount || 0);

  list.innerHTML = (data.employees || []).length > 0
    ? data.employees.map(item => buildPayoutEmployeeCardHtml(item, data)).join('')
    : '<p class="settlement-detail-empty">Brak pracowników z pensją, zaległościami lub zaliczkami do rozliczenia w tym miesiącu.</p>';

  normalizeCurrencySuffixSpacing(document.getElementById('payouts-view'));
}
