function getPayoutsStateBundle() {
  const exportState = Store.getExportData ? Store.getExportData() : { version: 'v3', common: {}, months: {} };
  return { ...exportState, selectedMonth: getSelectedMonthKey() };
}

function formatPayoutDateLabel(value = '') {
  if (!value) return '—';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleDateString('pl-PL', { year: 'numeric', month: '2-digit', day: '2-digit' });
}

function getPayoutEmployeeRecordState(item = {}) {
  if ((item.remainingAmount || 0) <= 0.005) return 'settled';
  if ((item.settledAmount || 0) > 0.005) return 'partial';
  return 'pending';
}

function getPayoutTypeLabel(type = '') {
  if (type === 'weekly') return 'Tygodniówka';
  if (type === 'custom') return 'Niestandardowa';
  return 'Miesięczna';
}

function buildPayoutHistoryEntryHtml(entry, personId, partnerMap) {
  const advancesTotal = entry.deductedAdvances.reduce((s, a) => s + (a.restoredToCosts ? 0 : a.amount), 0);
  const refundsPending = entry.advanceRefunds.filter(r => !r.returned);
  const refundsDone = entry.advanceRefunds.filter(r => r.returned);
  const payerName = entry.paidByPartnerId ? (partnerMap[entry.paidByPartnerId] || 'Nieznany') : '';
  const entryTotal = entry.cashAmount + advancesTotal;

  const advancesHtml = entry.deductedAdvances.length > 0 ? `
    <div class="payout-archived-advances">
      <div class="payout-section-label">Odliczone zaliczki:</div>
      ${entry.deductedAdvances.map(adv => {
        const giverName = adv.paidById ? (partnerMap[adv.paidById] || adv.paidById) : '';
        if (adv.restoredToCosts) {
          return `<div class="payout-archived-advance payout-archived-advance--restored">
            <span class="adv-date">${escapeReportHtml(adv.date ? formatPayoutDateLabel(adv.date) : '—')}</span>
            <span class="adv-name">${escapeReportHtml(adv.name || 'Zaliczka')}</span>
            <span class="adv-amount">${escapeReportHtml(formatSettlementCompactCurrency(adv.amount))}</span>
            ${giverName ? `<span class="adv-giver">dał: ${escapeReportHtml(giverName)}</span>` : ''}
            <span class="badge badge-working-partner" style="font-size:0.72rem;">Przywrócono do kosztów</span>
          </div>`;
        }
        return `<div class="payout-archived-advance" data-advance-id="${escapeReportHtml(adv.id)}">
          <span class="adv-date">${escapeReportHtml(adv.date ? formatPayoutDateLabel(adv.date) : '—')}</span>
          <span class="adv-name">${escapeReportHtml(adv.name || 'Zaliczka')}</span>
          <span class="adv-amount">${escapeReportHtml(formatSettlementCompactCurrency(adv.amount))}</span>
          ${giverName ? `<span class="adv-giver">dał: ${escapeReportHtml(giverName)}</span>` : ''}
          <button type="button" class="btn btn-secondary btn-xs btn-restore-advance"
            data-person-id="${escapeReportHtml(personId)}"
            data-entry-id="${escapeReportHtml(entry.id)}"
            data-advance-id="${escapeReportHtml(adv.id)}">Przywróć do kosztów</button>
        </div>`;
      }).join('')}
    </div>` : '';

  const refundsHtml = entry.advanceRefunds.length > 0 ? `
    <div class="payout-refunds-section">
      <div class="payout-section-label">Rozliczenia zaliczek między wspólnikami:</div>
      ${entry.advanceRefunds.map(ref => {
        const toName = ref.toPartnerId ? (partnerMap[ref.toPartnerId] || ref.toPartnerId) : '?';
        const fromName = entry.paidByPartnerId ? (partnerMap[entry.paidByPartnerId] || entry.paidByPartnerId) : '?';
        if (ref.returned) {
          return `<div class="payout-refund-row payout-refund-row--done">
            <span>✓ ${escapeReportHtml(fromName)} zwrócił ${escapeReportHtml(formatSettlementCompactCurrency(ref.amount))} → ${escapeReportHtml(toName)}</span>
            <span class="payout-card-note">${escapeReportHtml(formatHistoryTimestamp(ref.returnedAt))}</span>
          </div>`;
        }
        return `<div class="payout-refund-row payout-refund-row--pending">
          <span class="payout-refund-desc">⚠ ${escapeReportHtml(fromName)} winien ${escapeReportHtml(formatSettlementCompactCurrency(ref.amount))} → ${escapeReportHtml(toName)}</span>
          <button type="button" class="btn btn-secondary btn-xs btn-mark-refund-returned"
            data-person-id="${escapeReportHtml(personId)}"
            data-entry-id="${escapeReportHtml(entry.id)}"
            data-advance-id="${escapeReportHtml(ref.advanceId)}">Zwrócono</button>
        </div>`;
      }).join('')}
    </div>` : '';

  const sourceMonthLabel = entry.sourceMonth ? `za ${escapeReportHtml(formatMonthLabel(entry.sourceMonth))}` : '';
  const salaryRefundHtml = entry.salaryRefund?.toPartnerId
    ? (() => {
        const refundToName = entry.salaryRefund.toPartnerId ? (partnerMap[entry.salaryRefund.toPartnerId] || entry.salaryRefund.toPartnerId) : '?';
        if (entry.salaryRefund.returned) {
          return `<div class="payout-refund-row payout-refund-row--done">
            <span>✓ Rozliczono z ${escapeReportHtml(refundToName)}: ${escapeReportHtml(formatSettlementCompactCurrency(entry.salaryRefund.amount))}</span>
          </div>`;
        }
        return `<div class="payout-refunds-section">
          <div class="payout-section-label">Rozliczenie wynagrodzenia między wspólnikami:</div>
          <div class="payout-refund-row payout-refund-row--pending">
            <span class="payout-refund-desc">⚠ Należy zwrócić ${escapeReportHtml(formatSettlementCompactCurrency(entry.salaryRefund.amount))} → ${escapeReportHtml(refundToName)}</span>
            <button type="button" class="btn btn-secondary btn-xs btn-mark-salary-refund-returned"
              data-person-id="${escapeReportHtml(personId)}"
              data-entry-id="${escapeReportHtml(entry.id)}">Zwrócono</button>
          </div>
        </div>`;
      })()
    : '';

  return `
    <div class="payout-history-entry" data-entry-id="${escapeReportHtml(entry.id)}">
      <div class="payout-history-entry-header">
        <div class="payout-history-entry-meta">
          <span class="payout-history-entry-label">${escapeReportHtml(entry.label || getPayoutTypeLabel(entry.type))}</span>
          ${sourceMonthLabel ? `<span class="payout-source-month-tag">${sourceMonthLabel}</span>` : ''}
          ${entry.payoutDate ? `<span class="payout-history-entry-date">${escapeReportHtml(formatPayoutDateLabel(entry.payoutDate))}</span>` : ''}
          ${payerName ? `<span class="payout-history-entry-payer">płacił: ${escapeReportHtml(payerName)}</span>` : ''}
        </div>
        <div class="payout-history-entry-actions">
          <div class="payout-history-entry-amounts">
            ${entry.cashAmount > 0.005 ? `<span>gotówka: <strong>${escapeReportHtml(formatSettlementCompactCurrency(entry.cashAmount))}</strong></span>` : ''}
            ${advancesTotal > 0.005 ? `<span>zaliczki: <strong>${escapeReportHtml(formatSettlementCompactCurrency(advancesTotal))}</strong></span>` : ''}
            <span class="payout-history-entry-total">= ${escapeReportHtml(formatSettlementCompactCurrency(entryTotal))}</span>
          </div>
          <button type="button" class="btn btn-danger btn-xs btn-delete-payout"
            data-person-id="${escapeReportHtml(personId)}"
            data-entry-id="${escapeReportHtml(entry.id)}"
            title="Usuń wypłatę i przywróć zaliczki">✕</button>
        </div>
      </div>
      ${advancesHtml}
      ${refundsHtml}
      ${salaryRefundHtml}
    </div>`;
}

function buildPayoutHistoryHtml(payoutRecord, personId, partnerMap) {
  const payouts = payoutRecord?.payouts || [];
  if (payouts.length === 0) return '';

  const settledCash = payouts.reduce((s, p) => s + p.cashAmount, 0);
  const settledAdv = payouts.reduce((s, p) => s + p.deductedAdvances.filter(a => !a.restoredToCosts).reduce((ss, a) => ss + a.amount, 0), 0);
  const total = settledCash + settledAdv;

  // Group by type
  const byType = {};
  for (const p of payouts) {
    const t = p.type || 'monthly';
    if (!byType[t]) byType[t] = [];
    byType[t].push(p);
  }

  const groupsHtml = Object.entries(byType).map(([type, entries]) => {
    const groupTotal = entries.reduce((s, p) => s + p.cashAmount + p.deductedAdvances.filter(a => !a.restoredToCosts).reduce((ss, a) => ss + a.amount, 0), 0);
    const groupLabel = type === 'weekly' ? 'Tygodniówki' : type === 'custom' ? 'Niestandardowe' : 'Miesięczne';
    const entriesHtml = entries.map(e => buildPayoutHistoryEntryHtml(e, personId, partnerMap)).join('');

    if (entries.length === 1) {
      return entriesHtml;
    }
    return `
      <details class="payout-type-group">
        <summary class="payout-type-group-summary">
          <span>${escapeReportHtml(groupLabel)} (${entries.length})</span>
          <strong>${escapeReportHtml(formatSettlementCompactCurrency(groupTotal))}</strong>
        </summary>
        <div class="payout-type-group-entries">${entriesHtml}</div>
      </details>`;
  }).join('');

  return `
    <details class="payout-history-panel">
      <summary class="payout-history-panel-summary">
        <span>Historia wypłat (${payouts.length})</span>
        <strong>${escapeReportHtml(formatSettlementCompactCurrency(total))}</strong>
      </summary>
      <div class="payout-history-list">${groupsHtml}</div>
    </details>`;
}

function buildAdvancesChecklistHtml(allAdvances, defaultCheckedIds, remainingAmount) {
  if (allAdvances.length === 0) return '';

  const rows = allAdvances.map(adv => {
    const checked = defaultCheckedIds.has(adv.id);
    return `
      <label class="payout-advance-row">
        <input type="checkbox" class="payout-advance-checkbox"
          data-advance-id="${escapeReportHtml(adv.id)}"
          data-advance-amount="${parseFloat(adv.amount) || 0}"
          ${checked ? 'checked' : ''}>
        <span class="adv-date">${escapeReportHtml(formatPayoutDateLabel(adv.date))}</span>
        <span class="adv-name">${escapeReportHtml(adv.name || 'Zaliczka')}</span>
        <span class="adv-amount">${escapeReportHtml(formatSettlementCompactCurrency(adv.amount))}</span>
      </label>`;
  }).join('');

  const defaultTotal = allAdvances
    .filter(a => defaultCheckedIds.has(a.id))
    .reduce((s, a) => s + (parseFloat(a.amount) || 0), 0);

  return `
    <div class="payout-advances-section">
      <div class="payout-advances-header">
        <span class="payout-section-label">Zaliczki do odliczenia</span>
        <div class="payout-advances-header-actions">
          <button type="button" class="btn btn-secondary btn-xs btn-select-advances-to-date">Do dnia wypłaty</button>
          <button type="button" class="btn btn-secondary btn-xs btn-select-all-advances">Wszystkie</button>
          <button type="button" class="btn btn-secondary btn-xs btn-deselect-all-advances">Żadna</button>
        </div>
      </div>
      <div class="payout-advances-list">${rows}</div>
      <div class="payout-advances-total">
        Wybrane: <strong class="payout-selected-advances-total">${escapeReportHtml(formatSettlementCompactCurrency(defaultTotal))}</strong>
      </div>
    </div>`;
}

function buildPayoutEmployeeCardHtml(item = {}, data = {}) {
  const personId = item?.person?.id || '';
  const stateKey = getPayoutEmployeeRecordState(item);
  const statusMeta = stateKey === 'settled'
    ? { label: 'Rozliczono', className: 'badge-working-partner' }
    : (stateKey === 'partial'
        ? { label: 'Rozliczono częściowo', className: 'badge-partner' }
        : { label: 'Do rozliczenia', className: 'badge-employee' });

  // Build partner map for name lookups
  const partnerMap = {};
  (data.partners || []).forEach(p => { partnerMap[p.id] = getPersonDisplayName(p); });

  const historyHtml = buildPayoutHistoryHtml(item.payoutRecord, personId, partnerMap);

  const advancesHtml = item.remainingAmount > 0.005
    ? buildAdvancesChecklistHtml(
        item.allAvailableAdvanceExpenses || [],
        item.defaultCheckedIds || item.defaultCheckedAdvanceIds || new Set(),
        item.remainingAmount
      )
    : '';

  // Partner options for payer select — default = employee's employer
  const defaultPayerId = item.defaultPayerId || '';
  const partnerOptions = (data.partners || []).map(p =>
    `<option value="${escapeReportHtml(p.id)}" ${p.id === defaultPayerId ? 'selected' : ''}>${escapeReportHtml(getPersonDisplayName(p))}</option>`
  ).join('');

  // Next weekly payout number
  const weeklyCount = (item.payoutRecord?.payouts || []).filter(p => p.type === 'weekly').length;
  const defaultLabel = `Tygodniówka ${weeklyCount + 1}`;

  const firstDay = data?.payoutMonth ? `${data.payoutMonth}-01` : '';
  const [y, mo] = (data?.payoutMonth || '').split('-').map(Number);
  const lastDay = (y && mo) ? `${data.payoutMonth}-${String(new Date(y, mo, 0).getDate()).padStart(2, '0')}` : '';

  const formHtml = item.remainingAmount > 0.005 ? `
    <div class="payout-form-section">
      <div class="payout-section-label" style="margin-bottom:0.75rem;">Nowa wypłata</div>
      <div class="payout-form-grid">
        <div class="form-group" style="margin-bottom:0;">
          <label>Typ</label>
          <select class="payout-type-select">
            <option value="monthly">Miesięczna</option>
            <option value="weekly" selected>Tygodniówka</option>
            <option value="custom">Niestandardowa</option>
          </select>
        </div>
        <div class="form-group payout-label-group" style="margin-bottom:0; display:none;">
          <label>Nazwa / Opis</label>
          <input type="text" class="payout-label-input" placeholder="${escapeReportHtml(defaultLabel)}">
        </div>
        <div class="form-group" style="margin-bottom:0;">
          <label>Data wypłaty</label>
          <input type="date" class="payout-date-input" value="${escapeReportHtml(item.payoutDate || '')}" min="${firstDay}" max="${lastDay}">
        </div>
        <div class="form-group" style="margin-bottom:0;">
          <label>Kwota gotówka / przelew</label>
          <input type="number" step="0.01" min="0" class="payout-cash-input" placeholder="0,00">
        </div>
        ${partnerOptions ? `
        <div class="form-group" style="margin-bottom:0;">
          <label>Kto płaci?</label>
          <select class="payout-payer-select">
            <option value="">— nie określono —</option>
            ${partnerOptions}
          </select>
        </div>` : ''}
      </div>
      <div class="payout-form-summary">
        <span>Gotówka: <strong class="payout-form-cash-display">0,00 zł</strong></span>
        <span>+ Zaliczki: <strong class="payout-form-advances-display">0,00 zł</strong></span>
        <span>= Łącznie: <strong class="payout-form-total-display">0,00 zł</strong></span>
        <span class="payout-form-remaining-hint">Pozostaje: <strong class="payout-form-remaining-display">${escapeReportHtml(formatSettlementCompactCurrency(item.remainingAmount))}</strong></span>
      </div>
      <div class="payout-form-buttons">
        <button type="button" class="btn btn-primary btn-payout-confirm">Zatwierdź wypłatę</button>
      </div>
    </div>` : '';

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
        ${(item.carryoverAmount || 0) > 0.005 ? `
        <div class="glass-panel payout-summary-box">
          <div class="payout-summary-label">Nierozliczone wypłaty</div>
          <div class="payout-summary-value payout-summary-value--warning">${escapeReportHtml(formatSettlementCompactCurrency(item.carryoverAmount || 0))}</div>
        </div>` : ''}
        ${(item.currentMonthAmount || 0) > 0.005 ? `
        <div class="glass-panel payout-summary-box">
          <div class="payout-summary-label">Bieżący miesiąc</div>
          <div class="payout-summary-value">${escapeReportHtml(formatSettlementCompactCurrency(item.currentMonthAmount || 0))}</div>
        </div>` : ''}
        ${(item.settledAmount || 0) > 0.005 ? `
        <div class="glass-panel payout-summary-box">
          <div class="payout-summary-label">Łącznie wypłacono</div>
          <div class="payout-summary-value payout-summary-value--success">${escapeReportHtml(formatSettlementCompactCurrency(item.settledAmount || 0))}</div>
        </div>` : ''}
        <div class="glass-panel payout-summary-box">
          <div class="payout-summary-label">Pozostało do rozliczenia</div>
          <div class="payout-summary-value ${(item.remainingAmount || 0) <= 0.005 ? 'payout-summary-value--success' : 'payout-summary-value--warning'}">${escapeReportHtml(formatSettlementCompactCurrency(item.remainingAmount || 0))}</div>
        </div>
      </div>

      <div class="payout-config-row">
        <label class="choice-option payout-choice-option">
          <input type="checkbox" class="payout-include-carryover-toggle" ${item.payoutRecord?.includeCarryover !== false ? 'checked' : ''}>
          <span class="choice-option-text">Dolicz nierozliczone wypłaty z poprzednich miesięcy</span>
        </label>
        <label class="choice-option payout-choice-option">
          <input type="checkbox" class="payout-include-current-month-toggle" ${item.includeCurrentMonth ? 'checked' : ''}>
          <span class="choice-option-text">Dolicz wypłatę za bieżący miesiąc (tygodniówki)</span>
        </label>
      </div>

      ${historyHtml}
      ${advancesHtml}
      ${formHtml}
    </article>`;
}

function getSelectedAdvanceIds(card) {
  const ids = [];
  card.querySelectorAll('.payout-advance-checkbox:checked').forEach(cb => {
    ids.push(cb.dataset.advanceId);
  });
  return ids;
}

function getSelectedAdvancesTotal(card) {
  let total = 0;
  card.querySelectorAll('.payout-advance-checkbox:checked').forEach(cb => {
    total += parseFloat(cb.dataset.advanceAmount) || 0;
  });
  return total;
}

function updatePayoutFormSummary(card) {
  const cashInput = card.querySelector('.payout-cash-input');
  const cashVal = Math.max(0, parseFloat(cashInput?.value) || 0);
  const advancesVal = getSelectedAdvancesTotal(card);
  const total = cashVal + advancesVal;

  const fmt = v => formatSettlementCompactCurrency(v);
  const cashDisplay = card.querySelector('.payout-form-cash-display');
  const advDisplay = card.querySelector('.payout-form-advances-display');
  const totalDisplay = card.querySelector('.payout-form-total-display');
  if (cashDisplay) cashDisplay.textContent = fmt(cashVal);
  if (advDisplay) advDisplay.textContent = fmt(advancesVal);
  if (totalDisplay) totalDisplay.textContent = fmt(total);
}

function settlePayoutForEmployee(personId = '') {
  const currentMonth = getSelectedMonthKey();
  const payoutsData = Calculations.calculatePayoutsData(getPayoutsStateBundle(), currentMonth);
  const item = (payoutsData.employees || []).find(entry => entry?.person?.id === personId);
  const card = document.querySelector(`.payout-employee-card[data-person-id="${CSS.escape(personId)}"]`);
  if (!item || !card) return;

  if ((item.remainingAmount || 0) <= 0.005) {
    alert('Ta wypłata jest już rozliczona.');
    return;
  }

  const cashInput = card.querySelector('.payout-cash-input');
  const cashAmount = Math.max(0, parseFloat(cashInput?.value) || 0);
  const advanceExpenseIds = getSelectedAdvanceIds(card);
  const advancesTotal = getSelectedAdvancesTotal(card);

  if (cashAmount <= 0 && advancesTotal <= 0) {
    alert('Wpisz kwotę wypłaty lub zaznacz zaliczki do odliczenia.');
    return;
  }

  const totalThisPayout = cashAmount + advancesTotal;
  if (totalThisPayout > item.remainingAmount + 0.005) {
    alert(`Kwota ${formatSettlementCompactCurrency(totalThisPayout)} przekracza pozostałe do rozliczenia ${formatSettlementCompactCurrency(item.remainingAmount)}.`);
    return;
  }

  const typeSelect = card.querySelector('.payout-type-select');
  const payoutType = typeSelect?.value || 'monthly';
  const labelInput = card.querySelector('.payout-label-input');

  // Auto-generate label
  const existingPayouts = item.payoutRecord?.payouts || [];
  let payoutLabel = labelInput?.value?.trim() || '';
  if (!payoutLabel) {
    if (payoutType === 'weekly') {
      const weeklyCount = existingPayouts.filter(p => p.type === 'weekly').length;
      payoutLabel = `Tygodniówka ${weeklyCount + 1}`;
    } else if (payoutType === 'monthly') {
      payoutLabel = 'Wypłata miesięczna';
    }
  }

  const dateInput = card.querySelector('.payout-date-input');
  const payerSelect = card.querySelector('.payout-payer-select');

  // sourceMonth: current month for weekly (tygodniówki), previous month for standard
  const sourceMonth = item.includeCurrentMonth ? item.payoutMonth : item.previousMonth;

  const success = Store.settleEmployeePayout?.(personId, {
    payoutType,
    payoutLabel,
    payoutDate: dateInput?.value || item.payoutDate || '',
    paidByPartnerId: payerSelect?.value || '',
    cashAmount,
    advanceExpenseIds,
    sourceMonth,
    baseAmountSnapshot: item.baseAmount,
    carryoverAmountSnapshot: item.carryoverAmount,
    plannedAmountSnapshot: item.plannedAmount
  }, currentMonth);

  if (!success) {
    alert('Nie udało się zapisać wypłaty.');
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

    if (event.target.classList.contains('payout-include-current-month-toggle')) {
      Store.updatePayoutEmployeeConfig?.(personId, { includeCurrentMonth: event.target.checked }, getSelectedMonthKey());
      return;
    }

    if (event.target.classList.contains('payout-type-select')) {
      const labelGroup = card.querySelector('.payout-label-group');
      if (labelGroup) labelGroup.style.display = event.target.value === 'custom' ? '' : 'none';
      return;
    }

    // Advance checkbox change → update summary
    if (event.target.classList.contains('payout-advance-checkbox')) {
      updatePayoutFormSummary(card);
      return;
    }

    if (event.target.classList.contains('payout-cash-input')) {
      updatePayoutFormSummary(card);
    }
  });

  list.addEventListener('input', (event) => {
    const card = event.target.closest('.payout-employee-card');
    if (!card) return;
    if (event.target.classList.contains('payout-cash-input')) {
      updatePayoutFormSummary(card);
    }
  });

  list.addEventListener('click', (event) => {
    const card = event.target.closest('.payout-employee-card');
    const personId = card?.getAttribute('data-person-id') || '';

    // Confirm payout
    if (event.target.closest('.btn-payout-confirm')) {
      settlePayoutForEmployee(personId);
      return;
    }

    // Select advances to payout date
    if (event.target.closest('.btn-select-advances-to-date')) {
      const payoutsData = Calculations.calculatePayoutsData(getPayoutsStateBundle(), getSelectedMonthKey());
      const item = (payoutsData.employees || []).find(e => e?.person?.id === personId);
      const payoutDate = card.querySelector('.payout-date-input')?.value || item?.payoutDate || '';
      card.querySelectorAll('.payout-advance-checkbox').forEach(cb => {
        const advId = cb.dataset.advanceId;
        const adv = (item?.allAvailableAdvanceExpenses || []).find(a => a.id === advId);
        cb.checked = !!(adv && (!payoutDate || adv.date <= payoutDate));
      });
      updatePayoutFormSummary(card);
      return;
    }

    // Select all advances
    if (event.target.closest('.btn-select-all-advances')) {
      card.querySelectorAll('.payout-advance-checkbox').forEach(cb => { cb.checked = true; });
      updatePayoutFormSummary(card);
      return;
    }

    // Deselect all advances
    if (event.target.closest('.btn-deselect-all-advances')) {
      card.querySelectorAll('.payout-advance-checkbox').forEach(cb => { cb.checked = false; });
      updatePayoutFormSummary(card);
      return;
    }

    // Delete payout entry
    const deleteBtn = event.target.closest('.btn-delete-payout');
    if (deleteBtn) {
      const pId = deleteBtn.dataset.personId;
      const entryId = deleteBtn.dataset.entryId;
      if (pId && entryId && confirm('Usunąć wypłatę? Odliczone zaliczki zostaną przywrócone do bieżącego miesiąca.')) {
        const ok = Store.deletePayoutEntry?.(pId, entryId, getSelectedMonthKey());
        if (!ok) alert('Nie udało się usunąć wypłaty.');
      }
      return;
    }

    // Restore advance to costs
    const restoreBtn = event.target.closest('.btn-restore-advance');
    if (restoreBtn) {
      const pId = restoreBtn.dataset.personId;
      const entryId = restoreBtn.dataset.entryId;
      const advId = restoreBtn.dataset.advanceId;
      if (pId && entryId && advId) {
        const ok = Store.restoreAdvanceToCosts?.(pId, entryId, advId, getSelectedMonthKey());
        if (!ok) alert('Nie udało się przywrócić zaliczki do kosztów.');
      }
      return;
    }

    // Mark advance refund as returned
    const refundBtn = event.target.closest('.btn-mark-refund-returned');
    if (refundBtn) {
      const pId = refundBtn.dataset.personId;
      const entryId = refundBtn.dataset.entryId;
      const advId = refundBtn.dataset.advanceId;
      if (pId && entryId && advId) {
        const ok = Store.markAdvanceRefundReturned?.(pId, entryId, advId, getSelectedMonthKey());
        if (!ok) alert('Nie udało się oznaczyć zwrotu.');
      }
      return;
    }

    // Mark salary refund as returned (cross-partner)
    const salaryRefundBtn = event.target.closest('.btn-mark-salary-refund-returned');
    if (salaryRefundBtn) {
      const pId = salaryRefundBtn.dataset.personId;
      const entryId = salaryRefundBtn.dataset.entryId;
      if (pId && entryId) {
        const ok = Store.markSalaryRefundReturned?.(pId, entryId, getSelectedMonthKey());
        if (!ok) alert('Nie udało się oznaczyć zwrotu.');
      }
      return;
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
    subtitle.textContent = `Wypłaty w ${formatMonthLabel(data.payoutMonth)} dotyczą pensji z ${formatMonthLabel(data.previousMonth)} i mogą uwzględniać zaległości oraz zaliczki.`;
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

  // Render separate companies section
  const sepSection = document.getElementById('payouts-separate-companies-section');
  if (sepSection) sepSection.innerHTML = buildPayoutSeparateCompaniesSectionHtml(getPayoutsStateBundle(), data.payoutMonth);

  // Render global history
  const histSection = document.getElementById('payouts-global-history-section');
  if (histSection) histSection.innerHTML = buildPayoutGlobalHistoryHtml(getPayoutsStateBundle());

  normalizeCurrencySuffixSpacing(document.getElementById('payouts-view'));

  // Initialize form summaries
  list.querySelectorAll('.payout-employee-card').forEach(card => updatePayoutFormSummary(card));

  // Event delegation for separate companies toggles
  if (sepSection && !sepSection.dataset.bound) {
    sepSection.dataset.bound = 'true';
    sepSection.addEventListener('change', (event) => {
      if (event.target.classList.contains('payout-separate-company-toggle')) {
        const companyId = event.target.dataset.companyId;
        if (companyId) {
          Store.updateSeparateCompanyPayoutsConfig?.(companyId, { enablePayouts: event.target.checked }, getSelectedMonthKey());
        }
      }
    });
  }
}

function buildPayoutSeparateCompaniesSectionHtml(state, payoutMonth) {
  const allPersons = state?.common?.persons || state?.persons || [];
  const separateCompanies = allPersons.filter(p => p.type === 'SEPARATE_COMPANY' && p.isActive !== false);
  if (separateCompanies.length === 0) return '';

  const payoutSettings = Calculations.getPayoutSettings(state, payoutMonth);
  const rows = separateCompanies.map(company => {
    const enabled = payoutSettings.separateCompanies?.[company.id]?.enablePayouts === true;
    const empCount = allPersons.filter(p => p.type === 'EMPLOYEE' && p.employerId === company.id && p.isActive !== false).length;
    return `
      <label class="payout-separate-company-row choice-option">
        <input type="checkbox" class="payout-separate-company-toggle" data-company-id="${escapeReportHtml(company.id)}" ${enabled ? 'checked' : ''}>
        <span class="choice-option-text">
          <strong>${escapeReportHtml(getPersonDisplayName(company))}</strong>
          <span class="payout-card-note">${empCount} pracownik${empCount === 1 ? '' : empCount < 5 ? 'ów' : 'ów'}</span>
        </span>
      </label>`;
  }).join('');

  return `
    <details class="payout-separate-companies-panel">
      <summary class="payout-history-panel-summary">
        <span>Wypłaty dla pracowników osobnych firm</span>
        <span class="payout-card-note">domyślnie wyłączone</span>
      </summary>
      <div class="payout-separate-companies-list">${rows}</div>
    </details>`;
}

function buildPayoutGlobalHistoryHtml(state) {
  const allPersons = state?.common?.persons || state?.persons || [];
  const personById = Object.fromEntries(allPersons.map(p => [p.id, p]));
  const partnerMap = {};
  allPersons.filter(p => p.type === 'PARTNER' || p.type === 'WORKING_PARTNER').forEach(p => {
    partnerMap[p.id] = getPersonDisplayName(p);
  });

  // Collect all payout entries from all months
  const allEntries = [];
  Object.entries(state?.months || {}).forEach(([payoutMonth, monthData]) => {
    Object.entries(monthData?.payouts?.employees || {}).forEach(([personId, empRecord]) => {
      if (!empRecord || !Array.isArray(empRecord.payouts)) return;
      const person = personById[personId];
      empRecord.payouts.forEach(entry => {
        if (!entry) return;
        const advTotal = Array.isArray(entry.deductedAdvances)
          ? entry.deductedAdvances.filter(a => !a.restoredToCosts).reduce((s, a) => s + (a.amount || 0), 0)
          : 0;
        allEntries.push({
          personId,
          person,
          payoutMonth,
          entry,
          total: (entry.cashAmount || 0) + advTotal
        });
      });
    });
  });

  if (allEntries.length === 0) return '';

  // Sort by payoutDate desc, then createdAt desc
  allEntries.sort((a, b) => {
    const d = (b.entry.payoutDate || b.entry.createdAt || '').localeCompare(a.entry.payoutDate || a.entry.createdAt || '');
    return d !== 0 ? d : (b.entry.createdAt || '').localeCompare(a.entry.createdAt || '');
  });

  const rowsHtml = allEntries.map(({ person, entry, total }) => {
    const name = person ? escapeReportHtml(getPersonDisplayName(person)) : '—';
    const sourceLabel = entry.sourceMonth ? escapeReportHtml(formatMonthLabel(entry.sourceMonth)) : '—';
    const payoutDateLabel = entry.payoutDate ? escapeReportHtml(formatPayoutDateLabel(entry.payoutDate)) : '—';
    const typeLabel = escapeReportHtml(entry.label || getPayoutTypeLabel(entry.type));
    const payerName = entry.paidByPartnerId ? escapeReportHtml(partnerMap[entry.paidByPartnerId] || entry.paidByPartnerId) : '—';
    const totalFmt = escapeReportHtml(formatSettlementCompactCurrency(total));
    const advCount = Array.isArray(entry.deductedAdvances) ? entry.deductedAdvances.filter(a => !a.restoredToCosts).length : 0;
    return `<tr>
      <td>${name}</td>
      <td>${sourceLabel}</td>
      <td>${payoutDateLabel}</td>
      <td>${typeLabel}</td>
      <td class="text-right">${totalFmt}</td>
      <td>${advCount > 0 ? `${advCount} zal.` : '—'}</td>
      <td>${payerName}</td>
    </tr>`;
  }).join('');

  return `
    <details class="payout-global-history-panel">
      <summary class="payout-history-panel-summary">
        <span>Historia wszystkich wypłat (${allEntries.length})</span>
      </summary>
      <div class="payout-global-history-wrap">
        <table class="payout-global-history-table">
          <thead>
            <tr>
              <th>Pracownik</th>
              <th>Za miesiąc</th>
              <th>Data wypłaty</th>
              <th>Typ</th>
              <th>Kwota</th>
              <th>Zaliczki</th>
              <th>Płatnik</th>
            </tr>
          </thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>
    </details>`;
