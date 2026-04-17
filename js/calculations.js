const Calculations = {
  _wrapState: (state, month) => {
    if (!state || !state.version || state.version !== 'v3') return state;
    const selectedMonth = month || Calculations.getSelectedMonth(state);
    const m = (state.months || {})[selectedMonth] || {};
    return {
      ...state.common,
      ...m,
      selectedMonth,
      months: state.months,
      common: state.common,
      version: state.version
    };
  },

  getSelectedMonth: (state) => state?.selectedMonth || `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}`,

  isDateInMonth: (value, month) => !!value && !!month && value.startsWith(`${month}-`),

  sheetMatchesSelectedMonth: (sheet, selectedMonth) => sheet?.month === selectedMonth,

  getMonthSettings: (state, month = null) => {
    const selectedMonth = month || Calculations.getSelectedMonth(state);
    const directMonthSettings = state?.months?.[selectedMonth]?.monthSettings;
    if (directMonthSettings && typeof directMonthSettings === 'object' && !Array.isArray(directMonthSettings)) {
      return directMonthSettings;
    }

    const monthSettings = state?.monthSettings || {};
    if (monthSettings[selectedMonth] && typeof monthSettings[selectedMonth] === 'object' && !Array.isArray(monthSettings[selectedMonth])) {
      return monthSettings[selectedMonth];
    }

    if ((monthSettings.persons || monthSettings.clients || monthSettings.settlementConfig || monthSettings.personContractCharges || monthSettings.payouts || monthSettings.invoices)
      && typeof monthSettings === 'object'
      && !Array.isArray(monthSettings)) {
      return monthSettings;
    }

    return {
      persons: {},
      clients: {},
      settlementConfig: {},
      personContractCharges: {},
      payouts: { defaultDay: 15, employees: {} },
      invoices: { issueDate: '', emailIntro: '', issued: false, issuedAt: '', issuedSnapshot: null, clients: {}, extraInvoices: [] }
    };
  },

  getPreviousMonthKey: (month = '') => {
    if (!/^\d{4}-\d{2}$/.test((month || '').toString().trim())) return '';
    const [year, monthNumber] = month.split('-').map(Number);
    const previousMonthDate = new Date(year, (monthNumber || 1) - 2, 1);
    return `${previousMonthDate.getFullYear()}-${String(previousMonthDate.getMonth() + 1).padStart(2, '0')}`;
  },

  getPayoutSettings: (state, month = null) => {
    const monthSettings = Calculations.getMonthSettings(state, month);
    const payouts = monthSettings?.payouts || {};
    return {
      defaultDay: Math.max(1, Math.min(31, parseInt(payouts.defaultDay, 10) || 15)),
      employees: Object.entries(payouts.employees || {}).reduce((acc, [personId, record]) => {
        const normalizedPersonId = (personId || '').toString().trim();
        if (!normalizedPersonId) return acc;
        acc[normalizedPersonId] = {
          includeCarryover: record?.includeCarryover !== false,
          includeCurrentMonth: record?.includeCurrentMonth === true,
          deductAdvancesMode: ['none', 'default-day', 'custom-day'].includes(record?.deductAdvancesMode)
            ? record.deductAdvancesMode
            : 'default-day',
          customDeductionDate: /^\d{4}-\d{2}-\d{2}$/.test((record?.customDeductionDate || '').toString().trim())
            ? record.customDeductionDate.toString().trim()
            : '',
          sourceMonth: /^\d{4}-\d{2}$/.test((record?.sourceMonth || '').toString().trim())
            ? record.sourceMonth.toString().trim()
            : '',
          baseAmountSnapshot: Math.max(0, parseFloat(record?.baseAmountSnapshot) || 0),
          carryoverAmountSnapshot: Math.max(0, parseFloat(record?.carryoverAmountSnapshot) || 0),
          plannedAmountSnapshot: Math.max(0, parseFloat(record?.plannedAmountSnapshot) || 0),
          advanceDeductionAmountSnapshot: Math.max(0, parseFloat(record?.advanceDeductionAmountSnapshot) || 0),
          settledCashAmount: Math.max(0, parseFloat(record?.settledCashAmount) || 0),
          settledAdvanceAmount: Math.max(0, parseFloat(record?.settledAdvanceAmount) || 0),
          removedAdvanceExpenseIds: Array.isArray(record?.removedAdvanceExpenseIds)
            ? [...new Set(record.removedAdvanceExpenseIds.map(value => (value || '').toString().trim()).filter(Boolean))]
            : [],
          lastSettledAt: (record?.lastSettledAt || '').toString().trim(),
          payouts: Array.isArray(record?.payouts) ? record.payouts : []
        };
        return acc;
      }, {})
    };
  },

  getPayoutDefaultDate: (month = '', defaultDay = 15) => {
    if (!/^\d{4}-\d{2}$/.test((month || '').toString().trim())) return '';
    const [year, monthNumber] = month.split('-').map(Number);
    const lastDay = new Date(year, monthNumber, 0).getDate();
    const resolvedDay = Math.max(1, Math.min(lastDay, parseInt(defaultDay, 10) || 15));
    return `${month}-${String(resolvedDay).padStart(2, '0')}`;
  },

  getEmployeeUnresolvedPayoutTotal: (state, payoutMonth = '', personId = '') => {
    if (!payoutMonth || !personId) return 0;

    // Cache settlements computed during this call to avoid redundant work
    const settlementCache = {};
    const getOrComputeSettlement = (sMonth) => {
      if (!sMonth || !state?.months?.[sMonth]) return null;
      if (!settlementCache[sMonth]) {
        try {
          settlementCache[sMonth] = Calculations.generateSettlement(
            { ...state, selectedMonth: sMonth },
            { includeInvoiceTaxEqualization: false, includeInvoiceReconciliation: false }
          );
        } catch (e) {
          settlementCache[sMonth] = null;
        }
      }
      return settlementCache[sMonth];
    };

    return Object.keys(state?.months || {})
      .filter(monthKey => /^\d{4}-\d{2}$/.test(monthKey) && monthKey < payoutMonth)
      .reduce((sum, monthKey) => {
        const record = Calculations.getPayoutSettings({ ...state, selectedMonth: monthKey }, monthKey)?.employees?.[personId];
        const settledAmount = record
          ? (Math.max(0, parseFloat(record.settledCashAmount) || 0) + Math.max(0, parseFloat(record.settledAdvanceAmount) || 0))
          : 0;

        // Skip if this month's carryover is explicitly disabled
        if (record?.includeCarryover === false) return sum;

        let plannedAmount;
        if (record && (parseFloat(record.plannedAmountSnapshot) || 0) > 0.005) {
          // Use locked-in snapshot (settlement was started for this month)
          plannedAmount = Math.max(0, parseFloat(record.plannedAmountSnapshot));
        } else {
          // Dynamically compute: salary for payout month M comes from settlement of month M-1
          const salarySourceMonth = Calculations.getPreviousMonthKey(monthKey);
          if (!salarySourceMonth) return sum;
          const settlement = getOrComputeSettlement(salarySourceMonth);
          if (!settlement) return sum;
          const empEntry = (settlement.employees || []).find(e => e?.person?.id === personId);
          plannedAmount = Math.max(0, parseFloat(empEntry?.toPayout) || 0);
        }

        return sum + Math.max(0, plannedAmount - settledAmount);
      }, 0);
  },

  getEmployeePayoutAdvances: (state, payoutMonth = '', personId = '', cutoffDate = '') => {
    if (!payoutMonth || !personId || !cutoffDate) return { amount: 0, expenses: [] };
    const scopedState = Calculations._wrapState({ ...state, selectedMonth: payoutMonth }, payoutMonth);
    const expenses = (scopedState?.expenses || [])
      .filter(expense => expense?.type === 'ADVANCE' && expense?.advanceForId === personId && (expense?.date || '') <= cutoffDate)
      .sort((left, right) => (left?.date || '').localeCompare(right?.date || '', 'pl-PL'));

    return {
      amount: expenses.reduce((sum, expense) => sum + (parseFloat(expense.amount) || 0), 0),
      expenses
    };
  },

  selectAdvanceExpensesForPayout: (expenses = [], maxAmount = 0) => {
    const selectedExpenses = [];
    let totalAmount = 0;
    const limit = Math.max(0, parseFloat(maxAmount) || 0);

    for (const expense of (expenses || [])) {
      const amount = Math.max(0, parseFloat(expense?.amount) || 0);
      if (!(amount > 0)) continue;
      if ((totalAmount + amount) - limit > 0.005) continue;
      selectedExpenses.push(expense);
      totalAmount += amount;
    }

    return {
      amount: totalAmount,
      expenses: selectedExpenses
    };
  },

  calculatePayoutsData: (state, month = null) => {
    const payoutMonth = month || Calculations.getSelectedMonth(state);
    const previousMonth = Calculations.getPreviousMonthKey(payoutMonth);
    const payoutSettings = Calculations.getPayoutSettings(state, payoutMonth);
    const currentPayoutDate = Calculations.getPayoutDefaultDate(payoutMonth, payoutSettings.defaultDay);
    const [pyear, pmonthNum] = payoutMonth.split('-').map(Number);
    const lastDayOfPayoutMonth = `${payoutMonth}-${String(new Date(pyear, pmonthNum, 0).getDate()).padStart(2, '0')}`;
    const previousMonthSettlement = previousMonth
      ? Calculations.generateSettlement({ ...state, selectedMonth: previousMonth }, { includeInvoiceTaxEqualization: false, includeInvoiceReconciliation: false })
      : { employees: [] };
    const previousMonthEmployeeMap = Object.fromEntries((previousMonthSettlement?.employees || []).map(entry => [entry?.person?.id, entry]));
    const currentMonthScopedState = Calculations._wrapState({ ...state, selectedMonth: payoutMonth }, payoutMonth);
    const employeeIds = new Set([
      ...Object.keys(previousMonthEmployeeMap),
      ...Object.keys(payoutSettings.employees || {}),
      ...((currentMonthScopedState?.expenses || [])
        .filter(expense => expense?.type === 'ADVANCE' && expense?.advanceForId)
        .map(expense => expense.advanceForId))
    ]);

    // Partners list for the payer selector
    const allPersons = state?.common?.persons || state?.persons || [];
    const partners = allPersons
      .filter(p => (p.type === 'PARTNER' || p.type === 'WORKING_PARTNER') && p.isActive !== false)
      .sort((a, b) => Calculations.getPersonDisplayName(a).localeCompare(Calculations.getPersonDisplayName(b), 'pl-PL'));

    // Lazily computed current-month settlement (shared across employees)
    let currentMonthSettlement = null;

    const employees = [...employeeIds]
      .map(personId => {
        const person = allPersons.find(candidate => candidate.id === personId && candidate.type === 'EMPLOYEE');
        if (!person) return null;

        const payoutRecord = payoutSettings.employees?.[personId] || {
          includeCarryover: true,
          includeCurrentMonth: false,
          deductAdvancesMode: 'default-day',
          customDeductionDate: '',
          settledCashAmount: 0,
          settledAdvanceAmount: 0,
          plannedAmountSnapshot: 0,
          baseAmountSnapshot: 0,
          carryoverAmountSnapshot: 0,
          advanceDeductionAmountSnapshot: 0,
          removedAdvanceExpenseIds: [],
          payouts: []
        };
        const baseAmount = Math.max(0, parseFloat(previousMonthEmployeeMap?.[personId]?.toPayout) || 0);
        const carryoverAmount = Calculations.getEmployeeUnresolvedPayoutTotal(state, payoutMonth, personId);

        // Current payout month salary (always dynamic — for weekly payouts mid-month)
        const includeCurrentMonth = payoutRecord.includeCurrentMonth === true;
        let currentMonthAmount = 0;
        if (includeCurrentMonth) {
          if (!currentMonthSettlement) {
            currentMonthSettlement = Calculations.generateSettlement(
              { ...state, selectedMonth: payoutMonth },
              { includeInvoiceTaxEqualization: false, includeInvoiceReconciliation: false }
            );
          }
          const currentEntry = (currentMonthSettlement.employees || []).find(e => e?.person?.id === personId);
          currentMonthAmount = Math.max(0, parseFloat(currentEntry?.toPayout) || 0);
        }

        // Base + carryover snapshotted at first payout; currentMonthAmount is always live
        const snapshotBase = payoutRecord.plannedAmountSnapshot > 0
          ? payoutRecord.plannedAmountSnapshot
          : (baseAmount + (payoutRecord.includeCarryover !== false ? carryoverAmount : 0));
        const plannedAmount = snapshotBase + currentMonthAmount;

        const resolvedBaseAmount = payoutRecord.baseAmountSnapshot > 0 ? payoutRecord.baseAmountSnapshot : baseAmount;
        const resolvedCarryoverAmount = payoutRecord.plannedAmountSnapshot > 0
          ? payoutRecord.carryoverAmountSnapshot
          : (payoutRecord.includeCarryover !== false ? carryoverAmount : 0);
        const deductionDate = payoutRecord.deductAdvancesMode === 'none'
          ? ''
          : (payoutRecord.deductAdvancesMode === 'custom-day' && payoutRecord.customDeductionDate
              ? payoutRecord.customDeductionDate
              : currentPayoutDate);

        // All available advances (not yet removed), sorted by date
        const removedIds = new Set(payoutRecord.removedAdvanceExpenseIds || []);
        const allAdvancesForMonth = Calculations.getEmployeePayoutAdvances(
          state, payoutMonth, personId, lastDayOfPayoutMonth
        ).expenses.filter(e => !removedIds.has(e.id));

        // Default-checked advances based on deductAdvancesMode
        const defaultCheckedIds = new Set(
          payoutRecord.deductAdvancesMode === 'none'
            ? []
            : allAdvancesForMonth
                .filter(e => !deductionDate || e.date <= deductionDate)
                .map(e => e.id)
        );

        const settledAmount = Math.max(0, parseFloat(payoutRecord.settledCashAmount) || 0) + Math.max(0, parseFloat(payoutRecord.settledAdvanceAmount) || 0);
        const remainingAmount = Math.max(0, plannedAmount - settledAmount);

        // Auto-selected advances (for default payout amount hint)
        const selectableAdvanceData = payoutRecord.deductAdvancesMode === 'none'
          ? { amount: 0, expenses: [] }
          : Calculations.selectAdvanceExpensesForPayout(
              allAdvancesForMonth.filter(e => !deductionDate || e.date <= deductionDate),
              remainingAmount
            );
        const availableAdvanceAmount = selectableAdvanceData.amount;

        return {
          person,
          previousMonth,
          payoutMonth,
          payoutDate: currentPayoutDate,
          baseAmount: resolvedBaseAmount,
          carryoverAmount: resolvedCarryoverAmount,
          currentMonthAmount,
          includeCurrentMonth,
          plannedAmount,
          settledCashAmount: Math.max(0, parseFloat(payoutRecord.settledCashAmount) || 0),
          settledAdvanceAmount: Math.max(0, parseFloat(payoutRecord.settledAdvanceAmount) || 0),
          settledAmount,
          remainingAmount,
          deductAdvancesMode: payoutRecord.deductAdvancesMode,
          deductionDate,
          availableAdvanceAmount,
          availableAdvanceExpenses: selectableAdvanceData.expenses,
          allAvailableAdvanceExpenses: allAdvancesForMonth,
          defaultCheckedAdvanceIds: defaultCheckedIds,
          totalAdvanceAmountToDate: allAdvancesForMonth.reduce((s, e) => s + (parseFloat(e.amount) || 0), 0),
          payoutRecord,
          payoutNowAmount: Math.max(0, remainingAmount - availableAdvanceAmount)
        };
      })
      .filter(Boolean)
      .filter(item => item.baseAmount > 0 || item.carryoverAmount > 0 || item.currentMonthAmount > 0 || item.remainingAmount > 0 || item.availableAdvanceAmount > 0 || item.allAvailableAdvanceExpenses.length > 0 || (item.payoutRecord?.payouts?.length || 0) > 0)
      .sort((left, right) => Calculations.getPersonDisplayName(left.person).localeCompare(Calculations.getPersonDisplayName(right.person), 'pl-PL'));

    return {
      payoutMonth,
      previousMonth,
      defaultDay: payoutSettings.defaultDay,
      payoutDate: currentPayoutDate,
      employees,
      partners,
      totalBaseAmount: employees.reduce((sum, item) => sum + item.baseAmount, 0),
      totalCarryoverAmount: employees.reduce((sum, item) => sum + item.carryoverAmount, 0),
      totalRemainingAmount: employees.reduce((sum, item) => sum + item.remainingAmount, 0)
    };
  },

  getSettlementConfig: (state, month = null) => {
    const selectedMonth = month || Calculations.getSelectedMonth(state);
    const monthSettings = Calculations.getMonthSettings(state, selectedMonth);
    const monthTaxRate = parseFloat(monthSettings?.settlementConfig?.taxRate);
    const monthZusFixedAmount = parseFloat(monthSettings?.settlementConfig?.zusFixedAmount);

    return {
      taxRate: Number.isFinite(monthTaxRate) ? monthTaxRate : (parseFloat(state?.config?.taxRate) || 0),
      zusFixedAmount: Number.isFinite(monthZusFixedAmount) ? monthZusFixedAmount : (parseFloat(state?.config?.zusFixedAmount) || 0)
    };
  },

  isPersonActiveInMonth: (person, state, month = null) => {
    if (!person) return false;
    const monthSettings = Calculations.getMonthSettings(state, month);
    if (monthSettings.persons?.[person.id] !== undefined) {
      return monthSettings.persons[person.id] !== false;
    }
    return person.isActive !== false;
  },

  isClientActiveInMonth: (client, state, month = null) => {
    if (!client) return false;
    const monthSettings = Calculations.getMonthSettings(state, month);
    if (monthSettings.clients?.[client.id] !== undefined) {
      return monthSettings.clients[client.id] !== false;
    }
    return client.isActive !== false;
  },

  getActivePersons: (state, month = null) => {
    const selectedMonth = month || Calculations.getSelectedMonth(state);
    return (state?.persons || []).filter(person => Calculations.isPersonActiveInMonth(person, state, selectedMonth));
  },

  personParticipatesInCosts: (person) => {
    if (!person) return false;
    if (person.type === 'PARTNER') return person.participatesInCosts !== false;
    if (person.type === 'SEPARATE_COMPANY') return person.participatesInCosts !== false;
    if (person.type === 'WORKING_PARTNER') return person.participatesInCosts === true;
    return false;
  },

  clampPercent: (value, fallback = 100) => {
    const parsed = parseFloat(value);
    if (!Number.isFinite(parsed)) return fallback;
    return Math.max(0, Math.min(100, parsed));
  },

  isSeparateCompany: (person) => person?.type === 'SEPARATE_COMPANY',

  isPartnerLike: (person) => person?.type === 'PARTNER' || person?.type === 'WORKING_PARTNER' || person?.type === 'SEPARATE_COMPANY',

  getPersonDisplayName: (person) => {
    if (!person) return '';
    return [person.name, person.lastName].filter(value => !!value && value.toString().trim() !== '').join(' ').trim();
  },

  isEmployeeProfitRecipient: (person) => person?.type === 'PARTNER' || person?.type === 'SEPARATE_COMPANY',

  getEmployeeProfitRecipientWeight: (person) => {
    if (!person) return 0;
    if (person.type === 'PARTNER') return 100;
    if (person.type === 'SEPARATE_COMPANY' && person.receivesCompanyEmployeeProfits === true) {
      return Calculations.clampPercent(person.companyEmployeeProfitSharePercent, 100);
    }
    return 0;
  },

  getSeparateCompanySharedProfitPercent: (person) => {
    if (!Calculations.isSeparateCompany(person) || person.sharesEmployeeProfits !== true) return 0;
    return Calculations.clampPercent(person.employeeProfitSharePercent, 100);
  },

  countsEmployeeAccountingRefund: (person) => {
    if (!Calculations.isSeparateCompany(person)) return false;
    return person.countsEmployeeAccountingRefund !== false;
  },

  personHasSettlementActivity: (personId, state, month = null) => {
    const selectedMonth = month || Calculations.getSelectedMonth(state);

    const hasLegacyData = (state?.dayRecords || []).some(day => {
      if (!Calculations.isDateInMonth(day?.date, selectedMonth)) return false;
      const entry = day?.entries?.[personId];
      return !!(entry && entry.isPresent);
    });
    if (hasLegacyData) return true;

    const hasMonthlySheetData = (state?.monthlySheets || []).some(sheet => {
      if (!Calculations.sheetMatchesSelectedMonth(sheet, selectedMonth)) return false;
      return Object.values(sheet?.days || {}).some(day =>
        day?.hours && day.hours[personId] !== undefined && day.hours[personId] !== ''
      );
    });
    if (hasMonthlySheetData) return true;

    const hasWorksSheetData = (state?.worksSheets || []).some(sheet => {
      if (!Calculations.sheetMatchesSelectedMonth(sheet, selectedMonth)) return false;

      if ((sheet?.activePersons || []).includes(personId)) return true;
      if ((sheet?.partnerProfitOverrides || []).includes(personId)) return true;

      return (sheet?.entries || []).some(entry =>
        entry?.hours && entry.hours[personId] !== undefined && entry.hours[personId] !== ''
      );
    });
    if (hasWorksSheetData) return true;

    return (state?.expenses || []).some(expense =>
      Calculations.isDateInMonth(expense?.date, selectedMonth)
      && (expense?.paidById === personId || expense?.advanceForId === personId)
    );
  },

  getSettlementPersons: (state, month = null) => {
    const selectedMonth = month || Calculations.getSelectedMonth(state);
    return (state?.persons || []).filter(person =>
      Calculations.isPersonActiveInMonth(person, state, selectedMonth)
      || Calculations.personHasSettlementActivity(person.id, state, selectedMonth)
    );
  },

  normalizeText: (value) => (value || '').toString().trim().toLowerCase(),

  getClientByName: (state, clientName) => {
    const normalizedName = Calculations.normalizeText(clientName);
    if (!normalizedName || !state.clients) return null;
    return state.clients.find(client => Calculations.normalizeText(client.name) === normalizedName) || null;
  },

  getSheetClientRate: (sheet, state) => {
    if (sheet && !isNaN(parseFloat(sheet.clientRateOverride)) && parseFloat(sheet.clientRateOverride) > 0) {
      return parseFloat(sheet.clientRateOverride);
    }

    const client = Calculations.getClientByName(state, sheet?.client);
    if (client && !isNaN(parseFloat(client.hourlyRate)) && parseFloat(client.hourlyRate) > 0) {
      return parseFloat(client.hourlyRate);
    }
    return 0;
  },

  getSheetPersonRate: (person, sheet, state, fallbackRate = 0) => {
    const config = (sheet?.personsConfig && sheet.personsConfig[person.id]) ? sheet.personsConfig[person.id] : {};
    if (config.customRate !== undefined && config.customRate !== null && !isNaN(parseFloat(config.customRate))) {
      return parseFloat(config.customRate);
    }

    if (person.hourlyRate !== undefined && person.hourlyRate !== null && !isNaN(parseFloat(person.hourlyRate)) && parseFloat(person.hourlyRate) > 0) {
      return parseFloat(person.hourlyRate);
    }

    const clientRate = Calculations.getSheetClientRate(sheet, state);
    return clientRate > 0 ? clientRate : (fallbackRate || 0);
  },

  isPersonInactiveInMonthlySheetDay: (sheet, personId, dayNumber) => {
    if (!sheet?.days || !personId) return false;

    let isInactive = false;
    for (let day = 1; day <= dayNumber; day++) {
      const override = sheet?.days?.[day]?.activityOverrides?.[personId];
      if (override === 'inactive') {
        isInactive = true;
      } else if (override === 'active') {
        isInactive = false;
      }
    }

    return isInactive;
  },

  isPersonIncludedInMonthlySheet: (personId, sheet) => {
    if (!personId || !sheet) return false;
    if (Array.isArray(sheet.activePersons)) {
      return sheet.activePersons.includes(personId);
    }
    return true;
  },

  getSheetTotalHours: (sheet, personIds = null) => {
    if (!sheet?.days) return 0;

    return Object.entries(sheet.days).reduce((sheetSum, [dayKey, day]) => {
      if (!day?.hours) return sheetSum;

      const dayHours = Object.entries(day.hours).reduce((daySum, [personId, value]) => {
        if (personIds && !personIds.has(personId)) return daySum;
        if (!Calculations.isPersonIncludedInMonthlySheet(personId, sheet)) return daySum;
        if (Calculations.isPersonInactiveInMonthlySheetDay(sheet, personId, parseInt(dayKey, 10))) return daySum;
        const parsed = parseFloat(value);
        return daySum + (isNaN(parsed) ? 0 : parsed);
      }, 0);

      return sheetSum + dayHours;
    }, 0);
  },

  isRoboczogodzinyEntry: (entry, state) => {
    const workDef = (state?.worksCatalog || []).find(work => work.id === entry?.workId);
    return !!(workDef && workDef.coreId === 'c_roboczogodziny');
  },

  getWorksEntryMetrics: (entry, sheet, state) => {
    const isRoboczogodziny = Calculations.isRoboczogodzinyEntry(entry, state);
    const clientRate = Calculations.getSheetClientRate(sheet, state);
    let totalLoggedHours = 0;
    let employeeLoggedHours = 0;
    let employeeCost = 0;

    (sheet?.activePersons || []).forEach(personId => {
      if (!entry?.hours || typeof entry.hours[personId] === 'undefined' || entry.hours[personId] === '') return;

      const hours = parseFloat(entry.hours[personId]);
      if (!Number.isFinite(hours) || hours <= 0) return;

      totalLoggedHours += hours;

      const person = (state?.persons || []).find(candidate => candidate.id === personId);
      if (person && person.type === 'EMPLOYEE') {
        employeeLoggedHours += hours;
        employeeCost += hours * Calculations.getSheetPersonRate(person, sheet, state);
      }
    });

    const quantity = parseFloat(entry?.quantity) || 0;
    const price = parseFloat(entry?.price) || 0;
    const revenue = isRoboczogodziny ? totalLoggedHours * clientRate : quantity * price;
    const profitBaseRevenue = isRoboczogodziny ? employeeLoggedHours * clientRate : revenue;
    const profit = profitBaseRevenue - employeeCost;

    return {
      isRoboczogodziny,
      clientRate,
      totalLoggedHours,
      employeeLoggedHours,
      revenue,
      employeeCost,
      profitBaseRevenue,
      profit
    };
  },

  calculateEmployeeRevenueFromHourBasedWork: (person, state) => {
    const selectedMonth = Calculations.getSelectedMonth(state);
    const monthlyRevenue = (state.monthlySheets || []).reduce((sum, sheet) => {
      if (!Calculations.sheetMatchesSelectedMonth(sheet, selectedMonth)) return sum;
      if (!Calculations.isPersonIncludedInMonthlySheet(person.id, sheet)) return sum;

      const clientRate = Calculations.getSheetClientRate(sheet, state);
      if (!(clientRate > 0)) return sum;

      const personHours = Object.entries(sheet.days || {}).reduce((hoursSum, [dayKey, day]) => {
        if (!day?.hours || day.hours[person.id] === undefined || day.hours[person.id] === '') return hoursSum;
        if (Calculations.isPersonInactiveInMonthlySheetDay(sheet, person.id, parseInt(dayKey, 10))) return hoursSum;
        const hours = parseFloat(day.hours[person.id]);
        return hoursSum + (Number.isFinite(hours) ? hours : 0);
      }, 0);

      return sum + (personHours * clientRate);
    }, 0);

    const worksRevenue = (state.worksSheets || []).reduce((sum, sheet) => {
      if (!Calculations.sheetMatchesSelectedMonth(sheet, selectedMonth)) return sum;

      return sum + (sheet.entries || []).reduce((entrySum, entry) => {
        if (!Calculations.isRoboczogodzinyEntry(entry, state)) return entrySum;

        const clientRate = Calculations.getSheetClientRate(sheet, state);
        if (!(clientRate > 0)) return entrySum;

        const hours = parseFloat(entry?.hours?.[person.id]);
        if (!Number.isFinite(hours) || hours <= 0) return entrySum;

        return entrySum + (hours * clientRate);
      }, 0);
    }, 0);

    return monthlyRevenue + worksRevenue;
  },

  getAdjustedRevenueRateForPerson: (person, sheet, state) => {
    const clientRate = Calculations.getSheetClientRate(sheet, state);
    if (!(clientRate > 0)) {
      return { clientRate: 0, adjustedRate: 0, rateDifference: 0 };
    }

    let separateCompany = null;
    if (Calculations.isSeparateCompany(person)) {
      separateCompany = person;
    } else if (person?.employerId) {
      separateCompany = (state?.persons || []).find(candidate => candidate.id === person.employerId && Calculations.isSeparateCompany(candidate)) || null;
    }

    if (!separateCompany) {
      return { clientRate, adjustedRate: clientRate, rateDifference: 0 };
    }

    const separateCompanyRate = Calculations.getSheetPersonRate(separateCompany, sheet, state, clientRate);
    if (Number.isFinite(separateCompanyRate) && separateCompanyRate > 0 && separateCompanyRate < clientRate) {
      return {
        clientRate,
        adjustedRate: separateCompanyRate,
        rateDifference: clientRate - separateCompanyRate
      };
    }

    return { clientRate, adjustedRate: clientRate, rateDifference: 0 };
  },

  calculatePersonRevenueFromMonthlySheets: (person, state, options = {}) => {
    const selectedMonth = Calculations.getSelectedMonth(state);
    const useAdjustedRate = options.useAdjustedRate === true;
    return (state.monthlySheets || []).reduce((sum, sheet) => {
      if (!Calculations.sheetMatchesSelectedMonth(sheet, selectedMonth)) return sum;
      if (!Calculations.isPersonIncludedInMonthlySheet(person.id, sheet)) return sum;

      const rateInfo = Calculations.getAdjustedRevenueRateForPerson(person, sheet, state);
      const revenueRate = useAdjustedRate ? rateInfo.adjustedRate : rateInfo.clientRate;
      if (!(revenueRate > 0)) return sum;

      const personHours = Object.values(sheet.days || {}).reduce((hoursSum, day) => {
        if (!day?.hours || day.hours[person.id] === undefined || day.hours[person.id] === '') return hoursSum;
        const hours = parseFloat(day.hours[person.id]);
        return hoursSum + (Number.isFinite(hours) ? hours : 0);
      }, 0);

      return sum + (personHours * revenueRate);
    }, 0);
  },

  calculatePersonRevenueFromWorksRoboczogodziny: (person, state, options = {}) => {
    const selectedMonth = Calculations.getSelectedMonth(state);
    const useAdjustedRate = options.useAdjustedRate === true;
    return (state.worksSheets || []).reduce((sum, sheet) => {
      if (!Calculations.sheetMatchesSelectedMonth(sheet, selectedMonth)) return sum;

      return sum + (sheet.entries || []).reduce((entrySum, entry) => {
        if (!Calculations.isRoboczogodzinyEntry(entry, state)) return entrySum;

        const rateInfo = Calculations.getAdjustedRevenueRateForPerson(person, sheet, state);
        const revenueRate = useAdjustedRate ? rateInfo.adjustedRate : rateInfo.clientRate;
        if (!(revenueRate > 0)) return entrySum;

        const hours = parseFloat(entry?.hours?.[person.id]);
        if (!Number.isFinite(hours) || hours <= 0) return entrySum;

        return entrySum + (hours * revenueRate);
      }, 0);
    }, 0);
  },

  calculateRevenueBreakdown: (state, settlementPersons, profitSharePersons = settlementPersons) => {
    const selectedMonth = Calculations.getSelectedMonth(state);
    const settlementPersonIds = new Set(settlementPersons.map(person => person.id));
    const settlementEmployeeIds = new Set(settlementPersons.filter(person => person.type === 'EMPLOYEE').map(person => person.id));
    const profitSharePersonIds = new Set((profitSharePersons || []).map(person => person.id));

    let totalRevenue = 0;
    let employeeRevenue = 0;
    let worksRevenue = 0;
    const worksRevenueByPartner = {};

    if (state.dayRecords) {
      state.dayRecords.forEach(day => {
        if (!Calculations.isDateInMonth(day.date, selectedMonth)) return;
        Object.entries(day.entries || {}).forEach(([personId, entry]) => {
          if (!settlementPersonIds.has(personId) || !entry?.isPresent) return;

          const hours = entry.hoursReconciled !== undefined
            ? parseFloat(entry.hoursReconciled)
            : Calculations.calculateHours(entry.startTime || day.globalStartTime, entry.endTime || day.globalEndTime);
          const safeHours = isNaN(hours) ? 0 : hours;
          const revenue = 0;

          totalRevenue += revenue;
          if (settlementEmployeeIds.has(personId)) {
            employeeRevenue += revenue;
          }
        });
      });
    }

    if (state.monthlySheets) {
      state.monthlySheets.forEach(sheet => {
        if (!Calculations.sheetMatchesSelectedMonth(sheet, selectedMonth)) return;
        const clientRate = Calculations.getSheetClientRate(sheet, state);
        const sheetTotalHours = Calculations.getSheetTotalHours(sheet, settlementPersonIds);
        const sheetEmployeeHours = Calculations.getSheetTotalHours(sheet, settlementEmployeeIds);

        totalRevenue += sheetTotalHours * clientRate;
        employeeRevenue += sheetEmployeeHours * clientRate;
      });
    }

    if (state.worksSheets) {
      state.worksSheets.forEach(sheet => {
        if (!Calculations.sheetMatchesSelectedMonth(sheet, selectedMonth)) return;
        
        let sheetRev = 0;
        let sheetProfit = 0;

        (sheet.entries || []).forEach(e => {
          const metrics = Calculations.getWorksEntryMetrics(e, sheet, state);
          sheetRev += metrics.revenue;
          sheetProfit += metrics.profit;
        });
        
        totalRevenue += sheetRev;
        const sheetNetProfit = sheetProfit;
        worksRevenue += sheetNetProfit;

         const allPartners = state.persons.filter(p => Calculations.isPartnerLike(p) && profitSharePersonIds.has(p.id));
        const eligiblePartners = allPartners.filter(p => {
          const isParticipant = sheet.activePersons && sheet.activePersons.includes(p.id);
          const hasOverride = p.type === 'PARTNER' && sheet.partnerProfitOverrides && sheet.partnerProfitOverrides.includes(p.id);
          return isParticipant || hasOverride;
        });

        let shareGroup = eligiblePartners.length > 0 ? eligiblePartners : allPartners;
        
        if (shareGroup.length > 0) {
           const shareAmount = sheetNetProfit / shareGroup.length;
           shareGroup.forEach(p => {
             worksRevenueByPartner[p.id] = (worksRevenueByPartner[p.id] || 0) + shareAmount;
           });
        }
      });
    }

    return { totalRevenue, employeeRevenue, worksRevenue, worksRevenueByPartner };
  },

  calculateHours: (startTime, endTime) => {
    if (!startTime || !endTime) return 0;
    
    const [sH, sM] = startTime.split(':').map(Number);
    const [eH, eM] = endTime.split(':').map(Number);
    
    let diffMinutes = (eH * 60 + eM) - (sH * 60 + sM);
    if (diffMinutes < 0) {
      diffMinutes += 24 * 60;
    }
    return diffMinutes / 60;
  },

  getPersonActiveWorkDaysCount: (personId, state, month = null) => {
    if (!personId) return 0;

    const selectedMonth = month || Calculations.getSelectedMonth(state);
    const activeDayKeys = new Set();

    (state?.dayRecords || []).forEach(day => {
      if (!Calculations.isDateInMonth(day?.date, selectedMonth)) return;

      const entry = day?.entries?.[personId];
      if (!entry?.isPresent) return;

      const hours = entry?.hoursReconciled !== undefined
        ? parseFloat(entry.hoursReconciled)
        : Calculations.calculateHours(entry?.startTime || day?.globalStartTime, entry?.endTime || day?.globalEndTime);

      if (Number.isFinite(hours) && hours > 0) {
        activeDayKeys.add(day.date);
      }
    });

    (state?.monthlySheets || []).forEach(sheet => {
      if (!Calculations.sheetMatchesSelectedMonth(sheet, selectedMonth)) return;
      if (!Calculations.isPersonIncludedInMonthlySheet(personId, sheet)) return;

      Object.entries(sheet?.days || {}).forEach(([dayKey, day]) => {
        if (!day?.hours || day.hours[personId] === undefined || day.hours[personId] === '') return;
        if (Calculations.isPersonInactiveInMonthlySheetDay(sheet, personId, parseInt(dayKey, 10))) return;

        const hours = parseFloat(day.hours[personId]);
        if (Number.isFinite(hours) && hours > 0) {
          activeDayKeys.add(`${selectedMonth}-${String(dayKey).padStart(2, '0')}`);
        }
      });
    });

    (state?.worksSheets || []).forEach(sheet => {
      if (!Calculations.sheetMatchesSelectedMonth(sheet, selectedMonth)) return;

      (sheet?.entries || []).forEach(entry => {
        if (!Calculations.isDateInMonth(entry?.date, selectedMonth)) return;

        const hours = parseFloat(entry?.hours?.[personId]);
        if (Number.isFinite(hours) && hours > 0) {
          activeDayKeys.add(entry.date);
        }
      });
    });

    return activeDayKeys.size;
  },

  getDietaCalculationMode: (expense) => {
    if (expense?.type !== 'DIETA') return 'FIXED';

    const rawMode = (expense?.dietaCalculationMode || '').toString().trim().toUpperCase();
    if (rawMode === 'ACTIVE_DAYS' || rawMode === 'MANUAL_DAYS' || rawMode === 'FIXED') {
      return rawMode;
    }

    return expense?.dietaByActiveDays === true ? 'ACTIVE_DAYS' : 'FIXED';
  },

  isDietaCountedByActiveDays: (expense) => Calculations.getDietaCalculationMode(expense) === 'ACTIVE_DAYS',

  isDietaCountedByManualDays: (expense) => Calculations.getDietaCalculationMode(expense) === 'MANUAL_DAYS',

  getExpenseEffectiveDays: (expense, state, month = null) => {
    if (Calculations.isDietaCountedByManualDays(expense)) {
      return Math.max(0, parseInt(expense?.dietaDaysCount, 10) || 0);
    }

    if (!Calculations.isDietaCountedByActiveDays(expense)) return 0;

    const activeWorkDays = Calculations.getPersonActiveWorkDaysCount(expense?.advanceForId, state, month);
    const daysAdjustment = parseInt(expense?.dietDaysAdjustment, 10) || 0;

    return Math.max(0, activeWorkDays + daysAdjustment);
  },

  getExpenseEffectiveAmount: (expense, state, month = null) => {
    const baseAmount = parseFloat(expense?.amount) || 0;
    if (expense?.type !== 'DIETA') return baseAmount;
    if (!Calculations.isDietaCountedByActiveDays(expense) && !Calculations.isDietaCountedByManualDays(expense)) return baseAmount;

    return Number((Calculations.getExpenseEffectiveDays(expense, state, month) * baseAmount).toFixed(2));
  },

  isRefundSharedWithAllPartners: (expense) => {
    if (expense?.type !== 'REFUND') return false;
    return (expense?.advanceForId || '').toString().trim() === 'all_partners';
  },

  getRefundRecipientIds: (expense, state, month = null) => {
    if (expense?.type !== 'REFUND') return [];

    if (Calculations.isRefundSharedWithAllPartners(expense)) {
      return Calculations.getActivePersons(state, month)
        .filter(person => Calculations.isPartnerLike(person) && Calculations.personParticipatesInCosts(person))
        .map(person => person.id);
    }

    const recipientId = (expense?.advanceForId || '').toString().trim();
    return recipientId ? [recipientId] : [];
  },

  getDistributedAmountForPerson: (totalAmount = 0, recipientIds = [], personId = '') => {
    if (!personId || !Array.isArray(recipientIds) || recipientIds.length === 0) return 0;

    const normalizedRecipientIds = [...new Set(recipientIds.map(value => (value || '').toString().trim()).filter(Boolean))];
    const recipientIndex = normalizedRecipientIds.indexOf((personId || '').toString().trim());
    if (recipientIndex === -1) return 0;

    const totalCents = Math.max(0, Math.round((parseFloat(totalAmount) || 0) * 100));
    const baseCents = Math.floor(totalCents / normalizedRecipientIds.length);
    const remainderCents = totalCents % normalizedRecipientIds.length;
    const cents = baseCents + (recipientIndex < remainderCents ? 1 : 0);

    return cents / 100;
  },

  getRefundReceivedAmountForPerson: (expense, personId = '', state, month = null) => {
    if (expense?.type !== 'REFUND' || !personId) return 0;

    const recipientIds = Calculations.getRefundRecipientIds(expense, state, month);
    if (!recipientIds.includes(personId) || recipientIds.length === 0) return 0;

    return Calculations.getDistributedAmountForPerson(
      Calculations.getExpenseEffectiveAmount(expense, state, month),
      recipientIds,
      personId
    );
  },

  calculateAccountingTax: (grossAmount, taxRate) => {
    const taxableBase = Math.max(parseFloat(grossAmount) || 0, 0);
    return taxableBase * (parseFloat(taxRate) || 0);
  },

  calculateNetAfterAccounting: (grossAmount, taxAmount, zusAmount) => {
    return (parseFloat(grossAmount) || 0) - (parseFloat(taxAmount) || 0) - (parseFloat(zusAmount) || 0);
  },

  getPersonContractCharges: (person, state = null, month = null) => {
    const usesContractCharges = person?.type === 'EMPLOYEE' || person?.type === 'WORKING_PARTNER';
    if (!usesContractCharges) {
      return {
        taxAmount: 0,
        zusAmount: 0,
        total: 0,
        paidByEmployer: false
      };
    }

    const monthSettings = state ? Calculations.getMonthSettings(state, month) : null;
    const personMonthOverride = monthSettings?.personContractCharges?.[person.id] || {};
    const taxAmount = parseFloat(personMonthOverride.contractTaxAmount ?? person?.contractTaxAmount) || 0;
    const zusAmount = parseFloat(personMonthOverride.contractZusAmount ?? person?.contractZusAmount) || 0;
    return {
      taxAmount,
      zusAmount,
      total: taxAmount + zusAmount,
      paidByEmployer: person?.contractChargesPaidByEmployer === true
    };
  },

  calculatePersonStats: (person, state, fallbackRate = 0, options = {}) => {
    const selectedMonth = Calculations.getSelectedMonth(state);
    const {
      includeLegacy = true,
      includeMonthlySheets = true,
      includeWorksSheets = true,
      worksOnlyRoboczogodziny = false
    } = options;
    let totalHours = 0;
    let totalSalary = 0;

    if (includeLegacy && state.dayRecords) {
      let legacyHours = 0;
      state.dayRecords.forEach(day => {
        if (!Calculations.isDateInMonth(day.date, selectedMonth)) return;
        const entry = day.entries[person.id];
        if (entry && entry.isPresent) {
          if (entry.hoursReconciled !== undefined) {
            legacyHours += entry.hoursReconciled;
          } else {
            legacyHours += Calculations.calculateHours(entry.startTime || day.globalStartTime, entry.endTime || day.globalEndTime);
          }
        }
      });
      totalHours += legacyHours;
      totalSalary += legacyHours * (person.hourlyRate || fallbackRate);
    }

    if (includeMonthlySheets && state.monthlySheets) {
      state.monthlySheets.forEach(sheet => {
        if (!Calculations.sheetMatchesSelectedMonth(sheet, selectedMonth)) return;
        if (!Calculations.isPersonIncludedInMonthlySheet(person.id, sheet)) return;
        let sheetHours = 0;
        const effectiveRate = Calculations.getSheetPersonRate(person, sheet, state, fallbackRate);

        if (sheet.days) {
          Object.entries(sheet.days).forEach(([dayKey, day]) => {
            if (day?.hours && day.hours[person.id] !== undefined && day.hours[person.id] !== '') {
              if (Calculations.isPersonInactiveInMonthlySheetDay(sheet, person.id, parseInt(dayKey, 10))) return;
              const h = parseFloat(day.hours[person.id]);
              if (!isNaN(h)) {
                sheetHours += h;
              }
            }
          });
        }
        
        totalHours += sheetHours;
        totalSalary += sheetHours * effectiveRate;
      });
    }

    if (includeWorksSheets && state.worksSheets) {
      state.worksSheets.forEach(sheet => {
         if (!Calculations.sheetMatchesSelectedMonth(sheet, selectedMonth)) return;
         
         if (sheet.activePersons && sheet.activePersons.includes(person.id)) {
            let sheetWorksHours = 0;
            const effectiveRate = Calculations.getSheetPersonRate(person, sheet, state, fallbackRate);

            (sheet.entries || []).forEach(e => {
                if (worksOnlyRoboczogodziny && !Calculations.isRoboczogodzinyEntry(e, state)) return;
               if (e.hours && typeof e.hours[person.id] !== 'undefined' && e.hours[person.id] !== '') {
                  const h = parseFloat(e.hours[person.id]);
                  if (!isNaN(h)) {
                     sheetWorksHours += h;
                  }
               }
            });

            totalHours += sheetWorksHours;
            totalSalary += sheetWorksHours * effectiveRate;
         }
      });
    }

    return { totalHours, totalSalary };
  },

  formatMonthLabel: (monthKey) => {
    if (!monthKey || !monthKey.includes('-')) return monthKey || '';
    const [year, month] = monthKey.split('-').map(Number);
    const date = new Date(year, (month || 1) - 1, 1);
    return date.toLocaleDateString('pl-PL', { month: 'long', year: 'numeric' });
  },

  formatInvoiceCurrency: (value) => {
    const amount = parseFloat(value) || 0;
    return `${amount.toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}\u00A0zł`;
  },

  getMonthLastDate: (monthKey) => {
    if (!monthKey || !monthKey.includes('-')) return '';
    const [year, month] = monthKey.split('-').map(Number);
    if (!Number.isFinite(year) || !Number.isFinite(month)) return '';
    const lastDay = new Date(year, month, 0).getDate();
    return `${year}-${String(month).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
  },

  getInvoiceEligibleIssuers: (state, month = null) => {
    return Calculations.getSettlementPersons(state, month).filter(person =>
      person?.type === 'PARTNER'
      || person?.type === 'SEPARATE_COMPANY'
      || (person?.type === 'WORKING_PARTNER' && !person?.employerId)
    );
  },

  getInvoiceIssuerForPerson: (personId, state, eligibleIssuerIds = null) => {
    const person = (state?.persons || []).find(candidate => candidate.id === personId);
    if (!person) return null;

    const canUse = (candidateId) => !eligibleIssuerIds || eligibleIssuerIds.has(candidateId);
    if (Calculations.isPartnerLike(person) && canUse(person.id)) return person.id;
    if (person.employerId && canUse(person.employerId)) return person.employerId;
    return null;
  },

  getInvoiceClientIdentity: (state, clientName) => {
    const client = Calculations.getClientByName(state, clientName);
    const normalizedName = (client?.name || clientName || 'Nieznany klient').toString().trim();
    const clientId = client?.id || `name:${Calculations.normalizeText(normalizedName)}`;

    return {
      clientId,
      clientName: normalizedName,
      fullCompanyName: (client?.fullCompanyName || normalizedName).toString().trim(),
      address: (client?.address || '').toString().trim(),
      nip: (client?.nip || '').toString().trim(),
      krs: (client?.krs || '').toString().trim(),
      regon: (client?.regon || '').toString().trim()
    };
  },

  getInvoiceModeLabel: (mode) => {
    if (mode === 'SETTLEMENT_REVENUE') return 'Według przychodów z rozliczenia';
    if (mode === 'EQUAL_SPLIT') return 'Po równo';
    if (mode === 'PERCENTAGE_SPLIT') return 'Procentowo między zaznaczonych';
    if (mode === 'BALANCE_INVOICE_SUMS') return 'Wyrównywanie Sumy Faktur';
    if (mode === 'MANUAL') return 'Ręcznie';
    return 'Według przychodów z arkuszy';
  },

  normalizeInvoicePercentageAllocations: (issuerIds, percentageAllocations = {}) => {
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
  },

  getDeterministicInvoiceRandom: (seed) => {
    const text = (seed || '').toString();
    let hash = 2166136261;
    for (let i = 0; i < text.length; i++) {
      hash ^= text.charCodeAt(i);
      hash = Math.imul(hash, 16777619);
    }
    return ((hash >>> 0) % 1000000) / 1000000;
  },

  buildEqualSplitAllocations: (totalAmount, issuerIds, options = {}) => {
    if (!Array.isArray(issuerIds) || issuerIds.length === 0) return {};

    const equalAmount = issuerIds.length > 0 ? totalAmount / issuerIds.length : 0;
    const shouldRandomize = options.randomize === true && issuerIds.length > 1;
    const maxVariance = Math.max(0, Math.min(Math.abs(equalAmount), parseFloat(options.varianceAmount) || 0));

    if (!shouldRandomize || !(maxVariance > 0)) {
      return Object.fromEntries(issuerIds.map(issuerId => [issuerId, equalAmount]));
    }

    const rawOffsets = issuerIds.map(issuerId => Calculations.getDeterministicInvoiceRandom(`${options.seed || ''}|${issuerId}`) - 0.5);
    const avgOffset = rawOffsets.reduce((sum, value) => sum + value, 0) / rawOffsets.length;
    const centeredOffsets = rawOffsets.map(value => value - avgOffset);
    const maxAbsOffset = Math.max(...centeredOffsets.map(value => Math.abs(value)), 0);
    const scale = maxAbsOffset > 0 ? maxVariance / maxAbsOffset : 0;
    const allocations = {};
    let assignedTotal = 0;

    issuerIds.forEach((issuerId, index) => {
      if (index === issuerIds.length - 1) {
        allocations[issuerId] = Number((totalAmount - assignedTotal).toFixed(2));
        return;
      }

      const amount = Number((equalAmount + (centeredOffsets[index] * scale)).toFixed(2));
      allocations[issuerId] = amount;
      assignedTotal += amount;
    });

    return allocations;
  },

  buildPercentageSplitAllocations: (totalAmount, issuerIds, percentageAllocations = {}) => {
    if (!Array.isArray(issuerIds) || issuerIds.length === 0) return {};

    const normalizedPercentages = Calculations.normalizeInvoicePercentageAllocations(issuerIds, percentageAllocations);
    const allocations = {};
    let assignedTotal = 0;

    issuerIds.forEach((issuerId, index) => {
      if (index === issuerIds.length - 1) {
        allocations[issuerId] = Number((totalAmount - assignedTotal).toFixed(2));
        return;
      }

      const amount = Number((totalAmount * ((normalizedPercentages[issuerId] || 0) / 100)).toFixed(2));
      allocations[issuerId] = amount;
      assignedTotal += amount;
    });

    return allocations;
  },

  buildBalancedInvoiceSumAllocations: (totalAmount, issuerIds, yearToDateTotals = {}) => {
    if (!Array.isArray(issuerIds) || issuerIds.length === 0) return {};

    const allocations = Object.fromEntries(issuerIds.map(issuerId => [issuerId, 0]));
    let remainingAmount = parseFloat(totalAmount) || 0;
    if (!(remainingAmount > 0)) {
      return Object.fromEntries(issuerIds.map(issuerId => [issuerId, remainingAmount / issuerIds.length || 0]));
    }

    const levels = issuerIds
      .map(issuerId => ({
        issuerId,
        total: parseFloat(yearToDateTotals?.[issuerId]) || 0
      }))
      .sort((a, b) => a.total - b.total);

    for (let index = 0; index < levels.length && remainingAmount > 0; index++) {
      const currentGroup = levels.slice(0, index + 1);
      const nextLevel = levels[index + 1]?.total;

      if (!Number.isFinite(nextLevel)) {
        const equalShare = remainingAmount / currentGroup.length;
        currentGroup.forEach(item => {
          allocations[item.issuerId] += equalShare;
        });
        remainingAmount = 0;
        break;
      }

      const amountNeeded = (nextLevel - levels[index].total) * currentGroup.length;
      if (!(amountNeeded > 0)) continue;

      if (remainingAmount >= amountNeeded) {
        currentGroup.forEach(item => {
          allocations[item.issuerId] += nextLevel - item.total;
          item.total = nextLevel;
        });
        remainingAmount -= amountNeeded;
      } else {
        const equalShare = remainingAmount / currentGroup.length;
        currentGroup.forEach(item => {
          allocations[item.issuerId] += equalShare;
        });
        remainingAmount = 0;
        break;
      }
    }

    const roundedAllocations = {};
    let assignedTotal = 0;
    issuerIds.forEach((issuerId, index) => {
      if (index === issuerIds.length - 1) {
        roundedAllocations[issuerId] = Number((totalAmount - assignedTotal).toFixed(2));
        return;
      }

      roundedAllocations[issuerId] = Number((allocations[issuerId] || 0).toFixed(2));
      assignedTotal += roundedAllocations[issuerId];
    });

    return roundedAllocations;
  },

  buildInvoiceClientScopedState: (state, clientSource, month = null, options = {}) => {
    const selectedMonth = month || Calculations.getSelectedMonth(state);
    const clientId = clientSource?.clientId || '';
    const clientName = clientSource?.clientName || '';
    const normalizedClientName = Calculations.normalizeText(clientName);
    const includeClientCosts = options.includeClientCosts !== false;
    const clientCostShareRatio = Math.max(0, Math.min(1, parseFloat(options.clientCostShareRatio) || 0));
    const baseScopedState = {
      ...state,
      selectedMonth,
      monthlySheets: (state?.monthlySheets || []).filter(sheet =>
        Calculations.sheetMatchesSelectedMonth(sheet, selectedMonth)
        && Calculations.normalizeText(sheet?.client) === normalizedClientName
      ),
      worksSheets: (state?.worksSheets || []).filter(sheet =>
        Calculations.sheetMatchesSelectedMonth(sheet, selectedMonth)
        && Calculations.normalizeText(sheet?.client) === normalizedClientName
      ),
      expenses: []
    };
    const baseMonthSettings = Calculations.getMonthSettings(state, selectedMonth) || {};

    const getPersonClientShare = (() => {
      const cache = new Map();
      return (personId) => {
        if (!personId) return 0;
        if (cache.has(personId)) return cache.get(personId);

        const person = (state?.persons || []).find(candidate => candidate.id === personId);
        if (!person) {
          cache.set(personId, 0);
          return 0;
        }

        const totalStats = Calculations.calculatePersonStats(person, { ...state, selectedMonth }, 0);
        const clientStats = Calculations.calculatePersonStats(person, baseScopedState, 0);
        let share = 0;
        if ((totalStats.totalSalary || 0) > 0 && (clientStats.totalSalary || 0) > 0) {
          share = clientStats.totalSalary / totalStats.totalSalary;
        } else if ((totalStats.totalHours || 0) > 0 && (clientStats.totalHours || 0) > 0) {
          share = clientStats.totalHours / totalStats.totalHours;
        }

        const normalizedShare = Math.max(0, Math.min(1, share));
        cache.set(personId, normalizedShare);
        return normalizedShare;
      };
    })();

    const getPersonHasOwnMonthActivity = (() => {
      const cache = new Map();
      return (personId) => {
        if (!personId) return false;
        if (cache.has(personId)) return cache.get(personId);

        const person = (state?.persons || []).find(candidate => candidate.id === personId);
        if (!person) {
          cache.set(personId, false);
          return false;
        }

        const totalStats = Calculations.calculatePersonStats(person, { ...state, selectedMonth }, 0);
        const hasActivity = (totalStats.totalSalary || 0) > 0 || (totalStats.totalHours || 0) > 0;
        cache.set(personId, hasActivity);
        return hasActivity;
      };
    })();

    const getAdvanceClientShare = (personId) => {
      const directShare = getPersonClientShare(personId);
      if (directShare > 0) return directShare;

      const person = (state?.persons || []).find(candidate => candidate.id === personId);
      if (person?.employerId && !getPersonHasOwnMonthActivity(personId)) {
        return getPersonClientShare(person.employerId);
      }

      return 0;
    };

    const scopedExpenses = [];
    (state?.expenses || []).forEach(expense => {
      if (!Calculations.isDateInMonth(expense?.date, selectedMonth)) return;

      if (expense?.type === 'COST') {
        if (!includeClientCosts || !(clientCostShareRatio > 0)) return;
        const amount = Number(((parseFloat(expense.amount) || 0) * clientCostShareRatio).toFixed(2));
        if (!(Math.abs(amount) >= 0.005)) return;

        scopedExpenses.push({
          ...expense,
          id: `${expense.id}::client-cost::${clientId || normalizedClientName}`,
          amount
        });
        return;
      }

      if (expense?.type === 'BONUS' || expense?.type === 'DIETA') {
        const personShare = getAdvanceClientShare(expense?.advanceForId);
        if (!(personShare > 0)) return;

        const amount = Number((Calculations.getExpenseEffectiveAmount(expense, state, selectedMonth) * personShare).toFixed(2));
        if (!(Math.abs(amount) >= 0.005)) return;

        scopedExpenses.push({
          ...expense,
          id: `${expense.id}::client-${expense.type.toLowerCase()}::${clientId || normalizedClientName}`,
          amount
        });
        return;
      }

      if (expense?.type !== 'ADVANCE') return;

      if (expense?.paidById && expense.paidById.startsWith('client_')) {
        if (!!clientId && expense.paidById === `client_${clientId}`) {
          scopedExpenses.push({ ...expense });
        }
        return;
      }

      const personShare = getAdvanceClientShare(expense?.advanceForId);
      if (!(personShare > 0)) return;

      const amount = Number(((parseFloat(expense.amount) || 0) * personShare).toFixed(2));
      if (!(Math.abs(amount) >= 0.005)) return;

      scopedExpenses.push({
        ...expense,
        id: `${expense.id}::client-scope::${clientId || normalizedClientName}`,
        amount
      });
    });

    const scopedPersonContractCharges = {};
    (state?.persons || []).forEach(person => {
      if (person?.type !== 'EMPLOYEE' && person?.type !== 'WORKING_PARTNER') return;

      const personShare = getPersonClientShare(person.id);
      const charges = Calculations.getPersonContractCharges(person, state, selectedMonth);
      if (!(charges.taxAmount > 0 || charges.zusAmount > 0)) return;

      scopedPersonContractCharges[person.id] = {
        contractTaxAmount: Number((charges.taxAmount * personShare).toFixed(2)),
        contractZusAmount: Number((charges.zusAmount * personShare).toFixed(2))
      };
    });

    return {
      ...baseScopedState,
      monthSettings: {
        ...(state?.monthSettings || {}),
        [selectedMonth]: {
          ...baseMonthSettings,
          personContractCharges: {
            ...(baseMonthSettings.personContractCharges || {}),
            ...scopedPersonContractCharges
          }
        }
      },
      expenses: scopedExpenses
    };
  },

  getInvoiceSettlementAmounts: (state, clientSource, month = null, options = {}) => {
    const filteredState = Calculations.buildInvoiceClientScopedState(state, clientSource, month, options);
    const settlement = Calculations.generateSettlement(filteredState, { includeInvoiceTaxEqualization: false, includeInvoiceReconciliation: false });
    const result = {};

    [...(settlement.partners || []), ...(settlement.workingPartners || []), ...((settlement.separateCompanies || []))].forEach(entry => {
      const personId = entry?.person?.id;
      if (!personId) return;

      const grossWithEmployees = parseFloat(entry.grossWithEmployeeSalaries);
      const toPayout = parseFloat(entry.toPayout);
      const gross = Math.max(0, Number.isFinite(toPayout) ? toPayout : 0);
      const grossWithEmployeesSafe = Math.max(0, Number.isFinite(grossWithEmployees) ? grossWithEmployees : gross);
      result[personId] = {
        gross,
        grossWithEmployees: grossWithEmployeesSafe,
        invoiceAmount: Math.abs(grossWithEmployeesSafe - gross) >= 0.005 ? grossWithEmployeesSafe : gross,
        invoiceAmountLabel: Math.abs(grossWithEmployeesSafe - gross) >= 0.005 ? 'Przychód (Brutto) z Pensjami' : 'Przychód (Brutto)'
      };
    });

    return result;
  },

  calculateClientInvoiceSources: (state, month = null) => {
    const selectedMonth = month || Calculations.getSelectedMonth(state);
    const eligibleIssuers = Calculations.getInvoiceEligibleIssuers(state, selectedMonth);
    const eligibleIssuerIds = new Set(eligibleIssuers.map(person => person.id));
    const buckets = new Map();

    const ensureBucket = (clientName) => {
      const identity = Calculations.getInvoiceClientIdentity(state, clientName);
      if (!buckets.has(identity.clientId)) {
        buckets.set(identity.clientId, {
          ...identity,
          hourlyRevenue: 0,
          worksRevenue: 0,
          clientAdvances: 0,
          totalRevenue: 0,
          issuerWeights: {},
          issuerIds: []
        });
      }
      return buckets.get(identity.clientId);
    };

    const addWeight = (bucket, issuerId, amount) => {
      if (!issuerId || !Number.isFinite(amount) || Math.abs(amount) < 0.005) return;
      bucket.issuerWeights[issuerId] = (bucket.issuerWeights[issuerId] || 0) + amount;
    };

    (state?.monthlySheets || []).forEach(sheet => {
      if (!Calculations.sheetMatchesSelectedMonth(sheet, selectedMonth)) return;
      const clientRate = Calculations.getSheetClientRate(sheet, state);
      const bucket = ensureBucket(sheet.client);

      Object.entries(sheet.days || {}).forEach(([dayKey, day]) => {
        Object.entries(day?.hours || {}).forEach(([personId, value]) => {
          if (!Calculations.isPersonIncludedInMonthlySheet(personId, sheet)) return;
          if (Calculations.isPersonInactiveInMonthlySheetDay(sheet, personId, parseInt(dayKey, 10))) return;

          const hours = parseFloat(value);
          if (!Number.isFinite(hours) || hours <= 0) return;

          const amount = hours * clientRate;
          const issuerId = Calculations.getInvoiceIssuerForPerson(personId, state, eligibleIssuerIds);
          bucket.hourlyRevenue += amount;
          bucket.totalRevenue += amount;
          addWeight(bucket, issuerId, amount);
        });
      });
    });

    (state?.worksSheets || []).forEach(sheet => {
      if (!Calculations.sheetMatchesSelectedMonth(sheet, selectedMonth)) return;
      const bucket = ensureBucket(sheet.client);

      (sheet.entries || []).forEach(entry => {
        const metrics = Calculations.getWorksEntryMetrics(entry, sheet, state);
        if (!(metrics.revenue > 0)) return;

        bucket.worksRevenue += metrics.revenue;
        bucket.totalRevenue += metrics.revenue;

        if (metrics.isRoboczogodziny) {
          Object.entries(entry?.hours || {}).forEach(([personId, value]) => {
            const hours = parseFloat(value);
            if (!Number.isFinite(hours) || hours <= 0) return;
            const issuerId = Calculations.getInvoiceIssuerForPerson(personId, state, eligibleIssuerIds);
            addWeight(bucket, issuerId, hours * metrics.clientRate);
          });
          return;
        }

        const involvedIssuerIds = new Set();
        [...(sheet.activePersons || []), ...Object.keys(entry?.hours || {})].forEach(personId => {
          const issuerId = Calculations.getInvoiceIssuerForPerson(personId, state, eligibleIssuerIds);
          if (issuerId) involvedIssuerIds.add(issuerId);
        });

        const fallbackIssuerIds = Object.keys(bucket.issuerWeights).filter(issuerId => (bucket.issuerWeights[issuerId] || 0) > 0);
        const targetIssuerIds = involvedIssuerIds.size > 0
          ? [...involvedIssuerIds]
          : (fallbackIssuerIds.length > 0 ? fallbackIssuerIds : eligibleIssuers.map(person => person.id));

        if (!targetIssuerIds.length) return;
        const splitAmount = metrics.revenue / targetIssuerIds.length;
        targetIssuerIds.forEach(issuerId => addWeight(bucket, issuerId, splitAmount));
      });
    });

    (state?.expenses || []).forEach(expense => {
      if (!Calculations.isDateInMonth(expense?.date, selectedMonth)) return;
      if (expense?.type !== 'ADVANCE') return;
      if (!expense?.paidById || !expense.paidById.startsWith('client_')) return;

      const clientId = expense.paidById.substring(7);
      const client = (state?.clients || []).find(item => item.id === clientId);
      const bucket = ensureBucket(client?.name || 'Nieznany klient');
      bucket.clientAdvances += parseFloat(expense.amount) || 0;
    });

    const clientOrder = new Map((state?.clients || []).map((client, index) => [client.id, index]));
    return [...buckets.values()]
      .map(bucket => ({
        ...bucket,
        issuerIds: Object.keys(bucket.issuerWeights).filter(issuerId => (bucket.issuerWeights[issuerId] || 0) > 0)
      }))
      .filter(bucket => bucket.totalRevenue > 0)
      .sort((a, b) => (clientOrder.get(a.clientId) ?? Number.MAX_SAFE_INTEGER) - (clientOrder.get(b.clientId) ?? Number.MAX_SAFE_INTEGER) || a.clientName.localeCompare(b.clientName, 'pl-PL'));
  },

  calculateInvoiceClientCostShares: (sources, invoiceSettings = {}) => {
    const enabledSources = (sources || []).filter(source => {
      const config = invoiceSettings.clients?.[source.clientId] || invoiceSettings.clients?.[`name:${Calculations.normalizeText(source.clientName)}`] || {};
      return config.includeClientCosts !== false;
    });
    const totalRevenue = enabledSources.reduce((sum, source) => sum + (parseFloat(source.totalRevenue) || 0), 0);
    const shares = {};

    enabledSources.forEach(source => {
      shares[source.clientId] = totalRevenue > 0
        ? (parseFloat(source.totalRevenue) || 0) / totalRevenue
        : (enabledSources.length > 0 ? 1 / enabledSources.length : 0);
    });

    return shares;
  },

  getInvoiceYearToDateBaseTotals: (state, month = null, options = {}) => {
    const selectedMonth = month || Calculations.getSelectedMonth(state);
    if (!/^\d{4}-\d{2}$/.test(selectedMonth)) return {};

    const selectedYear = selectedMonth.slice(0, 4);
    const fullState = options.fullState || (typeof Store !== 'undefined' && Store.getExportData ? Store.getExportData() : null);
    const knownMonths = typeof Store !== 'undefined' && Store.getKnownMonths
      ? Store.getKnownMonths()
      : Object.keys(fullState?.months || {});
    const previousMonthKeys = knownMonths
      .filter(monthKey => /^\d{4}-\d{2}$/.test(monthKey) && monthKey.startsWith(`${selectedYear}-`) && monthKey < selectedMonth)
      .sort((a, b) => a.localeCompare(b, 'pl-PL'));

    const totals = {};
    previousMonthKeys.forEach(monthKey => {
      const monthState = typeof Store !== 'undefined' && Store.getStateForMonth
        ? Store.getStateForMonth(monthKey)
        : (() => {
            const monthRecord = fullState?.months?.[monthKey] || { monthlySheets: [], worksSheets: [], expenses: [], monthSettings: {} };
            const monthSettings = monthRecord.monthSettings || {};
            const common = monthSettings.commonSnapshot || fullState?.common || {};

            return {
              selectedMonth: monthKey,
              isArchived: monthSettings.isArchived === true,
              hasCommonSnapshot: !!monthSettings.commonSnapshot,
              persons: common.persons || [],
              clients: common.clients || [],
              worksCatalog: common.worksCatalog || [],
              config: common.config || {},
              settings: {},
              monthlySheets: monthRecord.monthlySheets || [],
              worksSheets: monthRecord.worksSheets || [],
              expenses: monthRecord.expenses || [],
              monthSettings: { [monthKey]: monthSettings }
            };
          })();
      const invoiceSummary = Calculations.calculateInvoices(monthState, monthKey, {
        includeYearToDate: false,
        yearToDateBaseTotals: totals,
        fullState
      });

      (invoiceSummary?.issuerSummaries || []).forEach(summary => {
        if (summary?.issuerType !== 'PARTNER' && summary?.issuerType !== 'WORKING_PARTNER') return;

        const totalAmount = parseFloat(summary.totalAmount) || 0;
        if (Math.abs(totalAmount) < 0.005) return;
        totals[summary.issuerId] = (totals[summary.issuerId] || 0) + Number(totalAmount.toFixed(2));
      });
    });

    return totals;
  },

  calculateInvoices: (state, month = null, options = {}) => {
    const selectedMonth = month || Calculations.getSelectedMonth(state);
    const scopedState = Calculations._wrapState(state, selectedMonth);
    state = scopedState; // Przekierowanie na opakowany stan dla reszty funkcji
    const includeYearToDate = options.includeYearToDate !== false;
    const monthSettings = Calculations.getMonthSettings(state, selectedMonth);
    const settlementConfig = Calculations.getSettlementConfig(state, selectedMonth);
    const invoiceSettings = monthSettings?.invoices || { issueDate: '', emailIntro: '', issued: false, issuedAt: '', issuedSnapshot: null, clients: {}, extraInvoices: [] };

    if (invoiceSettings.issued === true && invoiceSettings.issuedSnapshot && options.ignoreIssuedSnapshot !== true) {
      const issuedSnapshot = JSON.parse(JSON.stringify(invoiceSettings.issuedSnapshot || {}));
      const totalRevenue = parseFloat(issuedSnapshot.totalRevenue) || 0;
      const totalInvoices = parseFloat(issuedSnapshot.totalInvoices) || 0;

      return {
        month: selectedMonth,
        issueDate: (issuedSnapshot.issueDate || invoiceSettings.issueDate || Calculations.getMonthLastDate(selectedMonth)).toString().trim(),
        issuers: Array.isArray(issuedSnapshot.issuers) ? issuedSnapshot.issuers : [],
        clientInvoices: Array.isArray(issuedSnapshot.clientInvoices) ? issuedSnapshot.clientInvoices : [],
        extraInvoices: Array.isArray(issuedSnapshot.extraInvoices) ? issuedSnapshot.extraInvoices : [],
        issuerSummaries: Array.isArray(issuedSnapshot.issuerSummaries) ? issuedSnapshot.issuerSummaries : [],
        totalRevenue,
        totalInvoices,
        difference: Number.isFinite(parseFloat(issuedSnapshot.difference)) ? parseFloat(issuedSnapshot.difference) : (totalInvoices - totalRevenue),
        emailText: (issuedSnapshot.emailText || '').toString(),
        isIssuedSnapshot: true,
        issuedAt: (invoiceSettings.issuedAt || '').toString().trim()
      };
    }

    const issuers = Calculations.getInvoiceEligibleIssuers(state, selectedMonth);
    const issuerById = Object.fromEntries(issuers.map(person => [person.id, person]));
    const sources = Calculations.calculateClientInvoiceSources(state, selectedMonth);
    const clientCostShares = Calculations.calculateInvoiceClientCostShares(sources, invoiceSettings);
    const issueDate = invoiceSettings.issueDate || Calculations.getMonthLastDate(selectedMonth);
    const issuerTotals = {};
    const issuerYearToDateBaseTotals = Object.entries(options.yearToDateBaseTotals || {}).reduce((acc, [issuerId, amount]) => {
      const normalizedAmount = parseFloat(amount);
      if (Number.isFinite(normalizedAmount) && Math.abs(normalizedAmount) >= 0.005) {
        acc[issuerId] = Number(normalizedAmount.toFixed(2));
      }
      return acc;
    }, {});
    const hasExplicitYearToDateBaseTotals = Object.keys(issuerYearToDateBaseTotals).length > 0;
    const extraInvoices = (invoiceSettings.extraInvoices || [])
      .map(invoice => {
        const issuer = issuerById[invoice.issuerId];
        if (!issuer) return null;

        const client = invoice.clientId
          ? (state?.clients || []).find(item => item.id === invoice.clientId)
          : null;
        const clientName = (client?.name || invoice.clientName || 'Nieznany klient').toString().trim();
        const amount = parseFloat(invoice.amount) || 0;
        if (!(amount > 0) || !clientName) return null;

        return {
          ...invoice,
          clientName,
          issuerName: Calculations.getPersonDisplayName(issuer),
          issuerType: issuer.type,
          amount
        };
      })
      .filter(Boolean);

    if (includeYearToDate && /^\d{4}-\d{2}$/.test(selectedMonth) && !hasExplicitYearToDateBaseTotals) {
      Object.assign(issuerYearToDateBaseTotals, Calculations.getInvoiceYearToDateBaseTotals(state, selectedMonth, options));
    }

    const clientInvoices = sources.map(source => {
      const config = invoiceSettings.clients?.[source.clientId] || invoiceSettings.clients?.[`name:${Calculations.normalizeText(source.clientName)}`] || {};
      const configuredIds = Array.isArray(config.issuerIds) ? config.issuerIds.filter(id => issuerById[id]) : [];
      const selectedIssuerIds = configuredIds.length > 0
        ? configuredIds
        : (source.issuerIds.length > 0 ? source.issuerIds.filter(id => issuerById[id]) : issuers.map(person => person.id));
      const mode = config.mode || 'SETTLEMENT_REVENUE';
      const includeClientCosts = mode === 'SETTLEMENT_REVENUE' && config.includeClientCosts !== false;
      const settlementAmountsByIssuer = Calculations.getInvoiceSettlementAmounts(state, source, selectedMonth, {
        includeClientCosts,
        clientCostShareRatio: clientCostShares[source.clientId] || 0
      });
      const deductClientAdvances = config.deductClientAdvances !== false;
      const netRevenue = source.totalRevenue - (deductClientAdvances ? (source.clientAdvances || 0) : 0);
      const separateCompanyWithEmployees = config.separateCompanyWithEmployees || {};
      const allocationByIssuer = {};
      let remainingRevenue = netRevenue;
      const fixedIssuerIds = new Set();

      selectedIssuerIds.forEach(issuerId => {
        const issuer = issuerById[issuerId];
        const settlementAmounts = settlementAmountsByIssuer[issuerId] || { gross: 0, grossWithEmployees: 0, invoiceAmount: 0 };
        if (issuer?.type === 'SEPARATE_COMPANY' && separateCompanyWithEmployees[issuerId] === true) {
          allocationByIssuer[issuerId] = settlementAmounts.invoiceAmount || 0;
          remainingRevenue -= allocationByIssuer[issuerId];
          fixedIssuerIds.add(issuerId);
        }
      });

      const remainingIssuerIds = selectedIssuerIds.filter(issuerId => !fixedIssuerIds.has(issuerId));
      const normalizedPercentageAllocations = Calculations.normalizeInvoicePercentageAllocations(remainingIssuerIds, config.percentageAllocations || {});

      if (mode === 'MANUAL') {
        remainingIssuerIds.forEach(issuerId => {
          const amount = parseFloat(config?.manualAmounts?.[issuerId]);
          allocationByIssuer[issuerId] = Number.isFinite(amount) ? amount : 0;
        });

        const manualTotal = remainingIssuerIds.reduce((sum, issuerId) => sum + (allocationByIssuer[issuerId] || 0), 0);
        const difference = remainingRevenue - manualTotal;
        if (remainingIssuerIds.length > 0 && Math.abs(difference) >= 0.005) {
          const firstIssuerId = remainingIssuerIds[0];
          allocationByIssuer[firstIssuerId] = (allocationByIssuer[firstIssuerId] || 0) + difference;
        }
      } else if (mode === 'EQUAL_SPLIT') {
        const equalSplitAllocations = Calculations.buildEqualSplitAllocations(remainingRevenue, remainingIssuerIds, {
          randomize: config.randomizeEqualSplitInvoices === true,
          varianceAmount: config.equalSplitVarianceAmount,
          seed: `${selectedMonth}|${source.clientId}|${remainingRevenue.toFixed(2)}`
        });
        remainingIssuerIds.forEach(issuerId => {
          allocationByIssuer[issuerId] = equalSplitAllocations[issuerId] || 0;
        });
      } else if (mode === 'PERCENTAGE_SPLIT') {
        const percentageSplitAllocations = Calculations.buildPercentageSplitAllocations(remainingRevenue, remainingIssuerIds, normalizedPercentageAllocations);
        remainingIssuerIds.forEach(issuerId => {
          allocationByIssuer[issuerId] = percentageSplitAllocations[issuerId] || 0;
        });
      } else if (mode === 'BALANCE_INVOICE_SUMS') {
        const balancingTotals = Object.fromEntries(remainingIssuerIds.map(issuerId => [
          issuerId,
          (issuerYearToDateBaseTotals[issuerId] || 0) + (issuerTotals[issuerId] || 0)
        ]));
        const balancedAllocations = Calculations.buildBalancedInvoiceSumAllocations(remainingRevenue, remainingIssuerIds, balancingTotals);
        remainingIssuerIds.forEach(issuerId => {
          allocationByIssuer[issuerId] = balancedAllocations[issuerId] || 0;
        });
      } else if (mode === 'SETTLEMENT_REVENUE') {
        let totalSettlementWeight = 0;

        remainingIssuerIds.forEach(issuerId => {
          const settlementAmounts = settlementAmountsByIssuer[issuerId] || { gross: 0, grossWithEmployees: 0, invoiceAmount: 0 };
          totalSettlementWeight += settlementAmounts.invoiceAmount || 0;
        });

        if (remainingIssuerIds.length > 0) {
          if (totalSettlementWeight > 0) {
            remainingIssuerIds.forEach(issuerId => {
              const settlementAmounts = settlementAmountsByIssuer[issuerId] || { gross: 0, grossWithEmployees: 0, invoiceAmount: 0 };
              allocationByIssuer[issuerId] = remainingRevenue * ((settlementAmounts.invoiceAmount || 0) / totalSettlementWeight);
            });
          } else {
            const equalAmount = remainingRevenue / remainingIssuerIds.length;
            remainingIssuerIds.forEach(issuerId => {
              allocationByIssuer[issuerId] = equalAmount;
            });
          }
        }
      } else {
        const selectedWeight = remainingIssuerIds.reduce((sum, issuerId) => sum + (source.issuerWeights[issuerId] || 0), 0);
        if (selectedWeight > 0) {
          remainingIssuerIds.forEach(issuerId => {
              allocationByIssuer[issuerId] = remainingRevenue * ((source.issuerWeights[issuerId] || 0) / selectedWeight);
          });
        } else {
          const equalAmount = remainingIssuerIds.length > 0 ? remainingRevenue / remainingIssuerIds.length : 0;
          remainingIssuerIds.forEach(issuerId => {
            allocationByIssuer[issuerId] = equalAmount;
          });
        }
      }

      const allocations = selectedIssuerIds.map(issuerId => {
        const issuer = issuerById[issuerId];
        const amount = allocationByIssuer[issuerId] || 0;
        issuerTotals[issuerId] = (issuerTotals[issuerId] || 0) + amount;
        return {
          issuerId,
          issuerName: Calculations.getPersonDisplayName(issuer),
          issuerType: issuer?.type || '',
          amount
        };
      }).filter(allocation => Math.abs(allocation.amount) >= 0.005);

      return {
        ...source,
        mode,
        deductClientAdvances,
        includeClientCosts,
        clientCostShareRatio: clientCostShares[source.clientId] || 0,
        randomizeEqualSplitInvoices: config.randomizeEqualSplitInvoices === true,
        equalSplitVarianceAmount: Math.max(0, parseFloat(config.equalSplitVarianceAmount) || 10),
        percentageAllocations: Calculations.normalizeInvoicePercentageAllocations(selectedIssuerIds, config.percentageAllocations || {}),
        separateCompanyWithEmployees,
        settlementAmountsByIssuer,
        modeLabel: Calculations.getInvoiceModeLabel(mode),
        notes: (config.notes || '').toString().trim(),
        issueDate,
        netRevenue,
        allocations,
        allocatedTotal: allocations.reduce((sum, allocation) => sum + allocation.amount, 0),
        difference: allocations.reduce((sum, allocation) => sum + allocation.amount, 0) - netRevenue
      };
    });

    extraInvoices.forEach(invoice => {
      issuerTotals[invoice.issuerId] = (issuerTotals[invoice.issuerId] || 0) + invoice.amount;
    });

    const issuerSummaries = issuers
      .map(issuer => ({
        issuerId: issuer.id,
        issuerName: Calculations.getPersonDisplayName(issuer),
        issuerType: issuer.type,
        totalAmount: issuerTotals[issuer.id] || 0,
        settlementAmount: clientInvoices.reduce((sum, clientInvoice) => sum + (clientInvoice.allocations || [])
          .filter(allocation => allocation.issuerId === issuer.id)
          .reduce((allocationSum, allocation) => allocationSum + allocation.amount, 0), 0),
        extraInvoicesAmount: extraInvoices
          .filter(invoice => invoice.issuerId === issuer.id)
          .reduce((sum, invoice) => sum + invoice.amount, 0),
        taxRate: settlementConfig.taxRate,
        taxAmount: Calculations.calculateAccountingTax(issuerTotals[issuer.id] || 0, settlementConfig.taxRate),
        yearToDateTotal: issuer.type === 'PARTNER' || issuer.type === 'WORKING_PARTNER'
          ? ((issuerYearToDateBaseTotals[issuer.id] || 0) + (issuerTotals[issuer.id] || 0))
          : null
      }))
      .filter(summary => Math.abs(summary.totalAmount) >= 0.005);

    const totalRevenue = clientInvoices.reduce((sum, clientInvoice) => sum + clientInvoice.netRevenue, 0);
    const totalInvoices = issuerSummaries.reduce((sum, issuer) => sum + issuer.totalAmount, 0);

    const emailLines = [
      (invoiceSettings.emailIntro || `Witam, potrzebujemy następujących faktur za ${Calculations.formatMonthLabel(selectedMonth)}:`).trim(),
      ''
    ];
    let emailTotalInvoices = 0;

    clientInvoices.forEach((clientInvoice, index) => {
      const emailAllocations = (clientInvoice.allocations || []).filter(allocation => allocation.issuerType !== 'SEPARATE_COMPANY');
      const emailClientTotal = emailAllocations.reduce((sum, allocation) => sum + (allocation.amount || 0), 0);
      if (emailAllocations.length === 0) return;

      emailTotalInvoices += emailClientTotal;

      emailLines.push(`${index + 1}. ${clientInvoice.fullCompanyName || clientInvoice.clientName}`);
      emailLines.push('');
      if (clientInvoice.address) emailLines.push(`Adres: ${clientInvoice.address}`);
      if (clientInvoice.nip) emailLines.push(`NIP: ${clientInvoice.nip}`);
      if (clientInvoice.krs) emailLines.push(`KRS: ${clientInvoice.krs}`);
      if (clientInvoice.regon) emailLines.push(`REGON: ${clientInvoice.regon}`);
      emailLines.push('');
      emailLines.push(`Data wystawienia: ${clientInvoice.issueDate}`);
      emailLines.push('');
      emailAllocations.forEach((allocation, allocationIndex) => {
        emailLines.push(`${allocationIndex + 1}. ${allocation.issuerName} — ${Calculations.formatInvoiceCurrency(allocation.amount)}`);
      });
      emailLines.push('');
      emailLines.push(`Razem dla klienta: ${Calculations.formatInvoiceCurrency(emailClientTotal)}`);
      if (clientInvoice.notes) emailLines.push(`Uwagi: ${clientInvoice.notes}`);
      emailLines.push('');
      if (index < clientInvoices.length - 1) {
        emailLines.push('---------------------------------------------------------------------');
        emailLines.push('');
      }
    });

    emailLines.push(`Łączna suma faktur: ${Calculations.formatInvoiceCurrency(emailTotalInvoices)}`);

    return {
      month: selectedMonth,
      issueDate,
      issuers,
      clientInvoices,
      extraInvoices,
      issuerSummaries,
      totalRevenue,
      totalInvoices,
      difference: totalInvoices - totalRevenue,
      emailText: emailLines.join('\n').trim(),
      isIssuedSnapshot: false,
      issuedAt: (invoiceSettings.issuedAt || '').toString().trim()
    };
  },

  getSettlementInvoiceParticipants: (settlement) => {
    const entries = [
      ...(settlement?.partners || []),
      ...((settlement?.workingPartners || []).filter(entry => !entry?.person?.employerId)),
      ...((settlement?.separateCompanies || []))
    ];

    return entries
      .filter(entry => !!entry?.person?.id)
      .map(entry => {
        const gross = Number.isFinite(parseFloat(entry.toPayout))
          ? parseFloat(entry.toPayout)
          : 0;
        const grossWithEmployees = Number.isFinite(parseFloat(entry.grossWithEmployeeSalaries))
          ? parseFloat(entry.grossWithEmployeeSalaries)
          : gross;
        const desiredInvoiceAmount = Math.abs(grossWithEmployees - gross) >= 0.005 ? grossWithEmployees : gross;

        return {
          personId: entry.person.id,
          personName: Calculations.getPersonDisplayName(entry.person),
          personType: entry.person.type,
          desiredInvoiceAmount,
          desiredInvoiceLabel: Math.abs(grossWithEmployees - gross) >= 0.005 ? 'Przychód (Brutto) z Pensjami' : 'Przychód (Brutto)',
          settlementEntry: entry
        };
      });
  },

  buildInvoiceCoverageAllocations: (participants, actualIssuedAmountByIssuer = {}) => {
    const participantById = Object.fromEntries((participants || []).map(participant => [participant.personId, participant]));
    const issuerRemaining = Object.fromEntries((participants || []).map(participant => [participant.personId, Math.max(0, parseFloat(actualIssuedAmountByIssuer?.[participant.personId]) || 0)]));
    const beneficiaryRemaining = Object.fromEntries((participants || []).map(participant => [participant.personId, Math.max(0, parseFloat(participant.desiredInvoiceAmount) || 0)]));
    const allocations = [];

    const allocate = (issuerId, beneficiaryId, rawAmount) => {
      const amount = Math.max(0, Math.min(parseFloat(rawAmount) || 0, issuerRemaining[issuerId] || 0, beneficiaryRemaining[beneficiaryId] || 0));
      if (!(amount >= 0.005)) return;

      const issuer = participantById[issuerId];
      const beneficiary = participantById[beneficiaryId];
      allocations.push({
        issuerId,
        issuerName: issuer?.personName || 'Nieznany wystawca',
        issuerType: issuer?.personType || '',
        beneficiaryId,
        beneficiaryName: beneficiary?.personName || 'Nieznany odbiorca',
        beneficiaryType: beneficiary?.personType || '',
        amount
      });

      issuerRemaining[issuerId] -= amount;
      beneficiaryRemaining[beneficiaryId] -= amount;
    };

    (participants || []).forEach(participant => {
      allocate(participant.personId, participant.personId, Math.min(issuerRemaining[participant.personId] || 0, beneficiaryRemaining[participant.personId] || 0));
    });

    (participants || []).forEach(issuer => {
      (participants || []).forEach(beneficiary => {
        if (issuer.personId === beneficiary.personId) return;
        allocate(issuer.personId, beneficiary.personId, Math.min(issuerRemaining[issuer.personId] || 0, beneficiaryRemaining[beneficiary.personId] || 0));
      });
    });

    return {
      allocations,
      issuerRemaining,
      beneficiaryRemaining
    };
  },

  calculateInvoiceTaxEqualization: (state, settlement, month = null) => {
    const selectedMonth = month || Calculations.getSelectedMonth(state);
    const invoiceData = Calculations.calculateInvoices(state, selectedMonth);
    const issuerSummaryById = Object.fromEntries((invoiceData.issuerSummaries || []).map(summary => [summary.issuerId, summary]));
    const participants = Calculations.getSettlementInvoiceParticipants(settlement);
    const personById = Object.fromEntries((state?.persons || []).map(person => [person.id, person]));
    const participantById = Object.fromEntries(participants.map(participant => [participant.personId, participant]));
    const participantIdSet = new Set(participants.map(participant => participant.personId));
    const settlementEntryByPersonId = Object.fromEntries(participants.map(participant => [participant.personId, participant.settlementEntry || {}]));
    const actualIssuedAmountByIssuer = Object.fromEntries(participants.map(participant => [participant.personId, parseFloat(issuerSummaryById[participant.personId]?.totalAmount) || 0]));
    const invoiceTaxByIssuer = Object.fromEntries(participants.map(participant => [participant.personId, parseFloat(issuerSummaryById[participant.personId]?.taxAmount) || 0]));
    const issuerTaxRateByIssuer = Object.fromEntries(participants.map(participant => {
      const actualIssuedAmount = actualIssuedAmountByIssuer[participant.personId] || 0;
      const invoiceTaxAmount = invoiceTaxByIssuer[participant.personId] || 0;
      return [participant.personId, actualIssuedAmount > 0 ? invoiceTaxAmount / actualIssuedAmount : 0];
    }));
    const { allocations } = Calculations.buildInvoiceCoverageAllocations(participants, actualIssuedAmountByIssuer);
    const ownTaxByPerson = Object.fromEntries(participants.map(participant => [participant.personId, parseFloat(settlementEntryByPersonId[participant.personId]?.ownTaxAmount) || 0]));
    const sharedCompanyTaxByPerson = Object.fromEntries(participants.map(participant => [participant.personId, parseFloat(settlementEntryByPersonId[participant.personId]?.sharedCompanyTaxAmount) || 0]));
    const resolveRefundSettlementIssuerId = (personId = '') => {
      const normalizedPersonId = (personId || '').toString().trim();
      if (!normalizedPersonId) return '';
      if (participantIdSet.has(normalizedPersonId)) return normalizedPersonId;

      const person = personById[normalizedPersonId];
      const employerId = (person?.employerId || '').toString().trim();
      return employerId && participantIdSet.has(employerId) ? employerId : '';
    };
    const employerTeamRefundTaxByPerson = {};
    const taxPaidForOthersByIssuer = {};

    (state?.expenses || [])
      .filter(expense => expense?.type === 'REFUND' && Calculations.isDateInMonth(expense?.date, selectedMonth) && expense?.paidById)
      .forEach(expense => {
        const rawPayerId = (expense?.paidById || '').toString().trim();
        const payerIssuerId = resolveRefundSettlementIssuerId(rawPayerId);
        if (!payerIssuerId) return;

        const payerUsesEmployerBridge = payerIssuerId !== rawPayerId;
        const recipientIds = Calculations.getRefundRecipientIds(expense, state, selectedMonth);
        recipientIds.forEach(recipientId => {
          const normalizedRecipientId = (recipientId || '').toString().trim();
          const recipientIssuerId = resolveRefundSettlementIssuerId(normalizedRecipientId);
          if (!recipientIssuerId || recipientIssuerId === payerIssuerId) return;

          const recipientUsesEmployerBridge = recipientIssuerId !== normalizedRecipientId;
          if (!payerUsesEmployerBridge && !recipientUsesEmployerBridge) return;

          const distributedAmount = Calculations.getRefundReceivedAmountForPerson(expense, normalizedRecipientId, state, selectedMonth);
          const taxRate = issuerTaxRateByIssuer[payerIssuerId] || 0;
          const taxAmount = distributedAmount * taxRate;
          if (!(taxAmount >= 0.005)) return;

          employerTeamRefundTaxByPerson[recipientIssuerId] = (employerTeamRefundTaxByPerson[recipientIssuerId] || 0) + taxAmount;
          taxPaidForOthersByIssuer[payerIssuerId] = (taxPaidForOthersByIssuer[payerIssuerId] || 0) + taxAmount;
        });
      });

    const targetTaxByPerson = Object.fromEntries(participants.map(participant => [
      participant.personId,
      (parseFloat(settlementEntryByPersonId[participant.personId]?.taxAmount) || 0) + (employerTeamRefundTaxByPerson[participant.personId] || 0)
    ]));

    const separateCompanyReimbursements = allocations
      .filter(allocation => allocation.issuerId !== allocation.beneficiaryId && allocation.amount >= 0.005 && allocation.beneficiaryType === 'SEPARATE_COMPANY' && allocation.issuerType !== 'SEPARATE_COMPANY')
      .map(allocation => {
        const issuerTaxRate = issuerTaxRateByIssuer[allocation.issuerId] || 0;
        const taxAmount = allocation.amount * issuerTaxRate;

        return {
          payerId: allocation.beneficiaryId,
          payerName: allocation.beneficiaryName,
          payerType: allocation.beneficiaryType,
          recipientId: allocation.issuerId,
          recipientName: allocation.issuerName,
          recipientType: allocation.issuerType,
          revenueAmount: allocation.amount,
          taxAmount,
          taxRate: issuerTaxRate,
          type: `Zwrot podatku za ${Calculations.formatInvoiceCurrency(allocation.amount)} przychodu osobnej firmy`,
          category: 'separate-company-tax'
        };
      })
      .filter(entry => entry.taxAmount >= 0.005);

    const separateCompanyReimbursementsReceivedByIssuer = {};
    separateCompanyReimbursements.forEach(entry => {
      separateCompanyReimbursementsReceivedByIssuer[entry.recipientId] = (separateCompanyReimbursementsReceivedByIssuer[entry.recipientId] || 0) + (entry.taxAmount || 0);
    });

    const actualTaxBurdenByPerson = Object.fromEntries(participants.map(participant => {
      if (participant.personType === 'SEPARATE_COMPANY') {
        return [participant.personId, 0];
      }

      const actualInvoiceTax = invoiceTaxByIssuer[participant.personId] || 0;
      const separateCompanyReimbursement = separateCompanyReimbursementsReceivedByIssuer[participant.personId] || 0;
      return [participant.personId, Math.max(0, actualInvoiceTax - separateCompanyReimbursement)];
    }));

    const partnerTaxPayers = participants
      .filter(participant => participant.personType !== 'SEPARATE_COMPANY')
      .map(participant => ({
        personId: participant.personId,
        personName: participant.personName,
        personType: participant.personType,
        remainingAmount: Math.max(0, (targetTaxByPerson[participant.personId] || 0) - (actualTaxBurdenByPerson[participant.personId] || 0))
      }))
      .filter(entry => entry.remainingAmount >= 0.005);

    const partnerTaxRecipients = participants
      .filter(participant => participant.personType !== 'SEPARATE_COMPANY')
      .map(participant => ({
        personId: participant.personId,
        personName: participant.personName,
        personType: participant.personType,
        remainingAmount: Math.max(0, (actualTaxBurdenByPerson[participant.personId] || 0) - (targetTaxByPerson[participant.personId] || 0))
      }))
      .filter(entry => entry.remainingAmount >= 0.005);

    const partnerTaxReimbursements = [];
    partnerTaxRecipients.forEach(recipient => {
      partnerTaxPayers.forEach(payer => {
        if (!(recipient.remainingAmount >= 0.005) || !(payer.remainingAmount >= 0.005)) return;
        if (recipient.personId === payer.personId) return;

        const amount = Math.min(recipient.remainingAmount, payer.remainingAmount);
        partnerTaxReimbursements.push({
          payerId: payer.personId,
          payerName: payer.personName,
          payerType: payer.personType,
          recipientId: recipient.personId,
          recipientName: recipient.personName,
          recipientType: recipient.personType,
          revenueAmount: 0,
          taxAmount: amount,
          taxRate: 0,
          type: 'Wyrównanie podatku wspólnego i własnego',
          category: 'partner-tax-equalization'
        });

        payer.remainingAmount -= amount;
        recipient.remainingAmount -= amount;
      });
    });

    const reimbursements = [...separateCompanyReimbursements, ...partnerTaxReimbursements];
    const totalInvoiceTax = participants.reduce((sum, participant) => sum + (invoiceTaxByIssuer[participant.personId] || 0), 0);
    const totalOwnTax = participants.reduce((sum, participant) => sum + (ownTaxByPerson[participant.personId] || 0), 0);
    const sharedCompanyTaxTotal = participants.reduce((sum, participant) => sum + (sharedCompanyTaxByPerson[participant.personId] || 0), 0);
    const partnerSharedTaxParticipants = participants.filter(participant => (sharedCompanyTaxByPerson[participant.personId] || 0) > 0);
    const taxBalanceByPerson = Object.fromEntries(participants.map(participant => [
      participant.personId,
      (actualTaxBurdenByPerson[participant.personId] || 0) - (targetTaxByPerson[participant.personId] || 0)
    ]));

    return {
      month: selectedMonth,
      invoiceData,
      participants,
      actualIssuedAmountByIssuer,
      issuerTaxRateByIssuer,
      totalInvoiceTax,
      totalOwnTax,
      sharedCompanyTaxTotal,
      sharedCompanyTaxPerPerson: partnerSharedTaxParticipants.length > 0 ? sharedCompanyTaxTotal / partnerSharedTaxParticipants.length : 0,
      participantsCount: participants.length,
      invoiceTaxByIssuer,
      beneficiaryTaxByPerson: ownTaxByPerson,
      ownTaxByPerson,
      sharedCompanyTaxByPerson,
      targetTaxByPerson,
      actualTaxBurdenByPerson,
      taxBalanceByPerson,
      separateCompanyReimbursementsReceivedByIssuer,
      employerTeamRefundTaxByPerson,
      taxPaidForOthersByIssuer,
      allocations,
      reimbursements
    };
  },

  calculateInvoiceReconciliation: (state, month = null) => {
    const selectedMonth = month || Calculations.getSelectedMonth(state);
    const scopedState = Calculations._wrapState(state, selectedMonth);
    state = scopedState;
    const settlement = Calculations.generateSettlement(state, { includeInvoiceTaxEqualization: false, includeInvoiceReconciliation: false });
    const invoiceData = Calculations.calculateInvoices(state, selectedMonth);
    const invoiceTaxEqualization = Calculations.calculateInvoiceTaxEqualization(state, settlement, selectedMonth);
    const personById = Object.fromEntries((state?.persons || []).map(person => [person.id, person]));
    const settlementParticipants = Calculations.getSettlementInvoiceParticipants(settlement);
    const issuerSummaryById = Object.fromEntries((invoiceData.issuerSummaries || []).map(summary => [summary.issuerId, summary]));
    const issuers = settlementParticipants
      .map(participant => ({
        issuerId: participant.personId,
        issuerName: participant.personName,
        issuerType: participant.personType,
        receivedAmount: parseFloat(issuerSummaryById[participant.personId]?.totalAmount) || 0,
        ownTaxAmount: invoiceTaxEqualization.beneficiaryTaxByPerson?.[participant.personId] || 0,
        sharedCompanyTaxAmount: invoiceTaxEqualization.sharedCompanyTaxByPerson?.[participant.personId] || 0,
        targetTaxAmount: invoiceTaxEqualization.targetTaxByPerson?.[participant.personId] || 0,
        actualInvoiceTaxAmount: invoiceTaxEqualization.invoiceTaxByIssuer?.[participant.personId] || 0,
        actualTaxBurdenAmount: invoiceTaxEqualization.actualTaxBurdenByPerson?.[participant.personId] || 0,
        ownZusAmount: parseFloat(participant.settlementEntry?.zusAmount) || 0,
        companyTaxAmount: invoiceTaxEqualization.taxPaidForOthersByIssuer?.[participant.personId] || 0,
        targetAmount: participant.settlementEntry?.toPayout || 0,
        revenueTargetAmount: participant.desiredInvoiceAmount || 0,
        grossTargetAmount: participant.settlementEntry?.toPayout || 0,
        postTaxTargetAmount: participant.settlementEntry?.toPayout || 0,
        desiredInvoiceAmount: participant.desiredInvoiceAmount || 0,
        currentBalance: parseFloat(issuerSummaryById[participant.personId]?.totalAmount) || 0,
        revenueBalance: parseFloat(issuerSummaryById[participant.personId]?.totalAmount) || 0,
        salaryPayments: [],
        officePayments: [],
        transferPayments: [],
        incomingTransferPayments: []
      }))
      .filter(entry => Math.abs(entry.receivedAmount) >= 0.005 || Math.abs(entry.targetAmount) >= 0.005 || Math.abs(entry.ownTaxAmount) >= 0.005 || Math.abs(entry.sharedCompanyTaxAmount) >= 0.005 || Math.abs(entry.targetTaxAmount) >= 0.005 || Math.abs(entry.companyTaxAmount) >= 0.005);
    const issuerMap = Object.fromEntries(issuers.map(entry => [entry.issuerId, entry]));

    [...(settlement.partners || []), ...(settlement.workingPartners || []), ...((settlement.separateCompanies || []))].forEach(entry => {
      const issuerEntry = issuerMap[entry.person?.id];
      if (!issuerEntry) return;
      issuerEntry.targetAmount = entry.toPayout || 0;
    });

    const revenueTransferTargets = settlementParticipants
      .map(participant => ({
        personId: participant.personId,
        personName: participant.personName,
        personType: participant.personType,
        targetAmount: participant.desiredInvoiceAmount || 0
      }))
      .filter(entry => !!entry.personId);

    const revenueDeficits = revenueTransferTargets
      .map(target => ({
        ...target,
        remainingNeed: Math.max(0, (target.targetAmount || 0) - (issuerMap[target.personId]?.revenueBalance || 0))
      }))
      .filter(target => target.remainingNeed >= 0.005);

    const revenuePayers = issuers
      .map(entry => ({
        issuerId: entry.issuerId,
        remainingAmount: Math.max(0, (entry.revenueBalance || 0) - Math.max(0, entry.revenueTargetAmount || 0))
      }))
      .filter(entry => entry.remainingAmount >= 0.005);

    revenuePayers.forEach(payerEntry => {
      const payer = issuerMap[payerEntry.issuerId];
      revenueDeficits.forEach(deficit => {
        if (!(payerEntry.remainingAmount >= 0.005) || !(deficit.remainingNeed >= 0.005)) return;
        if (deficit.personId === payerEntry.issuerId) return;

        const amount = Math.min(payerEntry.remainingAmount, deficit.remainingNeed);
        const transferType = personById[deficit.personId]?.type === 'SEPARATE_COMPANY'
          ? 'Przelew przychodu do osobnej firmy'
          : 'Wyrównanie przychodu wspólnika';

        payer.transferPayments.push({
          recipientId: deficit.personId,
          recipientName: deficit.personName,
          amount,
          type: transferType,
          category: 'revenue-equalization'
        });

        const recipient = issuerMap[deficit.personId];
        if (recipient) {
          recipient.incomingTransferPayments.push({
            payerId: payerEntry.issuerId,
            payerName: payer.issuerName,
            amount,
            type: transferType,
            category: 'revenue-equalization'
          });
          recipient.revenueBalance += amount;
          recipient.currentBalance += amount;
        }

        payerEntry.remainingAmount -= amount;
        payer.revenueBalance -= amount;
        payer.currentBalance -= amount;
        deficit.remainingNeed -= amount;
      });
    });

    const salaryRecipients = [
      ...(settlement.employees || []).map(entry => ({
        recipientId: entry.person?.id,
        recipientName: Calculations.getPersonDisplayName(entry.person),
        payerId: entry.person?.employerId,
        amount: entry.toPayout || 0,
        type: 'Pensja pracownika'
      })),
      ...(settlement.workingPartners || [])
        .filter(entry => !!entry.person?.employerId)
        .map(entry => ({
          recipientId: entry.person?.id,
          recipientName: Calculations.getPersonDisplayName(entry.person),
          payerId: entry.person?.employerId,
          amount: entry.netAfterAccounting || 0,
          type: 'Wypłata wspólnika pracującego'
        }))
    ].filter(entry => !!entry.payerId && Math.abs(entry.amount) >= 0.005);

    salaryRecipients.forEach(payment => {
      const payer = issuerMap[payment.payerId];
      if (!payer) return;
      payer.salaryPayments.push(payment);
      payer.revenueBalance -= payment.amount || 0;
      payer.currentBalance -= payment.amount || 0;
    });

    const officePayments = [
      ...issuers
        .filter(entry => entry.issuerType === 'PARTNER' || entry.issuerType === 'WORKING_PARTNER')
        .map(entry => {
          const issuedAmount = invoiceTaxEqualization.actualIssuedAmountByIssuer?.[entry.issuerId] || 0;
          const taxAmount = invoiceTaxEqualization.invoiceTaxByIssuer?.[entry.issuerId] || 0;
          const taxRate = invoiceTaxEqualization.issuerTaxRateByIssuer?.[entry.issuerId] || 0;
          return {
            payerId: entry.issuerId,
            amount: taxAmount,
            type: `Podatek ryczałtu od faktur ${Calculations.formatInvoiceCurrency(issuedAmount)} × ${(taxRate * 100).toFixed(2)}%`,
            category: 'invoice-tax'
          };
        }),
      ...issuers
        .filter(entry => (entry.issuerType === 'PARTNER' || entry.issuerType === 'WORKING_PARTNER') && Math.abs(entry.ownZusAmount || 0) >= 0.005)
        .map(entry => ({
          payerId: entry.issuerId,
          amount: entry.ownZusAmount || 0,
          type: 'ZUS',
          category: 'own-zus'
        })),
      ...(settlement.partners || []).map(entry => ({
        payerId: entry.person?.id,
        amount: entry.employeeAccountingRefund || 0,
        type: 'Podatek i ZUS za pracowników do urzędu',
        category: 'employee-office'
      })),
      ...(settlement.workingPartners || []).map(entry => ({
        payerId: entry.person?.id,
        amount: entry.employeeAccountingRefund || 0,
        type: 'Podatek i ZUS za pracowników do urzędu',
        category: 'employee-office'
      })),
      ...((settlement.separateCompanies || []).map(entry => ({
        payerId: entry.person?.id,
        amount: entry.employeeAccountingRefund || 0,
        type: 'Podatek i ZUS za pracowników do urzędu',
        category: 'employee-office'
      })))
    ].filter(entry => !!entry.payerId && Math.abs(entry.amount) >= 0.005);

    officePayments.forEach(payment => {
      const payer = issuerMap[payment.payerId];
      if (!payer) return;
      payer.officePayments.push({
        recipientId: 'office',
        recipientName: 'Urząd',
        amount: payment.amount,
        type: payment.type,
        category: payment.category || 'employee-office'
      });
      if (payment.category === 'employee-office') {
        payer.revenueBalance -= payment.amount || 0;
      }
      payer.currentBalance -= payment.amount || 0;
    });

    issuers.forEach(entry => {
      entry.revenueRetainedAmount = entry.revenueBalance || 0;
    });

    (invoiceTaxEqualization.reimbursements || []).forEach(payment => {
      const payer = issuerMap[payment.payerId];
      const recipient = issuerMap[payment.recipientId];
      if (!payer || !recipient) return;

      payer.transferPayments.push({
        recipientId: payment.recipientId,
        recipientName: payment.recipientName,
        amount: payment.taxAmount,
        type: payment.type,
        category: 'tax-reimbursement'
      });
      recipient.incomingTransferPayments.push({
        payerId: payment.payerId,
        payerName: payment.payerName,
        amount: payment.taxAmount,
        type: payment.type,
        category: 'tax-reimbursement'
      });
      payer.currentBalance -= payment.taxAmount || 0;
      recipient.currentBalance += payment.taxAmount || 0;
    });

    issuers.forEach(entry => {
      const outgoingTaxReimbursements = (entry.transferPayments || [])
        .filter(payment => payment.category === 'tax-reimbursement')
        .reduce((sum, payment) => sum + (payment.amount || 0), 0);
      entry.postTaxTargetAmount = entry.issuerType === 'SEPARATE_COMPANY'
        ? ((entry.grossTargetAmount || 0) - outgoingTaxReimbursements)
        : ((entry.grossTargetAmount || 0) - (entry.targetTaxAmount || 0) - (entry.ownZusAmount || 0));
    });

    const unresolvedRecipients = issuers
      .map(entry => ({
        personId: entry.issuerId,
        personName: entry.issuerName,
        remainingNeed: Math.max(0, (entry.postTaxTargetAmount || 0) - (entry.currentBalance || 0))
      }))
      .filter(entry => entry.remainingNeed >= 0.005);

    return {
      month: selectedMonth,
      issuers: issuers.map(entry => ({
        ...entry,
        retainedAmount: entry.currentBalance || 0,
        differenceToTarget: (entry.currentBalance || 0) - (entry.postTaxTargetAmount || 0)
      })),
      unresolvedRecipients
    };
  },

  generateSettlement: (state, options = {}) => {
    const selectedMonth = Calculations.getSelectedMonth(state);
    const scopedState = Calculations._wrapState(state, selectedMonth);
    state = scopedState;
    const settlementConfig = Calculations.getSettlementConfig(state, selectedMonth);
    const includeInvoiceTaxEqualization = options.includeInvoiceTaxEqualization !== false;
    const includeInvoiceReconciliation = options.includeInvoiceReconciliation !== false;
    const activePersons = Calculations.getActivePersons(state, selectedMonth);
    const persons = Calculations.getSettlementPersons(state, selectedMonth);
    const selectedExpenses = (state.expenses || []).filter(expense => Calculations.isDateInMonth(expense.date, selectedMonth));
    const partners = persons.filter(p => p.type === 'PARTNER');
    const separateCompanies = persons.filter(p => p.type === 'SEPARATE_COMPANY');
    const workingPartners = persons.filter(p => p.type === 'WORKING_PARTNER');
    const employees = persons.filter(p => p.type === 'EMPLOYEE');
    const activePartnerIds = new Set(activePersons.filter(p => p.type === 'PARTNER').map(p => p.id));
    const activeSeparateCompanyIds = new Set(activePersons.filter(p => p.type === 'SEPARATE_COMPANY').map(p => p.id));
    const revenueBreakdown = Calculations.calculateRevenueBreakdown(state, persons, activePersons);
    const personById = Object.fromEntries((state.persons || []).map(person => [person.id, person]));

    let totalTeamHours = 0;

    const empStats = employees.map(emp => {
      const stats = Calculations.calculatePersonStats(emp, state, 0);
      const hourSheetStats = Calculations.calculatePersonStats(emp, state, 0, { includeWorksSheets: false });
      const revenueFromHourBasedWork = Calculations.calculateEmployeeRevenueFromHourBasedWork(emp, state);
      const hourBasedSalaryStats = Calculations.calculatePersonStats(emp, state, 0, { worksOnlyRoboczogodziny: true });
      const revenueFromHoursSheets = Calculations.calculatePersonRevenueFromMonthlySheets(emp, state);
      const adjustedRevenueFromHoursSheets = Calculations.calculatePersonRevenueFromMonthlySheets(emp, state, { useAdjustedRate: true });
      const isActiveInMonth = Calculations.isPersonActiveInMonth(emp, state, selectedMonth);
      totalTeamHours += stats.totalHours;
      return {
        person: emp,
        hours: stats.totalHours,
        salary: stats.totalSalary,
        salaryFromHoursSheets: hourSheetStats.totalSalary,
        revenueFromHourBasedWork,
        revenueFromHoursSheets,
        adjustedRevenueFromHoursSheets,
        rateDifferenceFromHoursSheets: revenueFromHoursSheets - adjustedRevenueFromHoursSheets,
        generatedProfit: adjustedRevenueFromHoursSheets + (revenueFromHourBasedWork - revenueFromHoursSheets) - hourBasedSalaryStats.totalSalary,
        generatedProfitFromHoursSheets: adjustedRevenueFromHoursSheets - hourSheetStats.totalSalary,
        isActiveInMonth,
        effectiveRate: stats.totalHours > 0 ? stats.totalSalary / stats.totalHours : (emp.hourlyRate || 0)
      };
    });

    const partnerStats = partners.map(p => {
      const stats = Calculations.calculatePersonStats(p, state, 0, { worksOnlyRoboczogodziny: true });
      totalTeamHours += stats.totalHours;
      return {
        person: p,
        hours: stats.totalHours,
        salary: stats.totalSalary,
        effectiveRate: stats.totalHours > 0 ? stats.totalSalary / stats.totalHours : (p.hourlyRate || 0)
      };
    });

    const separateCompanyStats = separateCompanies.map(company => {
      const stats = Calculations.calculatePersonStats(company, state, 0, { worksOnlyRoboczogodziny: true });
      const revenueFromHoursSheets = Calculations.calculatePersonRevenueFromMonthlySheets(company, state);
      const adjustedRevenueFromHoursSheets = Calculations.calculatePersonRevenueFromMonthlySheets(company, state, { useAdjustedRate: true });
      const revenueFromWorksRoboczogodziny = Calculations.calculatePersonRevenueFromWorksRoboczogodziny(company, state);
      const adjustedRevenueFromWorksRoboczogodziny = Calculations.calculatePersonRevenueFromWorksRoboczogodziny(company, state, { useAdjustedRate: true });
      totalTeamHours += stats.totalHours;
      return {
        person: company,
        hours: stats.totalHours,
        salary: stats.totalSalary,
        effectiveRate: stats.totalHours > 0 ? stats.totalSalary / stats.totalHours : (company.hourlyRate || 0),
        ownHoursProfit: (adjustedRevenueFromHoursSheets + adjustedRevenueFromWorksRoboczogodziny) - stats.totalSalary,
        rateDifferenceProfit: (revenueFromHoursSheets - adjustedRevenueFromHoursSheets) + (revenueFromWorksRoboczogodziny - adjustedRevenueFromWorksRoboczogodziny)
      };
    });

    const workingPartnerStats = workingPartners.map(wp => {
      const stats = Calculations.calculatePersonStats(wp, state, 0, { worksOnlyRoboczogodziny: true });
      totalTeamHours += stats.totalHours;
      const effectiveRate = stats.totalHours > 0 ? stats.totalSalary / stats.totalHours : (wp.hourlyRate || 0);
      return { person: wp, hours: stats.totalHours, salary: stats.totalSalary, effectiveRate };
    });

    const commonRevenue = revenueBreakdown.totalRevenue;

    const clientAdvances = selectedExpenses
      .filter(e => e.type === 'ADVANCE' && e.paidById && e.paidById.startsWith('client_'))
      .reduce((sum, e) => sum + e.amount, 0);

    const employeeBonusByRecipientId = selectedExpenses
      .filter(expense => expense.type === 'BONUS' && expense.advanceForId)
      .reduce((totals, expense) => {
        totals[expense.advanceForId] = (totals[expense.advanceForId] || 0) + Calculations.getExpenseEffectiveAmount(expense, state, selectedMonth);
        return totals;
      }, {});

    const employeeDietaByRecipientId = selectedExpenses
      .filter(expense => expense.type === 'DIETA' && expense.advanceForId)
      .reduce((totals, expense) => {
        totals[expense.advanceForId] = (totals[expense.advanceForId] || 0) + Calculations.getExpenseEffectiveAmount(expense, state, selectedMonth);
        return totals;
      }, {});

    const refundExpenses = selectedExpenses.filter(expense => expense.type === 'REFUND' && expense.paidById);

    const processedEmployees = empStats.map(emp => {
      const employer = emp.person.employerId ? personById[emp.person.employerId] : null;
      const skipSettlementContractCharges = Calculations.isSeparateCompany(employer)
        && !Calculations.countsEmployeeAccountingRefund(employer);
      const contractCharges = emp.isActiveInMonth && !skipSettlementContractCharges
        ? Calculations.getPersonContractCharges(emp.person, state, selectedMonth)
        : { taxAmount: 0, zusAmount: 0, total: 0, paidByEmployer: false };
      const paidCosts = selectedExpenses
        .filter(e => e.type === 'COST' && e.paidById === emp.person.id)
        .reduce((sum, e) => sum + e.amount, 0);
      const paidAdvances = selectedExpenses
        .filter(e => e.type === 'ADVANCE' && e.paidById === emp.person.id)
        .reduce((sum, e) => sum + e.amount, 0);
      const advancesTaken = selectedExpenses
        .filter(e => e.type === 'ADVANCE' && e.advanceForId === emp.person.id)
        .reduce((sum, e) => sum + e.amount, 0);
      const refundsPaid = refundExpenses
        .filter(expense => expense.paidById === emp.person.id)
        .reduce((sum, expense) => sum + Calculations.getExpenseEffectiveAmount(expense, state, selectedMonth), 0);
      const refundsReceived = refundExpenses
        .reduce((sum, expense) => sum + Calculations.getRefundReceivedAmountForPerson(expense, emp.person.id, state, selectedMonth), 0);
      const bonusAmount = employeeBonusByRecipientId[emp.person.id] || 0;
      const dietaAmount = employeeDietaByRecipientId[emp.person.id] || 0;
      const totalBenefitAmount = bonusAmount + dietaAmount;
      const deductedContractTaxAmount = contractCharges.paidByEmployer ? 0 : contractCharges.taxAmount;
      const deductedContractZusAmount = contractCharges.paidByEmployer ? 0 : contractCharges.zusAmount;
      const employerPaidContractTaxAmount = contractCharges.paidByEmployer ? contractCharges.taxAmount : 0;
      const employerPaidContractZusAmount = contractCharges.paidByEmployer ? contractCharges.zusAmount : 0;

      return {
        ...emp,
        employer,
        paidCosts,
        paidAdvances,
        advancesTaken,
        refundsPaid,
        refundsReceived,
        bonusAmount,
        dietaAmount,
        totalBenefitAmount,
        contractTaxAmount: contractCharges.taxAmount,
        contractZusAmount: contractCharges.zusAmount,
        deductedContractTaxAmount,
        deductedContractZusAmount,
        employerPaidContractTaxAmount,
        employerPaidContractZusAmount,
        employerPaidContractCharges: employerPaidContractTaxAmount + employerPaidContractZusAmount,
        contractChargesPaidByEmployer: contractCharges.paidByEmployer,
        toPayout: emp.salary + totalBenefitAmount + paidCosts + paidAdvances + refundsReceived - refundsPaid - advancesTaken - deductedContractTaxAmount - deductedContractZusAmount
      };
    });

    const processedEmployeesVisible = processedEmployees.filter(person =>
      Calculations.isPersonActiveInMonth(person.person, state, selectedMonth) || Math.abs(person.toPayout) >= 0.005
    );

    const totalCommonCosts = (state?.expenses || [])
      .filter(e => Calculations.isDateInMonth(e.date, selectedMonth) && e.type === 'COST')
      .reduce((sum, e) => sum + e.amount, 0);
    const activeCostSharePersonIds = new Set(
      activePersons
        .filter(person => Calculations.isPartnerLike(person) && Calculations.personParticipatesInCosts(person))
        .map(person => person.id)
    );
    const costShare = activeCostSharePersonIds.size > 0 ? totalCommonCosts / activeCostSharePersonIds.size : 0;

    const processedWorkingPartnersBase = workingPartnerStats.map(wp => {
      const employer = wp.person.employerId ? personById[wp.person.employerId] : null;
      const skipSettlementContractCharges = Calculations.isSeparateCompany(employer)
        && !Calculations.countsEmployeeAccountingRefund(employer);
      const contractCharges = !skipSettlementContractCharges
        ? Calculations.getPersonContractCharges(wp.person, state, selectedMonth)
        : { taxAmount: 0, zusAmount: 0, total: 0, paidByEmployer: false };
      const paidCosts = selectedExpenses
        .filter(e => e.type === 'COST' && e.paidById === wp.person.id)
        .reduce((sum, e) => sum + e.amount, 0);
      const paidAdvances = selectedExpenses
        .filter(e => e.type === 'ADVANCE' && e.paidById === wp.person.id)
        .reduce((sum, e) => sum + e.amount, 0);
      const advancesTaken = selectedExpenses
        .filter(e => e.type === 'ADVANCE' && e.advanceForId === wp.person.id)
        .reduce((sum, e) => sum + e.amount, 0);
      const refundsPaid = refundExpenses
        .filter(expense => expense.paidById === wp.person.id)
        .reduce((sum, expense) => sum + Calculations.getExpenseEffectiveAmount(expense, state, selectedMonth), 0);
      const refundsReceived = refundExpenses
        .reduce((sum, expense) => sum + Calculations.getRefundReceivedAmountForPerson(expense, wp.person.id, state, selectedMonth), 0);
      const worksShare = revenueBreakdown.worksRevenueByPartner[wp.person.id] || 0;
      const appliedCostShare = activeCostSharePersonIds.has(wp.person.id) ? costShare : 0;
      const grossBeforeAccounting = wp.salary + worksShare + paidCosts + paidAdvances + refundsReceived - refundsPaid - appliedCostShare - advancesTaken;
      const ryczaltTaxAmount = Calculations.calculateAccountingTax(grossBeforeAccounting, settlementConfig.taxRate);
      const deductedContractTaxAmount = contractCharges.paidByEmployer ? 0 : contractCharges.taxAmount;
      const deductedContractZusAmount = contractCharges.paidByEmployer ? 0 : contractCharges.zusAmount;
      const employerPaidContractTaxAmount = contractCharges.paidByEmployer ? contractCharges.taxAmount : 0;
      const employerPaidContractZusAmount = contractCharges.paidByEmployer ? contractCharges.zusAmount : 0;
      const taxAmount = ryczaltTaxAmount + deductedContractTaxAmount;
      const zusAmount = deductedContractZusAmount;

      return {
        ...wp,
        employer,
        worksShare,
        costShareApplied: appliedCostShare,
        isActiveInMonth: Calculations.isPersonActiveInMonth(wp.person, state, selectedMonth),
        paidCosts,
        paidAdvances,
        advancesTaken,
        refundsPaid,
        refundsReceived,
        ryczaltTaxAmount,
        contractTaxAmount: contractCharges.taxAmount,
        contractZusAmount: contractCharges.zusAmount,
        deductedContractTaxAmount,
        deductedContractZusAmount,
        employerPaidContractTaxAmount,
        employerPaidContractZusAmount,
        employerPaidContractCharges: employerPaidContractTaxAmount + employerPaidContractZusAmount,
        contractChargesPaidByEmployer: contractCharges.paidByEmployer,
        taxAmount,
        zusAmount,
        toPayout: grossBeforeAccounting,
        netAfterAccounting: Calculations.calculateNetAfterAccounting(grossBeforeAccounting, taxAmount, zusAmount)
      };
    });

    const employeeProfitWeights = activePersons.reduce((acc, person) => {
      if (!Calculations.isEmployeeProfitRecipient(person)) return acc;
      const weight = Calculations.getEmployeeProfitRecipientWeight(person);
      if (weight > 0) acc[person.id] = weight;
      return acc;
    }, {});
    const totalEmployeeProfitWeight = Object.values(employeeProfitWeights).reduce((sum, weight) => sum + weight, 0);
    const distributedEmployeeProfitByRecipient = {};
    const retainedEmployeeProfitByCompany = {};
    const companyRecipientPercents = Object.entries(employeeProfitWeights)
      .map(([personId, weight]) => ({ personId, weight, person: personById[personId] }))
      .filter(entry => Calculations.isSeparateCompany(entry.person))
      .map(entry => ({
        personId: entry.personId,
        name: Calculations.getPersonDisplayName(entry.person) || 'Osobna Firma',
        percent: totalEmployeeProfitWeight > 0 ? (entry.weight / totalEmployeeProfitWeight) * 100 : 0
      }));
    const ownEmployeeProfitItems = [];
    const separateCompanyEmployeeProfitItems = [];
    const separateCompanyRateDifferenceItems = [];

    const distributeEmployeeProfit = (amount) => {
      if (!Number.isFinite(amount) || Math.abs(amount) < 0.005 || totalEmployeeProfitWeight <= 0) return;
      Object.entries(employeeProfitWeights).forEach(([personId, weight]) => {
        distributedEmployeeProfitByRecipient[personId] = (distributedEmployeeProfitByRecipient[personId] || 0) + (amount * weight / totalEmployeeProfitWeight);
      });
    };

    processedEmployees.forEach(employee => {
      const employerCoveredContractCharges = Calculations.isSeparateCompany(employee.employer)
        ? (Calculations.countsEmployeeAccountingRefund(employee.employer) ? (employee.employerPaidContractCharges || 0) : 0)
        : (employee.employerPaidContractCharges || 0);
      const profitContribution = (employee.generatedProfitFromHoursSheets || 0) - employerCoveredContractCharges;
      const adjustedProfitContribution = profitContribution - (employee.totalBenefitAmount || 0);
      const rateDifferenceProfit = employee.rateDifferenceFromHoursSheets || 0;

      if (Calculations.isSeparateCompany(employee.employer)) {
        const sharePercent = Calculations.getSeparateCompanySharedProfitPercent(employee.employer);
        const sharedAmount = adjustedProfitContribution * (sharePercent / 100);
        const retainedAmount = adjustedProfitContribution - sharedAmount;
        retainedEmployeeProfitByCompany[employee.employer.id] = (retainedEmployeeProfitByCompany[employee.employer.id] || 0) + retainedAmount;
        if (Number.isFinite(adjustedProfitContribution) && Math.abs(adjustedProfitContribution) >= 0.005) {
          separateCompanyEmployeeProfitItems.push({
            companyId: employee.employer.id,
            companyName: Calculations.getPersonDisplayName(employee.employer),
            employeeId: employee.person.id,
            employeeName: Calculations.getPersonDisplayName(employee.person),
            profit: adjustedProfitContribution,
            sharePercent,
            distributedProfit: sharedAmount,
            retainedProfit: retainedAmount
          });
          distributeEmployeeProfit(sharedAmount);
        }
        if (Number.isFinite(rateDifferenceProfit) && Math.abs(rateDifferenceProfit) >= 0.005) {
          separateCompanyRateDifferenceItems.push({
            companyId: employee.employer.id,
            companyName: Calculations.getPersonDisplayName(employee.employer),
            sourceName: Calculations.getPersonDisplayName(employee.person),
            profit: rateDifferenceProfit,
            type: 'employee-rate-difference'
          });
          distributeEmployeeProfit(rateDifferenceProfit);
        }
        return;
      }

      if (Number.isFinite(adjustedProfitContribution) && Math.abs(adjustedProfitContribution) >= 0.005) {
        ownEmployeeProfitItems.push({
          employeeId: employee.person.id,
          employeeName: Calculations.getPersonDisplayName(employee.person),
          profit: adjustedProfitContribution,
          separateCompanyPercents: companyRecipientPercents
        });
        distributeEmployeeProfit(adjustedProfitContribution);
      }
    });

    separateCompanyStats.forEach(company => {
      if (Number.isFinite(company.ownHoursProfit || 0) && Math.abs(company.ownHoursProfit || 0) >= 0.005) {
        separateCompanyEmployeeProfitItems.push({
          companyId: company.person.id,
          companyName: Calculations.getPersonDisplayName(company.person),
          employeeId: company.person.id,
          employeeName: `${Calculations.getPersonDisplayName(company.person)} (własne godziny)`,
          profit: company.ownHoursProfit || 0,
          sharePercent: 100,
          distributedProfit: company.ownHoursProfit || 0,
          retainedProfit: 0
        });
        distributeEmployeeProfit(company.ownHoursProfit || 0);
      }
      if (Number.isFinite(company.rateDifferenceProfit || 0) && Math.abs(company.rateDifferenceProfit || 0) >= 0.005) {
        separateCompanyRateDifferenceItems.push({
          companyId: company.person.id,
          companyName: Calculations.getPersonDisplayName(company.person),
          sourceName: `${Calculations.getPersonDisplayName(company.person)} (różnica stawki)`,
          profit: company.rateDifferenceProfit || 0,
          type: 'company-rate-difference'
        });
        distributeEmployeeProfit(company.rateDifferenceProfit || 0);
      }
    });

    const totalDistributedEmployeeProfit = Object.values(distributedEmployeeProfitByRecipient).reduce((sum, value) => sum + value, 0);
    const totalOwnEmployeeProfit = ownEmployeeProfitItems.reduce((sum, item) => sum + item.profit, 0);
    const totalSeparateCompanyEmployeeProfit = separateCompanyEmployeeProfitItems.reduce((sum, item) => sum + item.distributedProfit, 0);
    const totalSeparateCompanyRateDifferenceProfit = separateCompanyRateDifferenceItems.reduce((sum, item) => sum + item.profit, 0);
    const totalEmployeeBonuses = processedEmployees.reduce((sum, employee) => sum + (employee.bonusAmount || 0), 0);
    const totalEmployeeDietas = processedEmployees.reduce((sum, employee) => sum + (employee.dietaAmount || 0), 0);
    const totalEmployeesSalary = processedEmployees.reduce((sum, employee) => {
      if (Calculations.isSeparateCompany(employee.employer)) return sum;
      return sum + (employee.salaryFromHoursSheets || 0);
    }, 0);
    const employeeRevenue = processedEmployees.reduce((sum, employee) => {
      if (Calculations.isSeparateCompany(employee.employer)) return sum;
      return sum + (employee.revenueFromHoursSheets || 0);
    }, 0);

    const employerEmployeeSalaryByEmployer = processedEmployeesVisible.map(employee => ({
      employerId: employee.person.employerId,
      payout: employee.toPayout || 0,
      employerPaidContractCharges: employee.employerPaidContractCharges || 0,
      bonusAmount: employee.bonusAmount || 0,
      dietaAmount: employee.dietaAmount || 0,
      totalBenefitAmount: employee.totalBenefitAmount || 0
    })).reduce((totals, entry) => {
      const employerId = entry.employerId;
      if (!employerId) return totals;

      if (!totals[employerId]) {
        totals[employerId] = {
          positive: 0,
          receivable: 0,
          signed: 0,
          accountingRefund: 0,
          bonuses: 0,
          dietas: 0,
          benefits: 0
        };
      }

      const payout = entry.payout;
      if (payout >= 0) {
        totals[employerId].positive += payout;
      } else {
        totals[employerId].receivable += Math.abs(payout);
      }
      totals[employerId].signed += payout;
      totals[employerId].accountingRefund += entry.employerPaidContractCharges || 0;
      totals[employerId].bonuses += entry.bonusAmount || 0;
      totals[employerId].dietas += entry.dietaAmount || 0;
      totals[employerId].benefits += entry.totalBenefitAmount || 0;

      return totals;
    }, {});

    processedWorkingPartnersBase.forEach(wp => {
      if (!wp.person.employerId) return;

      const employerId = wp.person.employerId;
      if (!employerEmployeeSalaryByEmployer[employerId]) {
        employerEmployeeSalaryByEmployer[employerId] = {
          positive: 0,
          receivable: 0,
          signed: 0,
          accountingRefund: 0,
          bonuses: 0,
          dietas: 0,
          benefits: 0
        };
      }

      const employerTotals = employerEmployeeSalaryByEmployer[employerId];
      const netPayout = wp.netAfterAccounting || 0;
      if (netPayout >= 0) {
        employerTotals.positive += netPayout;
      } else {
        employerTotals.receivable += Math.abs(netPayout);
      }
      employerTotals.signed += netPayout;
      employerTotals.accountingRefund += (wp.ryczaltTaxAmount || 0)
        + (wp.deductedContractTaxAmount || 0)
        + (wp.deductedContractZusAmount || 0)
        + (wp.employerPaidContractCharges || 0);
    });

    let processedPartners = partnerStats.map(p => {
      const paidCosts = selectedExpenses
        .filter(e => e.type === 'COST' && e.paidById === p.person.id)
        .reduce((sum, e) => sum + e.amount, 0);
      const paidAdvances = selectedExpenses
        .filter(e => e.type === 'ADVANCE' && e.paidById === p.person.id)
        .reduce((sum, e) => sum + e.amount, 0);
      const advancesTaken = selectedExpenses
        .filter(e => e.type === 'ADVANCE' && e.advanceForId === p.person.id)
        .reduce((sum, e) => sum + e.amount, 0);
      const refundsPaid = refundExpenses
        .filter(expense => expense.paidById === p.person.id)
        .reduce((sum, expense) => sum + Calculations.getExpenseEffectiveAmount(expense, state, selectedMonth), 0);
      const refundsReceived = refundExpenses
        .reduce((sum, expense) => sum + Calculations.getRefundReceivedAmountForPerson(expense, p.person.id, state, selectedMonth), 0);
      const worksShare = revenueBreakdown.worksRevenueByPartner[p.person.id] || 0;
      const revenueShare = distributedEmployeeProfitByRecipient[p.person.id] || 0;
      const appliedCostShare = activeCostSharePersonIds.has(p.person.id) ? costShare : 0;
      const totalRevenueShare = revenueShare + worksShare;
      const employerEmployeeTotals = employerEmployeeSalaryByEmployer[p.person.id] || { positive: 0, receivable: 0, signed: 0, accountingRefund: 0, bonuses: 0, dietas: 0, benefits: 0 };
      const employeeSalaries = employerEmployeeTotals.positive;
      const employeeReceivables = employerEmployeeTotals.receivable;
      const employeeAccountingRefund = employerEmployeeTotals.accountingRefund;
      const employeeBonuses = employerEmployeeTotals.bonuses || 0;
      const employeeDietas = employerEmployeeTotals.dietas || 0;
      const employeeBenefits = employerEmployeeTotals.benefits || 0;
      const ownGrossAmount = p.salary + totalRevenueShare;
      const sharedCompanyTaxBase = paidAdvances + employeeSalaries + employeeAccountingRefund - employeeReceivables;
      const toPayout = ownGrossAmount + paidCosts + paidAdvances + refundsReceived - refundsPaid - appliedCostShare - advancesTaken;
      const zusAmount = settlementConfig.zusFixedAmount;
      const ownTaxBase = ownGrossAmount;
      const ownTaxAmount = Calculations.calculateAccountingTax(ownTaxBase, settlementConfig.taxRate);
      const taxAmount = ownTaxAmount;

      return {
        ...p,
        revenueShare,
        worksShare,
        totalRevenueShare,
        employeeSalaries,
        employeeReceivables,
        employeeAccountingRefund,
        employeeBonuses,
        employeeDietas,
        employeeBenefits,
        grossWithEmployeeSalaries: toPayout + employerEmployeeTotals.signed + employeeAccountingRefund,
        ownGrossAmount,
        sharedCompanyTaxBase,
        costShareApplied: appliedCostShare,
        isActiveInMonth: activePartnerIds.has(p.person.id),
        paidCosts,
        paidAdvances,
        advancesTaken,
        refundsPaid,
        refundsReceived,
        ownTaxBase,
        ownTaxAmount,
        sharedCompanyTaxAmount: 0,
        taxAmount,
        zusAmount,
        toPayout,
        netAfterAccounting: Calculations.calculateNetAfterAccounting(toPayout, taxAmount, zusAmount)
      };
    }).filter(person => person.isActiveInMonth || Math.abs(person.toPayout) >= 0.005 || Math.abs(person.employeeSalaries) >= 0.005 || Math.abs(person.employeeReceivables) >= 0.005);

    const sharedCompanyTaxParticipants = processedPartners.filter(person => person.isActiveInMonth);
    const partnerSharedCompanyTaxBaseTotal = sharedCompanyTaxParticipants.reduce((sum, person) => sum + (parseFloat(person.sharedCompanyTaxBase) || 0), 0);
    const partnerSharedCompanyTaxTotal = Calculations.calculateAccountingTax(partnerSharedCompanyTaxBaseTotal, settlementConfig.taxRate);
    const partnerSharedCompanyTaxPerPerson = sharedCompanyTaxParticipants.length > 0
      ? partnerSharedCompanyTaxTotal / sharedCompanyTaxParticipants.length
      : 0;

    processedPartners = processedPartners.map(person => {
      const sharedCompanyTaxAmount = person.isActiveInMonth ? partnerSharedCompanyTaxPerPerson : 0;
      const taxAmount = (parseFloat(person.ownTaxAmount) || 0) + sharedCompanyTaxAmount;
      return {
        ...person,
        sharedCompanyTaxParticipantsCount: sharedCompanyTaxParticipants.length,
        sharedCompanyTaxBaseTotal: partnerSharedCompanyTaxBaseTotal,
        sharedCompanyTaxTotal: partnerSharedCompanyTaxTotal,
        sharedCompanyTaxAmount,
        taxAmount,
        netAfterAccounting: Calculations.calculateNetAfterAccounting(person.toPayout, taxAmount, person.zusAmount)
      };
    });

    const processedSeparateCompanies = separateCompanyStats.map(company => {
      const paidCosts = selectedExpenses
        .filter(e => e.type === 'COST' && e.paidById === company.person.id)
        .reduce((sum, e) => sum + e.amount, 0);
      const paidAdvances = selectedExpenses
        .filter(e => e.type === 'ADVANCE' && e.paidById === company.person.id)
        .reduce((sum, e) => sum + e.amount, 0);
      const advancesTaken = selectedExpenses
        .filter(e => e.type === 'ADVANCE' && e.advanceForId === company.person.id)
        .reduce((sum, e) => sum + e.amount, 0);
      const refundsPaid = refundExpenses
        .filter(expense => expense.paidById === company.person.id)
        .reduce((sum, expense) => sum + Calculations.getExpenseEffectiveAmount(expense, state, selectedMonth), 0);
      const refundsReceived = refundExpenses
        .reduce((sum, expense) => sum + Calculations.getRefundReceivedAmountForPerson(expense, company.person.id, state, selectedMonth), 0);
      const worksShare = revenueBreakdown.worksRevenueByPartner[company.person.id] || 0;
      const revenueShare = (distributedEmployeeProfitByRecipient[company.person.id] || 0) + (retainedEmployeeProfitByCompany[company.person.id] || 0);
      const appliedCostShare = activeCostSharePersonIds.has(company.person.id) ? costShare : 0;
      const totalRevenueShare = revenueShare + worksShare;
      const employerEmployeeTotals = employerEmployeeSalaryByEmployer[company.person.id] || { positive: 0, receivable: 0, signed: 0, accountingRefund: 0, bonuses: 0, dietas: 0, benefits: 0 };
      const employeeSalaries = employerEmployeeTotals.positive;
      const employeeReceivables = employerEmployeeTotals.receivable;
      const employeeAccountingRefund = Calculations.countsEmployeeAccountingRefund(company.person)
        ? employerEmployeeTotals.accountingRefund
        : 0;
      const employeeBonuses = employerEmployeeTotals.bonuses || 0;
      const employeeDietas = employerEmployeeTotals.dietas || 0;
      const employeeBenefits = employerEmployeeTotals.benefits || 0;
      const toPayout = company.salary + totalRevenueShare + paidCosts + paidAdvances + refundsReceived - refundsPaid - appliedCostShare - advancesTaken;

      return {
        ...company,
        revenueShare,
        worksShare,
        totalRevenueShare,
        employeeSalaries,
        employeeReceivables,
        employeeAccountingRefund,
        employeeBonuses,
        employeeDietas,
        employeeBenefits,
        grossWithEmployeeSalaries: toPayout + employerEmployeeTotals.signed + employeeAccountingRefund,
        costShareApplied: appliedCostShare,
        isActiveInMonth: activeSeparateCompanyIds.has(company.person.id),
        paidCosts,
        paidAdvances,
        advancesTaken,
        refundsPaid,
        refundsReceived,
        taxAmount: 0,
        zusAmount: 0,
        toPayout,
        netAfterAccounting: toPayout,
        ownHoursProfit: company.ownHoursProfit || 0
      };
    }).filter(person => person.isActiveInMonth || Math.abs(person.toPayout) >= 0.005 || Math.abs(person.employeeSalaries) >= 0.005 || Math.abs(person.employeeReceivables) >= 0.005);

    let processedWorkingPartners = processedWorkingPartnersBase.map(wp => {
      const employerEmployeeTotals = employerEmployeeSalaryByEmployer[wp.person.id] || { positive: 0, receivable: 0, signed: 0, accountingRefund: 0, bonuses: 0, dietas: 0, benefits: 0 };
      const employeeSalaries = employerEmployeeTotals.positive;
      const employeeReceivables = employerEmployeeTotals.receivable;
      const employeeAccountingRefund = employerEmployeeTotals.accountingRefund;
      const employeeBonuses = employerEmployeeTotals.bonuses || 0;
      const employeeDietas = employerEmployeeTotals.dietas || 0;
      const employeeBenefits = employerEmployeeTotals.benefits || 0;
      return {
        ...wp,
        employeeSalaries,
        employeeReceivables,
        employeeAccountingRefund,
        employeeBonuses,
        employeeDietas,
        employeeBenefits,
        ownGrossAmount: wp.salary + wp.worksShare,
        ownTaxBase: wp.salary + wp.worksShare - (wp.costShareApplied || 0),
        ownTaxAmount: wp.ryczaltTaxAmount || 0,
        invoiceTaxAmount: 0,
        sharedCompanyTaxAmount: 0,
        grossWithEmployeeSalaries: wp.toPayout + employerEmployeeTotals.signed + employeeAccountingRefund
      };
    }).filter(person => person.isActiveInMonth || Math.abs(person.toPayout) >= 0.005 || Math.abs(person.employeeSalaries) >= 0.005 || Math.abs(person.employeeReceivables) >= 0.005);

    const totalEmployerPaidContractCharges = processedEmployees.reduce((sum, employee) => sum + (employee.employerPaidContractCharges || 0), 0)
      + processedWorkingPartnersBase.reduce((sum, workingPartner) => sum + (workingPartner.employerPaidContractCharges || 0), 0);
    const employerCosts = [...processedPartners, ...processedWorkingPartners]
      .reduce((sum, person) => sum + person.taxAmount + person.zusAmount, 0) + totalEmployerPaidContractCharges;

    const totalWorksProfit = revenueBreakdown.worksRevenue || 0;
    const profitToSplit = totalDistributedEmployeeProfit + totalWorksProfit;

    return {
      commonRevenue,
      totalTeamHours,
      employeeRevenue,
      employeeSalaryFromHours: totalEmployeesSalary,
      totalEmployeeBonuses,
      totalEmployeeDietas,
      employeeProfitShared: totalDistributedEmployeeProfit,
      profitBreakdown: {
        ownEmployees: ownEmployeeProfitItems,
        ownEmployeesTotal: totalOwnEmployeeProfit,
        separateCompanyEmployees: separateCompanyEmployeeProfitItems,
        separateCompanyEmployeesTotal: totalSeparateCompanyEmployeeProfit,
        separateCompanyRateDifferences: separateCompanyRateDifferenceItems,
        separateCompanyRateDifferencesTotal: totalSeparateCompanyRateDifferenceProfit,
        separateCompanyRecipientPercents: companyRecipientPercents
      },
      partners: processedPartners,
      separateCompanies: processedSeparateCompanies,
      workingPartners: processedWorkingPartners,
      employees: processedEmployeesVisible,
      settlementConfig,
      employerCosts,
      employerPaidContractCharges: totalEmployerPaidContractCharges,
      partnerSharedCompanyTaxBaseTotal,
      partnerSharedCompanyTaxTotal,
      partnerSharedCompanyTaxPerPerson,
      partnerSharedCompanyTaxParticipantsCount: sharedCompanyTaxParticipants.length,
      profitToSplit,
      totalWorksProfit,
      clientAdvances,
      invoiceReconciliation: includeInvoiceTaxEqualization && includeInvoiceReconciliation ? Calculations.calculateInvoiceReconciliation(state, selectedMonth) : null
    };
  }
};
