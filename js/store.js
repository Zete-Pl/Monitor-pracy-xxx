// Helper for unique IDs
const generateId = () => crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).substring(2, 9);
const getCurrentMonthKey = () => {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
};

const WORK_TRACKER_COMMON_KEY = 'workTrackerCommon';
const WORK_TRACKER_MONTH_PREFIX = 'workTrackerMonth_';
const WORK_TRACKER_LOCAL_SETTINGS_KEY = 'workTrackerLocalSettings';
const WORK_TRACKER_SYNC_META_KEY = 'workTrackerSyncMeta';
const WORK_TRACKER_LEGACY_KEY = 'workTrackerState'; // Dla migracji
const LEGACY_WORK_TRACKER_INVOICE_TOTALS_KEY = 'workTrackerInvoiceTotals';

const SHARED_META_ROOT = 'shared_meta';
const SHARED_HISTORY_ROOT = 'shared_history';
const SHARED_HISTORY_INDEX_ROOT = `${SHARED_HISTORY_ROOT}/index`;
const SHARED_HISTORY_ENTRIES_ROOT = `${SHARED_HISTORY_ROOT}/entries`;
const SHARED_DAILY_BACKUPS_ROOT = 'shared_backups_daily';
const SHARED_DAILY_BACKUPS_INDEX_ROOT = `${SHARED_DAILY_BACKUPS_ROOT}/index`;
const SHARED_DAILY_BACKUPS_ENTRIES_ROOT = `${SHARED_DAILY_BACKUPS_ROOT}/entries`;
const SHARED_MONTHLY_BACKUPS_ROOT = 'shared_backups_monthly';
const SHARED_MONTHLY_BACKUPS_INDEX_ROOT = `${SHARED_MONTHLY_BACKUPS_ROOT}/index`;
const SHARED_MONTHLY_BACKUPS_ENTRIES_ROOT = `${SHARED_MONTHLY_BACKUPS_ROOT}/entries`;
const LEGACY_SHARED_INVOICE_TOTALS_ROOT = 'shared_meta/invoiceTotals';

const FIREBASE_WRITE_DEBOUNCE_MS = 1500;

const LOCAL_HISTORY_LIMIT = 50;
const SHARED_HISTORY_LIMIT = 20;
const SHARED_DAILY_BACKUPS_LIMIT = 10;
const SHARED_MONTHLY_BACKUPS_LIMIT = 12;

const COMMON_SCOPE_KEYS = ['persons', 'clients', 'worksCatalog', 'config'];
const MONTH_SCOPE_KEYS = [
  'monthlySheets',
  'worksSheets',
  'expenses',
  'monthSettings.persons',
  'monthSettings.clients',
  'monthSettings.settlementConfig',
  'monthSettings.personContractCharges',
  'monthSettings.payouts',
  'monthSettings.invoices',
  'monthSettings.archive',
  'monthSettings.commonSnapshot.persons',
  'monthSettings.commonSnapshot.clients',
  'monthSettings.commonSnapshot.worksCatalog',
  'monthSettings.commonSnapshot.config'
];

const COMMON_SYNC_SCOPE_KEYS = ['common'];
const MONTH_SYNC_SCOPE_KEYS = [
  'monthlySheets',
  'worksSheets',
  'expenses',
  'monthSettings',
  'monthSettings.commonSnapshot'
];

const safeParseStorage = (key) => {
  try {
    const rawValue = localStorage.getItem(key);
    return rawValue ? JSON.parse(rawValue) : null;
  } catch {
    return null;
  }
};

const getScopeSnapshotFromState = (stateData = {}, scopeKey, month = Store.getSelectedMonth()) => {
  const normalizedState = normalizeState(stateData || {});
  const monthKey = month || Store.getSelectedMonth() || getCurrentMonthKey();
  const monthRecord = normalizedState.months?.[monthKey] || normalizeMonth(null);
  const monthSettings = monthRecord.monthSettings || normalizeMonth(null).monthSettings;

  switch (scopeKey) {
    case 'common.persons':
      return cloneData(normalizedState.common?.persons || []);
    case 'common.clients':
      return cloneData(normalizedState.common?.clients || []);
    case 'common.worksCatalog':
      return cloneData(normalizedState.common?.worksCatalog || []);
    case 'common.config':
      return cloneData(normalizedState.common?.config || {});
    case 'month.monthlySheets':
      return cloneData(monthRecord.monthlySheets || []);
    case 'month.worksSheets':
      return cloneData(monthRecord.worksSheets || []);
    case 'month.expenses':
      return cloneData(monthRecord.expenses || []);
    case 'month.monthSettings.persons':
      return cloneData(monthSettings.persons || {});
    case 'month.monthSettings.clients':
      return cloneData(monthSettings.clients || {});
    case 'month.monthSettings.settlementConfig':
      return cloneData(monthSettings.settlementConfig || {});
    case 'month.monthSettings.personContractCharges':
      return cloneData(monthSettings.personContractCharges || {});
    case 'month.monthSettings.payouts':
      return cloneData(monthSettings.payouts || normalizePayoutSettings({}));
    case 'month.monthSettings.invoices':
      return cloneData(monthSettings.invoices || normalizeInvoicesConfig({}));
    case 'month.monthSettings.archive':
      return cloneData({
        isArchived: monthSettings.isArchived === true,
        commonSnapshot: monthSettings.commonSnapshot || null
      });
    case 'month.monthSettings.commonSnapshot.persons':
      return cloneData(monthSettings.commonSnapshot?.persons || []);
    case 'month.monthSettings.commonSnapshot.clients':
      return cloneData(monthSettings.commonSnapshot?.clients || []);
    case 'month.monthSettings.commonSnapshot.worksCatalog':
      return cloneData(monthSettings.commonSnapshot?.worksCatalog || []);
    case 'month.monthSettings.commonSnapshot.config':
      return cloneData(monthSettings.commonSnapshot?.config || {});
    default:
      return null;
  }
};

const DEFAULT_WORKS_CATALOG = [
  { coreId: 'c_roboczogodziny', name: 'Roboczogodziny', unit: 'h', defaultPrice: 0 },
  { coreId: 'c_lawy', name: 'Ławy bet. konst.', unit: 'm³', defaultPrice: 200 },
  { coreId: 'c_chudziak', name: 'Chudziak', unit: 'm²', defaultPrice: 150 },
  { coreId: 'c_slupy', name: 'Słupy', unit: 'm³', defaultPrice: 800 },
  { coreId: 'c_sciany', name: 'Ściany', unit: 'm²', defaultPrice: 85 },
  { coreId: 'c_strop', name: 'Strop', unit: 'm²', defaultPrice: 120 },
  { coreId: 'c_scianki', name: 'Ścianki', unit: 'm²', defaultPrice: 180 },
  { coreId: 'c_podciagi', name: 'Podciągi', unit: 'm³', defaultPrice: 1150 },
  { coreId: 'c_slupki', name: 'Słupki duperele', unit: 'szt.', defaultPrice: 1200 },
  { coreId: 'c_plyta', name: 'Płyta fund.', unit: 'm³', defaultPrice: 200 },
  { coreId: 'c_zbrojenie', name: 'Zbrojenie', unit: 'kg', defaultPrice: 2 },
  { coreId: 'c_fundamenty', name: 'Fundamenty', unit: 'm²', defaultPrice: 80 },
  { coreId: 'c_schody', name: 'Schody', unit: 'szt.', defaultPrice: 200 }
];

const DEFAULT_CONTRACT_TAX_AMOUNT = 104;
const DEFAULT_CONTRACT_ZUS_AMOUNT = 490.27;
const DEFAULT_PROFIT_SHARE_PERCENT = 100;

const createDefaultLocalSettings = () => ({
  theme: 'system',
  scaleLarge: 100,
  scaleVertical: 100,
  selectedMonth: getCurrentMonthKey()
});

const clampPercent = (value, fallback = DEFAULT_PROFIT_SHARE_PERCENT) => {
  const parsed = parseFloat(value);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.max(0, Math.min(100, parsed));
};

const clampScale = (value, fallback = 100) => {
  const parsed = parseInt(value, 10);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.max(20, Math.min(200, parsed));
};

const normalizeLocalSettings = (settings = {}) => {
  const theme = typeof settings.theme === 'string' && ['system', 'light', 'dark'].includes(settings.theme)
    ? settings.theme
    : 'system';

  const selectedMonth = typeof settings.selectedMonth === 'string' && /^\d{4}-\d{2}$/.test(settings.selectedMonth)
    ? settings.selectedMonth
    : getCurrentMonthKey();

  return {
    theme,
    selectedMonth,
    scaleLarge: clampScale(settings.scaleLarge, 100),
    scaleVertical: clampScale(settings.scaleVertical, 100)
  };
};

const normalizePayoutDay = (value, fallback = 15) => {
  const parsed = parseInt(value, 10);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.max(1, Math.min(31, parsed));
};

const normalizePayoutEntry = (entry = {}) => {
  const deductedAdvances = Array.isArray(entry.deductedAdvances)
    ? entry.deductedAdvances
        .map(adv => ({
          id: (adv.id || '').toString().trim(),
          date: (adv.date || '').toString().trim(),
          name: (adv.name || '').toString().trim(),
          amount: Math.max(0, parseFloat(adv.amount) || 0),
          paidById: (adv.paidById || '').toString().trim(),
          restoredToCosts: adv.restoredToCosts === true,
          restoredAt: (adv.restoredAt || '').toString().trim()
        }))
        .filter(adv => adv.id)
    : [];
  const advanceRefunds = Array.isArray(entry.advanceRefunds)
    ? entry.advanceRefunds
        .map(ref => ({
          advanceId: (ref.advanceId || '').toString().trim(),
          toPartnerId: (ref.toPartnerId || '').toString().trim(),
          amount: Math.max(0, parseFloat(ref.amount) || 0),
          returned: ref.returned === true,
          returnedAt: (ref.returnedAt || '').toString().trim()
        }))
        .filter(ref => ref.advanceId && ref.toPartnerId)
    : [];
  const salaryRefund = entry.salaryRefund && typeof entry.salaryRefund === 'object'
    ? {
        toPartnerId: (entry.salaryRefund.toPartnerId || '').toString().trim(),
        amount: Math.max(0, parseFloat(entry.salaryRefund.amount) || 0),
        returned: entry.salaryRefund.returned === true,
        returnedAt: (entry.salaryRefund.returnedAt || '').toString().trim()
      }
    : null;

  return {
    id: (entry.id || generateId()).toString().trim(),
    type: ['monthly', 'weekly', 'custom'].includes(entry.type) ? entry.type : 'monthly',
    label: (entry.label || '').toString().trim(),
    sourceMonth: /^\d{4}-\d{2}$/.test((entry.sourceMonth || '').toString().trim())
      ? entry.sourceMonth.toString().trim() : '',
    payoutDate: /^\d{4}-\d{2}-\d{2}$/.test((entry.payoutDate || '').trim())
      ? entry.payoutDate.trim() : '',
    cashAmount: Math.max(0, parseFloat(entry.cashAmount) || 0),
    paidByPartnerId: (entry.paidByPartnerId || '').toString().trim(),
    deductedAdvances,
    advanceRefunds,
    salaryRefund: salaryRefund?.toPartnerId ? salaryRefund : null,
    createdAt: (entry.createdAt || '').toString().trim()
  };
};

const normalizePayoutEmployeeRecord = (record = {}) => {
  const rawPayouts = Array.isArray(record.payouts)
    ? record.payouts.map(normalizePayoutEntry).filter(p => p.id)
    : [];

  // Migration: old-style data (no payouts array) → create synthetic entry
  const payouts = rawPayouts.length === 0 && (
    (parseFloat(record.settledCashAmount) || 0) > 0.005 ||
    (parseFloat(record.settledAdvanceAmount) || 0) > 0.005
  ) ? [normalizePayoutEntry({
    id: 'legacy-migration',
    type: 'monthly',
    label: 'Wypłata (historyczna)',
    payoutDate: (record.lastSettledAt || '').substring(0, 10),
    cashAmount: Math.max(0, parseFloat(record.settledCashAmount) || 0),
    paidByPartnerId: '',
    deductedAdvances: (Array.isArray(record.removedAdvanceExpenseIds)
      ? record.removedAdvanceExpenseIds.filter(Boolean).map(id => ({
          id: id.toString().trim(),
          date: '',
          name: 'Zaliczka (historyczna)',
          amount: 0,
          paidById: '',
          restoredToCosts: false,
          restoredAt: ''
        }))
      : []),
    advanceRefunds: [],
    createdAt: record.lastSettledAt || ''
  })] : rawPayouts;

  // Derive legacy fields from payouts array (for backward compat with calculations)
  const settledCashAmount = payouts.reduce((sum, p) => sum + p.cashAmount, 0);
  const settledAdvanceAmount = payouts.reduce(
    (sum, p) => sum + p.deductedAdvances.filter(a => !a.restoredToCosts).reduce((s, a) => s + a.amount, 0), 0
  );
  const removedAdvanceExpenseIds = [
    ...new Set(payouts.flatMap(p => p.deductedAdvances.filter(a => !a.restoredToCosts).map(a => a.id)))
  ];
  const sorted = [...payouts].sort((a, b) => (a.createdAt || '').localeCompare(b.createdAt || ''));
  const lastSettledAt = sorted.length > 0 ? sorted[sorted.length - 1].createdAt : (record.lastSettledAt || '');

  return {
    includeCarryover: record.includeCarryover !== false,
    deductAdvancesMode: ['none', 'default-day', 'custom-day'].includes(record.deductAdvancesMode)
      ? record.deductAdvancesMode : 'default-day',
    customDeductionDate: /^\d{4}-\d{2}-\d{2}$/.test((record.customDeductionDate || '').toString().trim())
      ? record.customDeductionDate.toString().trim() : '',
    sourceMonth: /^\d{4}-\d{2}$/.test((record.sourceMonth || '').toString().trim())
      ? record.sourceMonth.toString().trim() : '',
    baseAmountSnapshot: Math.max(0, parseFloat(record.baseAmountSnapshot) || 0),
    carryoverAmountSnapshot: Math.max(0, parseFloat(record.carryoverAmountSnapshot) || 0),
    plannedAmountSnapshot: Math.max(0, parseFloat(record.plannedAmountSnapshot) || 0),
    advanceDeductionAmountSnapshot: Math.max(0, parseFloat(record.advanceDeductionAmountSnapshot) || 0),
    includeCurrentMonth: record.includeCurrentMonth === true,
    payouts,
    settledCashAmount,
    settledAdvanceAmount,
    removedAdvanceExpenseIds,
    lastSettledAt
  };
};

const normalizePayoutSettings = (settings = {}) => {
  const employees = Object.entries(settings.employees || {}).reduce((acc, [personId, employeeRecord]) => {
    const normalizedPersonId = (personId || '').toString().trim();
    if (!normalizedPersonId) return acc;
    acc[normalizedPersonId] = normalizePayoutEmployeeRecord(employeeRecord || {});
    return acc;
  }, {});

  const separateCompanies = Object.entries(settings.separateCompanies || {}).reduce((acc, [companyId, config]) => {
    const normalizedId = (companyId || '').toString().trim();
    if (!normalizedId) return acc;
    acc[normalizedId] = { enablePayouts: config?.enablePayouts === true };
    return acc;
  }, {});

  return {
    defaultDay: normalizePayoutDay(settings.defaultDay, 15),
    employees,
    separateCompanies
  };
};

const hasMeaningfulPayoutSettings = (settings = {}) => {
  const normalized = normalizePayoutSettings(settings || {});
  return normalized.defaultDay !== 15 || Object.keys(normalized.employees || {}).length > 0;
};

const hasMeaningfulInvoicesConfig = (config = {}) => {
  const normalized = normalizeInvoicesConfig(config || {});
  return (normalized.issueDate || '').trim() !== ''
    || (normalized.emailIntro || '').trim() !== ''
    || Object.keys(normalized.clients || {}).length > 0
    || (normalized.extraInvoices || []).length > 0;
};

const buildRemoteMonthSettingsPayload = (monthSettings = {}, options = {}) => {
  const { includeCommonSnapshot = true } = options;
  const payload = {};

  if (monthSettings?.persons && Object.keys(monthSettings.persons).length > 0) {
    payload.persons = cloneData(monthSettings.persons);
  }

  if (monthSettings?.clients && Object.keys(monthSettings.clients).length > 0) {
    payload.clients = cloneData(monthSettings.clients);
  }

  if (monthSettings?.settlementConfig && Object.keys(monthSettings.settlementConfig).length > 0) {
    payload.settlementConfig = cloneData(monthSettings.settlementConfig);
  }

  if (monthSettings?.personContractCharges && Object.keys(monthSettings.personContractCharges).length > 0) {
    payload.personContractCharges = cloneData(monthSettings.personContractCharges);
  }

  if (hasMeaningfulPayoutSettings(monthSettings?.payouts || {})) {
    payload.payouts = normalizePayoutSettings(monthSettings.payouts || {});
  }

  if (hasMeaningfulInvoicesConfig(monthSettings?.invoices || {})) {
    payload.invoices = normalizeInvoicesConfig(monthSettings.invoices || {});
  }

  if (monthSettings?.isArchived === true) {
    payload.isArchived = true;
  }

  if (includeCommonSnapshot && monthSettings?.commonSnapshot) {
    payload.commonSnapshot = normalizeCommonData(monthSettings.commonSnapshot, createDefaultState().common);
  }

  return Object.keys(payload).length > 0 ? payload : null;
};

const serializeRemoteMonthRecord = (monthRecord = {}) => {
  const normalizedMonth = normalizeMonth(monthRecord || {});
  const monthSettingsPayload = buildRemoteMonthSettingsPayload(normalizedMonth.monthSettings || {}, { includeCommonSnapshot: true });

  return {
    monthlySheets: serializeRemoteCollectionMap(normalizedMonth.monthlySheets || []),
    worksSheets: serializeRemoteCollectionMap(normalizedMonth.worksSheets || []),
    expenses: serializeRemoteCollectionMap(normalizedMonth.expenses || []),
    monthSettings: monthSettingsPayload
  };
};

const buildRemoteCommonPayload = (common = {}) => normalizeCommonData(common, createDefaultState().common);

const getRemoteMonthScopeData = (monthRecord = {}, syncScopeKey = '') => {
  const normalizedMonth = normalizeMonth(monthRecord || {});

  switch (syncScopeKey) {
    case 'monthlySheets':
      return serializeRemoteCollectionMap(normalizedMonth.monthlySheets || []);
    case 'worksSheets':
      return serializeRemoteCollectionMap(normalizedMonth.worksSheets || []);
    case 'expenses':
      return serializeRemoteCollectionMap(normalizedMonth.expenses || []);
    case 'monthSettings':
      return buildRemoteMonthSettingsPayload(normalizedMonth.monthSettings || {}, { includeCommonSnapshot: false });
    case 'monthSettings.commonSnapshot':
      return normalizedMonth.monthSettings?.commonSnapshot
        ? normalizeCommonData(normalizedMonth.monthSettings.commonSnapshot, createDefaultState().common)
        : null;
    default:
      return null;
  }
};

const hasRemoteMonthScopeData = (monthRecord = {}, syncScopeKey = '') => getRemoteMonthScopeData(monthRecord, syncScopeKey) !== null;

const normalizeClientRecord = (client = {}, existingClient = {}) => {
  const normalizeString = (value, fallback = '') => (value ?? fallback).toString().trim();
  const hourlyRate = parseFloat(client.hourlyRate ?? existingClient.hourlyRate);

  return {
    ...existingClient,
    ...client,
    id: client.id || existingClient.id || generateId(),
    name: normalizeString(client.name, existingClient.name),
    hourlyRate: Number.isFinite(hourlyRate) ? hourlyRate : 0,
    customWorkPrices: client.customWorkPrices ?? existingClient.customWorkPrices ?? {},
    fullCompanyName: normalizeString(client.fullCompanyName, existingClient.fullCompanyName),
    address: normalizeString(client.address, existingClient.address),
    nip: normalizeString(client.nip, existingClient.nip).replace(/\s+/g, ''),
    krs: normalizeString(client.krs, existingClient.krs).replace(/\s+/g, ''),
    regon: normalizeString(client.regon, existingClient.regon).replace(/\s+/g, ''),
    accountNumbers: Array.isArray(client.accountNumbers)
      ? client.accountNumbers.filter(value => !!value).map(value => value.toString().trim())
      : (Array.isArray(existingClient.accountNumbers) ? existingClient.accountNumbers : []),
    fetchedFromRegistryAt: normalizeString(client.fetchedFromRegistryAt, existingClient.fetchedFromRegistryAt),
    registrySourceUrl: normalizeString(client.registrySourceUrl, existingClient.registrySourceUrl)
  };
};

const getActivePersonIdsForMonthFromState = (state, month = '') => {
  const monthKey = month || Store.getSelectedMonth() || getCurrentMonthKey();
  const m = state?.months?.[monthKey] || {};
  const monthPersonStatuses = m.monthSettings?.persons || {};

  const persons = m.monthSettings?.commonSnapshot?.persons || state?.common?.persons || state?.persons || [];

  return persons
    .filter(person => {
      if (monthPersonStatuses[person.id] !== undefined) {
        return monthPersonStatuses[person.id] !== false;
      }
      return person.isActive !== false;
    })
    .map(person => person.id);
};

const normalizeInvoiceClientConfig = (config = {}) => {
  const issuerIds = Array.isArray(config.issuerIds)
    ? [...new Set(config.issuerIds.filter(value => typeof value === 'string' && value.trim() !== ''))]
    : [];
  const mode = typeof config.mode === 'string' && config.mode.trim() !== '' ? config.mode : 'SETTLEMENT_REVENUE';
  const notes = (config.notes || '').toString().trim();
  const deductClientAdvances = config.deductClientAdvances !== false;
  const includeClientCosts = config.includeClientCosts !== false;
  const randomizeEqualSplitInvoices = config.randomizeEqualSplitInvoices === true;
  const equalSplitVarianceAmount = Math.max(0, parseFloat(config.equalSplitVarianceAmount) || 10);
  const percentageAllocations = {};
  const percentageTouchedIssuerIds = Array.isArray(config.percentageTouchedIssuerIds)
    ? [...new Set(config.percentageTouchedIssuerIds.filter(value => typeof value === 'string' && value.trim() !== ''))]
    : [];
  const lastEditedPercentageIssuerId = typeof config.lastEditedPercentageIssuerId === 'string' && config.lastEditedPercentageIssuerId.trim() !== ''
    ? config.lastEditedPercentageIssuerId
    : '';
  const manualAmounts = {};
  const separateCompanyWithEmployees = {};

  Object.entries(config.percentageAllocations || {}).forEach(([personId, value]) => {
    const amount = Math.max(0, Math.min(100, Math.round(parseFloat(value))));
    if (Number.isFinite(amount)) {
      percentageAllocations[personId] = amount;
    }
  });

  Object.entries(config.manualAmounts || {}).forEach(([personId, value]) => {
    const amount = parseFloat(value);
    if (Number.isFinite(amount)) {
      manualAmounts[personId] = amount;
    }
  });

  Object.entries(config.separateCompanyWithEmployees || {}).forEach(([personId, value]) => {
    separateCompanyWithEmployees[personId] = value === true;
  });

  return {
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
    notes
  };
};

const normalizeExtraInvoiceRecord = (invoice = {}, existingInvoice = {}) => {
  const amount = parseFloat(invoice.amount ?? existingInvoice.amount);
  const normalizedClientId = (invoice.clientId ?? existingInvoice.clientId ?? '').toString().trim();
  const normalizedClientName = (invoice.clientName ?? existingInvoice.clientName ?? '').toString().trim();
  const normalizedIssuerId = (invoice.issuerId ?? existingInvoice.issuerId ?? '').toString().trim();

  return {
    ...existingInvoice,
    ...invoice,
    id: (invoice.id ?? existingInvoice.id ?? generateId()).toString(),
    clientId: normalizedClientId,
    clientName: normalizedClientName,
    issuerId: normalizedIssuerId,
    amount: Number.isFinite(amount) ? amount : 0
  };
};

const normalizeIssuedInvoiceSnapshot = (snapshot = {}, fallbackMonth = '') => {
  if (!snapshot || typeof snapshot !== 'object' || Array.isArray(snapshot)) return null;

  const normalizeAmount = (value, fallback = 0) => {
    const amount = parseFloat(value);
    return Number.isFinite(amount) ? amount : fallback;
  };

  return {
    month: (snapshot.month || fallbackMonth || '').toString().trim(),
    issueDate: (snapshot.issueDate || '').toString().trim(),
    issuers: Array.isArray(snapshot.issuers)
      ? snapshot.issuers
        .filter(issuer => issuer && typeof issuer === 'object' && issuer.id)
        .map(issuer => ({ ...cloneData(issuer) }))
      : [],
    clientInvoices: Array.isArray(snapshot.clientInvoices)
      ? snapshot.clientInvoices
        .filter(invoice => invoice && typeof invoice === 'object')
        .map(invoice => ({ ...cloneData(invoice) }))
      : [],
    extraInvoices: Array.isArray(snapshot.extraInvoices)
      ? snapshot.extraInvoices
        .filter(invoice => invoice && typeof invoice === 'object')
        .map(invoice => ({ ...cloneData(invoice), amount: normalizeAmount(invoice.amount) }))
      : [],
    issuerSummaries: Array.isArray(snapshot.issuerSummaries)
      ? snapshot.issuerSummaries
        .filter(summary => summary && typeof summary === 'object' && summary.issuerId)
        .map(summary => ({
          ...cloneData(summary),
          totalAmount: normalizeAmount(summary.totalAmount),
          settlementAmount: normalizeAmount(summary.settlementAmount),
          extraInvoicesAmount: normalizeAmount(summary.extraInvoicesAmount),
          taxRate: normalizeAmount(summary.taxRate),
          taxAmount: normalizeAmount(summary.taxAmount),
          yearToDateTotal: summary.yearToDateTotal === null ? null : normalizeAmount(summary.yearToDateTotal)
        }))
      : [],
    totalRevenue: normalizeAmount(snapshot.totalRevenue),
    totalInvoices: normalizeAmount(snapshot.totalInvoices),
    difference: normalizeAmount(snapshot.difference, normalizeAmount(snapshot.totalInvoices) - normalizeAmount(snapshot.totalRevenue)),
    emailText: (snapshot.emailText || '').toString().trim()
  };
};

const normalizeInvoicesConfig = (config = {}) => {
  const clients = {};
  Object.entries(config.clients || {}).forEach(([clientId, clientConfig]) => {
    clients[clientId] = normalizeInvoiceClientConfig(clientConfig);
  });

  const issuedSnapshot = normalizeIssuedInvoiceSnapshot(config.issuedSnapshot, config.issueDate || '');

  return {
    issueDate: (config.issueDate || '').toString().trim(),
    emailIntro: (config.emailIntro || '').toString().trim(),
    issued: config.issued === true,
    issuedAt: (config.issuedAt || '').toString().trim(),
    issuedSnapshot,
    clients,
    extraInvoices: Array.isArray(config.extraInvoices)
      ? config.extraInvoices.map(invoice => normalizeExtraInvoiceRecord(invoice)).filter(invoice => invoice.issuerId && invoice.amount > 0 && (invoice.clientId || invoice.clientName))
      : []
  };
};

const createDefaultState = () => ({
  version: 'v3',
  common: {
    persons: [],
    clients: [],
    worksCatalog: DEFAULT_WORKS_CATALOG.map(w => ({ ...w, id: generateId() })),
    config: {
      taxRate: 0.055,
      zusFixedAmount: 1600.27
    }
  },
  months: {} 
});

const cloneData = (value) => JSON.parse(JSON.stringify(value));

const normalizeRemoteCollectionArray = (value = null) => {
  if (Array.isArray(value)) return value;
  if (!value || typeof value !== 'object') return [];

  return Object.values(value)
    .filter(item => item && typeof item === 'object')
    .sort((left, right) => {
      const leftOrder = Number.isFinite(parseInt(left?.sortOrder, 10)) ? parseInt(left.sortOrder, 10) : Number.MAX_SAFE_INTEGER;
      const rightOrder = Number.isFinite(parseInt(right?.sortOrder, 10)) ? parseInt(right.sortOrder, 10) : Number.MAX_SAFE_INTEGER;
      return leftOrder - rightOrder;
    })
    .map(item => {
      const normalizedItem = cloneData(item);
      delete normalizedItem.sortOrder;
      return normalizedItem;
    });
};

const serializeRemoteCollectionMap = (items = []) => {
  const normalizedItems = Array.isArray(items)
    ? items.filter(item => item && typeof item === 'object' && item.id)
    : [];

  if (normalizedItems.length === 0) return null;

  return normalizedItems.reduce((acc, item, index) => {
    acc[item.id] = {
      ...cloneData(item),
      sortOrder: index
    };
    return acc;
  }, {});
};

const serializeRemoteCollectionItem = (item = {}, sortOrder = 0) => {
  if (!item || typeof item !== 'object' || !item.id) return null;
  return {
    ...cloneData(item),
    sortOrder
  };
};

const buildCanonicalComparableValue = (value = null) => {
  if (Array.isArray(value)) {
    return value.map(item => buildCanonicalComparableValue(item));
  }

  if (!value || typeof value !== 'object') {
    return value;
  }

  return Object.keys(value)
    .sort((leftKey, rightKey) => leftKey.localeCompare(rightKey, 'pl-PL'))
    .reduce((acc, key) => {
      acc[key] = buildCanonicalComparableValue(value[key]);
      return acc;
    }, {});
};

const deepEqual = (left, right) => JSON.stringify(buildCanonicalComparableValue(left)) === JSON.stringify(buildCanonicalComparableValue(right));

const getIsoTimestamp = () => new Date().toISOString();

const getTodayKey = () => getIsoTimestamp().slice(0, 10);

const trimLimitedEntries = (entries, limit) => entries.slice(0, Math.max(0, limit));

const createDefaultSyncMeta = () => ({
  common: {},
  months: {}
});

const createDefaultFirebaseTransferStats = () => ({
  downloadBytes: 0,
  uploadBytes: 0,
  readCount: 0,
  writeCount: 0,
  lastReadAt: '',
  lastWriteAt: ''
});

const estimateSerializedBytes = (value = null) => {
  try {
    return new Blob([JSON.stringify(value ?? null)]).size;
  } catch {
    return 0;
  }
};

const normalizeSyncMetaEntry = (entry = {}) => ({
  revision: Math.max(0, parseInt(entry.revision, 10) || 0),
  updatedAt: (entry.updatedAt || '').toString().trim()
});

const isSyncMetaLeafEntry = (value) => {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  return Object.prototype.hasOwnProperty.call(value, 'revision')
    || Object.prototype.hasOwnProperty.call(value, 'updatedAt');
};

const getSyncScopeKey = (scopeKey = '') => {
  const normalizedScopeKey = (scopeKey || '').toString().trim();
  if (!normalizedScopeKey) return '';

  if (normalizedScopeKey === 'common' || normalizedScopeKey.startsWith('common.')) {
    return 'common';
  }

  const monthScopeKey = normalizedScopeKey.replace(/^month\./, '');
  if (monthScopeKey === 'monthlySheets' || monthScopeKey.startsWith('monthlySheets.')) return 'monthlySheets';
  if (monthScopeKey === 'worksSheets' || monthScopeKey.startsWith('worksSheets.')) return 'worksSheets';
  if (monthScopeKey === 'expenses' || monthScopeKey.startsWith('expenses.')) return 'expenses';
  if (monthScopeKey === 'monthSettings.commonSnapshot' || monthScopeKey.startsWith('monthSettings.commonSnapshot.')) return 'monthSettings.commonSnapshot';
  return 'monthSettings';
};

const getSyncMetaMonthEntry = (meta = {}, monthKey = '') => {
  if (!monthKey) return {};
  return meta?.months?.[monthKey] && typeof meta.months[monthKey] === 'object' && !Array.isArray(meta.months[monthKey])
    ? meta.months[monthKey]
    : {};
};

const normalizeSyncMetadata = (meta = {}) => {
  const normalized = createDefaultSyncMeta();

  if (isSyncMetaLeafEntry(meta.common)) {
    normalized.common = normalizeSyncMetaEntry(meta.common);
  }

  Object.entries(meta.months || {}).forEach(([monthKey, rawMonthMeta]) => {
    if (!/^\d{4}-\d{2}$/.test(monthKey) || !rawMonthMeta || typeof rawMonthMeta !== 'object' || Array.isArray(rawMonthMeta)) return;

    const normalizedMonthMeta = {};

    ['monthlySheets', 'worksSheets', 'expenses'].forEach(scopeKey => {
      if (isSyncMetaLeafEntry(rawMonthMeta?.[scopeKey])) {
        normalizedMonthMeta[scopeKey] = normalizeSyncMetaEntry(rawMonthMeta[scopeKey]);
      }
    });

    const rawMonthSettingsMeta = rawMonthMeta?.monthSettings;
    if (rawMonthSettingsMeta && typeof rawMonthSettingsMeta === 'object' && !Array.isArray(rawMonthSettingsMeta)) {
      if (isSyncMetaLeafEntry(rawMonthSettingsMeta)) {
        normalizedMonthMeta.monthSettings = normalizeSyncMetaEntry(rawMonthSettingsMeta);
      }

      if (isSyncMetaLeafEntry(rawMonthSettingsMeta.commonSnapshot)) {
        if (!normalizedMonthMeta.monthSettings) normalizedMonthMeta.monthSettings = {};
        normalizedMonthMeta.monthSettings.commonSnapshot = normalizeSyncMetaEntry(rawMonthSettingsMeta.commonSnapshot);
      }
    }

    if (Object.keys(normalizedMonthMeta).length > 0) {
      normalized.months[monthKey] = normalizedMonthMeta;
    }
  });

  return normalized;
};

const serializeSyncMetadataForRemote = (meta = {}) => {
  const normalized = normalizeSyncMetadata(meta);
  const serialized = createDefaultSyncMeta();

  if (hasUsableSyncMetaEntry(normalized.common)) {
    serialized.common = normalizeSyncMetaEntry(normalized.common);
  }

  Object.entries(normalized.months || {}).forEach(([monthKey, monthMeta]) => {
    const serializedMonthMeta = {};

    ['monthlySheets', 'worksSheets', 'expenses'].forEach(scopeKey => {
      if (hasUsableSyncMetaEntry(monthMeta?.[scopeKey])) {
        serializedMonthMeta[scopeKey] = normalizeSyncMetaEntry(monthMeta[scopeKey]);
      }
    });

    const serializedMonthSettingsMeta = {};
    if (hasUsableSyncMetaEntry(monthMeta?.monthSettings)) {
      Object.assign(serializedMonthSettingsMeta, normalizeSyncMetaEntry(monthMeta.monthSettings));
    }
    if (hasUsableSyncMetaEntry(monthMeta?.monthSettings?.commonSnapshot)) {
      serializedMonthSettingsMeta.commonSnapshot = normalizeSyncMetaEntry(monthMeta.monthSettings.commonSnapshot);
    }
    if (Object.keys(serializedMonthSettingsMeta).length > 0) {
      serializedMonthMeta.monthSettings = serializedMonthSettingsMeta;
    }

    if (Object.keys(serializedMonthMeta).length > 0) {
      serialized.months[monthKey] = serializedMonthMeta;
    }
  });

  return serialized;
};

const hasUsableSyncMetaEntry = (entry = {}) => {
  const normalized = normalizeSyncMetaEntry(entry);
  return normalized.revision > 0 && normalized.updatedAt !== '';
};

const hasStateDataForSyncScope = (state = {}, scopeKey = '', monthKey = '') => {
  if (scopeKey === 'common') {
    return true;
  }

  return hasRemoteMonthScopeData(state?.months?.[monthKey] || {}, scopeKey);
};

const pruneMonthSyncMetadata = (monthMeta = {}) => {
  const nextMonthMeta = cloneData(monthMeta || {});

  ['monthlySheets', 'worksSheets', 'expenses'].forEach(scopeKey => {
    if (!hasUsableSyncMetaEntry(nextMonthMeta?.[scopeKey])) {
      delete nextMonthMeta[scopeKey];
    }
  });

  if (nextMonthMeta.monthSettings) {
    if (!hasUsableSyncMetaEntry(nextMonthMeta.monthSettings)) {
      delete nextMonthMeta.monthSettings.revision;
      delete nextMonthMeta.monthSettings.updatedAt;
    }

    if (!hasUsableSyncMetaEntry(nextMonthMeta.monthSettings.commonSnapshot)) {
      delete nextMonthMeta.monthSettings.commonSnapshot;
    }

    if (Object.keys(nextMonthMeta.monthSettings).length === 0) {
      delete nextMonthMeta.monthSettings;
    }
  }

  return nextMonthMeta;
};

const mergeSyncMetadataWithState = (state, baseMeta = createDefaultSyncMeta(), author = getCurrentActorId(), timestamp = getIsoTimestamp()) => {
  const nextMeta = normalizeSyncMetadata(baseMeta || createDefaultSyncMeta());
  let changed = false;

  if (!hasUsableSyncMetaEntry(nextMeta.common)) {
    nextMeta.common = {
      revision: 1,
      updatedAt: timestamp
    };
    changed = true;
  }

  Object.keys(nextMeta.months || {}).forEach(monthKey => {
    const hasAnyStateData = MONTH_SYNC_SCOPE_KEYS.some(scopeKey => hasStateDataForSyncScope(state || {}, scopeKey, monthKey));
    if (!hasAnyStateData && Object.keys(nextMeta.months[monthKey] || {}).length > 0) {
      delete nextMeta.months[monthKey];
      changed = true;
    }
  });

  Object.keys(state?.months || {}).forEach(monthKey => {
    const nextMonthMeta = getSyncMetaMonthEntry(nextMeta, monthKey);

    ['monthlySheets', 'worksSheets', 'expenses'].forEach(scopeKey => {
      if (!hasStateDataForSyncScope(state || {}, scopeKey, monthKey)) {
        if (nextMonthMeta[scopeKey]) {
          delete nextMonthMeta[scopeKey];
          changed = true;
        }
        return;
      }

      if (!hasUsableSyncMetaEntry(nextMonthMeta[scopeKey])) {
        nextMonthMeta[scopeKey] = {
          revision: 1,
          updatedAt: timestamp
        };
        changed = true;
      }
    });

    if (!hasStateDataForSyncScope(state || {}, 'monthSettings', monthKey)) {
      if (nextMonthMeta.monthSettings && hasUsableSyncMetaEntry(nextMonthMeta.monthSettings)) {
        delete nextMonthMeta.monthSettings.revision;
        delete nextMonthMeta.monthSettings.updatedAt;
        changed = true;
      }
    } else if (!hasUsableSyncMetaEntry(nextMonthMeta.monthSettings)) {
      nextMonthMeta.monthSettings = {
        ...(nextMonthMeta.monthSettings || {}),
        revision: 1,
        updatedAt: timestamp
      };
      changed = true;
    }

    if (!hasStateDataForSyncScope(state || {}, 'monthSettings.commonSnapshot', monthKey)) {
      if (nextMonthMeta.monthSettings?.commonSnapshot) {
        delete nextMonthMeta.monthSettings.commonSnapshot;
        changed = true;
      }
    } else if (!hasUsableSyncMetaEntry(nextMonthMeta.monthSettings?.commonSnapshot)) {
      if (!nextMonthMeta.monthSettings) nextMonthMeta.monthSettings = {};
      nextMonthMeta.monthSettings.commonSnapshot = {
        revision: 1,
        updatedAt: timestamp
      };
      changed = true;
    }

    const prunedMonthMeta = pruneMonthSyncMetadata(nextMonthMeta);
    if (Object.keys(prunedMonthMeta).length > 0) {
      nextMeta.months[monthKey] = prunedMonthMeta;
    } else if (nextMeta.months[monthKey]) {
      delete nextMeta.months[monthKey];
      changed = true;
    }
  });

  return { meta: nextMeta, changed };
};

const collectSyncConflictEntries = (localMeta = {}, remoteMeta = {}) => {
  const localNormalized = normalizeSyncMetadata(localMeta);
  const remoteNormalized = normalizeSyncMetadata(remoteMeta);
  const localNewer = [];
  const remoteNewer = [];

  const commonLocalEntry = normalizeSyncMetaEntry(localNormalized.common);
  const commonRemoteEntry = normalizeSyncMetaEntry(remoteNormalized.common);
  if (commonLocalEntry.revision !== commonRemoteEntry.revision) {
    const payload = {
      scope: 'common',
      scopeLabel: getScopeDisplayLabel('common'),
      month: '',
      localEntry: commonLocalEntry,
      remoteEntry: commonRemoteEntry
    };

    if (commonLocalEntry.revision > commonRemoteEntry.revision) {
      localNewer.push(payload);
    } else {
      remoteNewer.push(payload);
    }
  }

  const knownMonths = [...new Set([
    ...Object.keys(localNormalized.months || {}),
    ...Object.keys(remoteNormalized.months || {})
  ])].sort((a, b) => b.localeCompare(a, 'pl-PL'));

  knownMonths.forEach(monthKey => {
    MONTH_SYNC_SCOPE_KEYS.forEach(scopeKey => {
      const localMonthMeta = getSyncMetaMonthEntry(localNormalized, monthKey);
      const remoteMonthMeta = getSyncMetaMonthEntry(remoteNormalized, monthKey);
      const localEntry = scopeKey === 'monthSettings.commonSnapshot'
        ? normalizeSyncMetaEntry(localMonthMeta?.monthSettings?.commonSnapshot)
        : normalizeSyncMetaEntry(scopeKey === 'monthSettings' ? localMonthMeta?.monthSettings : localMonthMeta?.[scopeKey]);
      const remoteEntry = scopeKey === 'monthSettings.commonSnapshot'
        ? normalizeSyncMetaEntry(remoteMonthMeta?.monthSettings?.commonSnapshot)
        : normalizeSyncMetaEntry(scopeKey === 'monthSettings' ? remoteMonthMeta?.monthSettings : remoteMonthMeta?.[scopeKey]);
      const payload = {
        scope: `month.${scopeKey}`,
        scopeLabel: getScopeDisplayLabel(`month.${scopeKey}`),
        month: monthKey,
        localEntry,
        remoteEntry
      };

      if (localEntry.revision > remoteEntry.revision) {
        localNewer.push(payload);
      } else if (remoteEntry.revision > localEntry.revision) {
        remoteNewer.push(payload);
      }
    });
  });

  return { localNewer, remoteNewer };
};

const normalizeFirebaseEntryCollectionInput = (entries = []) => {
  if (Array.isArray(entries)) return entries;
  if (entries && typeof entries === 'object') return Object.values(entries);
  return [];
};

const historyChangeHasLoadedDetails = (change = {}) => Object.prototype.hasOwnProperty.call(change || {}, 'beforeSnapshot')
  || Object.prototype.hasOwnProperty.call(change || {}, 'afterSnapshot');

const historyEntryHasLoadedDetails = (entry = {}) => entry?.detailsLoaded === true
  || (Array.isArray(entry?.changes) && entry.changes.some(historyChangeHasLoadedDetails));

const backupEntryHasLoadedDetails = (entry = {}) => entry?.detailsLoaded === true
  || Object.prototype.hasOwnProperty.call(entry || {}, 'commonSnapshot')
  || Object.prototype.hasOwnProperty.call(entry || {}, 'monthSnapshot')
  || Object.prototype.hasOwnProperty.call(entry || {}, 'stateSnapshot');

const getHistoryEntryCollectionFlags = (scope = '', beforeSnapshot = null, afterSnapshot = null) => {
  const beforeItems = scope === 'month.monthSettings.invoices'
    ? (Array.isArray(beforeSnapshot?.extraInvoices) ? beforeSnapshot.extraInvoices : [])
    : (Array.isArray(beforeSnapshot) ? beforeSnapshot : []);
  const afterItems = scope === 'month.monthSettings.invoices'
    ? (Array.isArray(afterSnapshot?.extraInvoices) ? afterSnapshot.extraInvoices : [])
    : (Array.isArray(afterSnapshot) ? afterSnapshot : []);
  const { addedItems, removedItems } = getCollectionAddRemoveDiff(beforeItems, afterItems);

  return {
    hasAddedItems: addedItems.length > 0,
    hasRemovedItems: removedItems.length > 0
  };
};

const buildSharedHistoryIndexEntry = (entry = {}) => ({
  id: (entry.id || generateId()).toString(),
  timestamp: (entry.timestamp || '').toString(),
  author: (entry.author || '').toString(),
  label: (entry.label || 'Zmiana').toString(),
  month: (entry.month || '').toString(),
  detailsLoaded: false,
  hasDetails: true,
  changes: normalizeFirebaseEntryCollectionInput(entry.changes)
    .filter(change => change && typeof change === 'object')
    .map(change => ({
      scope: (change.scope || '').toString(),
      month: (change.month || '').toString(),
      scopeLabel: (change.scopeLabel || '').toString(),
      beforeRevision: Math.max(0, parseInt(change.beforeRevision, 10) || 0),
      afterRevision: Math.max(0, parseInt(change.afterRevision, 10) || 0),
      ...getHistoryEntryCollectionFlags(change.scope, change.beforeSnapshot, change.afterSnapshot)
    }))
});

const buildBackupIndexEntry = (entry = {}) => ({
  id: (entry.id || generateId()).toString(),
  type: (entry.type || '').toString(),
  timestamp: (entry.timestamp || '').toString(),
  author: (entry.author || '').toString(),
  month: (entry.month || '').toString(),
  dayKey: (entry.dayKey || '').toString(),
  detailsLoaded: false,
  hasDetails: true
});

const prependNormalizedEntriesWithLimit = (currentEntries = [], nextEntry = null, limit = 20, normalizer = (items = []) => items) => {
  const normalizedEntries = normalizer([nextEntry, ...currentEntries]);
  return {
    entries: normalizedEntries.slice(0, limit),
    removedEntries: normalizedEntries.slice(limit)
  };
};

const buildIndexedRemoteUpdatesForEntries = (indexRoot = '', entriesRoot = '', nextEntries = [], removedEntries = [], indexBuilder = (entry) => entry) => {
  const updates = {};

  nextEntries.forEach(entry => {
    if (!entry?.id) return;
    updates[`/${indexRoot}/${entry.id}`] = indexBuilder(entry);
    if (historyEntryHasLoadedDetails(entry) || backupEntryHasLoadedDetails(entry)) {
      updates[`/${entriesRoot}/${entry.id}`] = cloneData(entry);
    }
  });

  removedEntries.forEach(entry => {
    if (!entry?.id) return;
    updates[`/${indexRoot}/${entry.id}`] = null;
    updates[`/${entriesRoot}/${entry.id}`] = null;
  });

  return updates;
};

const buildIndexedRemoteUpsertUpdatesFromCollection = (entries = [], indexRoot = '', entriesRoot = '', indexBuilder = (entry) => entry, normalizer = (items = []) => items) => {
  return normalizer(entries).reduce((updates, entry) => ({
    ...updates,
    ...buildIndexedRemoteUpdatesForEntries(indexRoot, entriesRoot, [entry], [], indexBuilder)
  }), {});
};

const normalizeSharedHistoryEntries = (entries = []) => normalizeFirebaseEntryCollectionInput(entries)
  .filter(entry => entry && typeof entry === 'object')
  .map(entry => {
    const normalizedChanges = normalizeFirebaseEntryCollectionInput(entry.changes)
      .filter(change => change && typeof change === 'object')
      .map(change => {
        const normalizedChange = {
          ...change,
          scope: (change.scope || '').toString(),
          month: (change.month || '').toString(),
          scopeLabel: (change.scopeLabel || '').toString(),
          beforeRevision: Math.max(0, parseInt(change.beforeRevision, 10) || 0),
          afterRevision: Math.max(0, parseInt(change.afterRevision, 10) || 0),
          hasAddedItems: change.hasAddedItems === true,
          hasRemovedItems: change.hasRemovedItems === true
        };

        if (Object.prototype.hasOwnProperty.call(change, 'beforeSnapshot')) {
          normalizedChange.beforeSnapshot = cloneData(change.beforeSnapshot ?? null);
        }
        if (Object.prototype.hasOwnProperty.call(change, 'afterSnapshot')) {
          normalizedChange.afterSnapshot = cloneData(change.afterSnapshot ?? null);
        }

        return normalizedChange;
      });

    return {
      ...entry,
      id: (entry.id || generateId()).toString(),
      timestamp: (entry.timestamp || '').toString(),
      author: (entry.author || '').toString(),
      label: (entry.label || 'Zmiana').toString(),
      month: (entry.month || '').toString(),
      detailsLoaded: historyEntryHasLoadedDetails({ ...entry, changes: normalizedChanges }),
      hasDetails: entry.hasDetails !== false,
      changes: normalizedChanges
    };
  })
  .sort((a, b) => (b.timestamp || '').localeCompare(a.timestamp || ''));

const limitSharedHistoryEntries = (entries = []) => {
  const normalizedEntries = normalizeSharedHistoryEntries(entries);
  return {
    entries: normalizedEntries.slice(0, SHARED_HISTORY_LIMIT),
    removedEntries: normalizedEntries.slice(SHARED_HISTORY_LIMIT)
  };
};

const isFirebasePermissionDeniedError = (error) => {
  const code = (error?.code || '').toString().toUpperCase();
  const message = (error?.message || '').toString().toLowerCase();
  return code === 'PERMISSION_DENIED' || message.includes('permission_denied');
};

const buildSharedDataOnlyUpdates = (updates = {}) => Object.fromEntries(
  Object.entries(updates).filter(([path]) => path.startsWith('/shared_data/'))
);

const normalizeBackupEntries = (entries = []) => normalizeFirebaseEntryCollectionInput(entries)
  .filter(entry => entry && typeof entry === 'object')
  .map(entry => {
    const normalizedEntry = {
      ...entry,
      id: (entry.id || generateId()).toString(),
      type: (entry.type || '').toString(),
      timestamp: (entry.timestamp || '').toString(),
      author: (entry.author || '').toString(),
      month: (entry.month || '').toString(),
      dayKey: (entry.dayKey || '').toString(),
      detailsLoaded: backupEntryHasLoadedDetails(entry),
      hasDetails: entry.hasDetails !== false
    };

    if (Object.prototype.hasOwnProperty.call(entry, 'commonSnapshot')) {
      normalizedEntry.commonSnapshot = entry.commonSnapshot ? cloneData(entry.commonSnapshot) : null;
    }
    if (Object.prototype.hasOwnProperty.call(entry, 'monthSnapshot')) {
      normalizedEntry.monthSnapshot = entry.monthSnapshot ? cloneData(entry.monthSnapshot) : null;
    }
    if (Object.prototype.hasOwnProperty.call(entry, 'stateSnapshot')) {
      normalizedEntry.stateSnapshot = entry.stateSnapshot ? cloneData(entry.stateSnapshot) : null;
    }

    return normalizedEntry;
  })
  .sort((a, b) => (b.timestamp || '').localeCompare(a.timestamp || ''));

const buildMonthlySnapshotRestoreChanges = (snapshotState = {}) => {
  const normalizedSnapshotState = normalizeState(snapshotState || {});
  const selectedMonth = Store.getSelectedMonth() || getCurrentMonthKey();
  const knownMonths = [...new Set([
    ...Object.keys(appState?.months || {}),
    ...Object.keys(normalizedSnapshotState?.months || {})
  ])];

  return [
    ...COMMON_SCOPE_KEYS.map(scopeKey => ({ scope: `common.${scopeKey}`, month: selectedMonth })),
    ...knownMonths.flatMap(monthKey => MONTH_SCOPE_KEYS.map(scopeKey => ({ scope: `month.${scopeKey}`, month: monthKey })))
  ];
};

const isCommonScope = (scopeKey = '') => scopeKey.startsWith('common.');
const isMonthScope = (scopeKey = '') => scopeKey.startsWith('month.');
const getMetaScopeKey = (scopeKey = '') => getSyncScopeKey(scopeKey);

const getCurrentActorId = () => {
  try {
    if (typeof firebase !== 'undefined' && Array.isArray(firebase.apps) && firebase.apps.length > 0 && firebase.auth) {
      const currentUserEmail = firebase.auth().currentUser?.email;
      if (currentUserEmail) return currentUserEmail;
    }
  } catch {
    // Firebase może nie być jeszcze zainicjalizowany podczas ładowania store.js.
  }
  if (window.isOfflineMode) return 'offline@local';
  return 'local@browser';
};

const getScopeDisplayLabel = (scopeKey = '') => {
  const map = {
    'common': 'Dane wspólne',
    'common.persons': 'Osoby',
    'common.clients': 'Klienci',
    'common.worksCatalog': 'Katalog prac',
    'common.config': 'Konfiguracja wspólna',
    'month.monthlySheets': 'Arkusze godzin',
    'month.worksSheets': 'Arkusze prac',
    'month.expenses': 'Koszty i zaliczki',
    'month.monthSettings': 'Ustawienia miesiąca',
    'month.monthSettings.commonSnapshot': 'Snapshot miesiąca',
    'month.monthSettings.persons': 'Status osób w miesiącu',
    'month.monthSettings.clients': 'Status klientów w miesiącu',
    'month.monthSettings.settlementConfig': 'Konfiguracja rozliczenia miesiąca',
    'month.monthSettings.personContractCharges': 'Podatek i ZUS UZ miesiąca',
    'month.monthSettings.payouts': 'Wypłaty miesiąca',
    'month.monthSettings.invoices': 'Faktury miesiąca',
    'month.monthSettings.archive': 'Archiwizacja miesiąca',
    'month.monthSettings.commonSnapshot.persons': 'Snapshot osób miesiąca',
    'month.monthSettings.commonSnapshot.clients': 'Snapshot klientów miesiąca',
    'month.monthSettings.commonSnapshot.worksCatalog': 'Snapshot katalogu prac miesiąca',
    'month.monthSettings.commonSnapshot.config': 'Snapshot konfiguracji miesiąca'
  };

  return map[scopeKey] || scopeKey;
};

const getSyncMetaEntry = (scopeKey, month = '') => {
  const metaScopeKey = getSyncScopeKey(scopeKey);
  if (metaScopeKey === 'common') {
    return normalizeSyncMetaEntry(syncMetadata.common);
  }
  const monthKey = month || Store.getSelectedMonth() || getCurrentMonthKey();
  const monthMeta = getSyncMetaMonthEntry(syncMetadata, monthKey);

  if (metaScopeKey === 'monthSettings.commonSnapshot') {
    return normalizeSyncMetaEntry(monthMeta?.monthSettings?.commonSnapshot);
  }

  if (metaScopeKey === 'monthSettings') {
    return normalizeSyncMetaEntry(monthMeta?.monthSettings);
  }

  return normalizeSyncMetaEntry(monthMeta?.[metaScopeKey]);
};

const setSyncMetaEntry = (scopeKey, month = '', entry = {}) => {
  const metaScopeKey = getSyncScopeKey(scopeKey);
  if (metaScopeKey === 'common') {
    syncMetadata.common = hasUsableSyncMetaEntry(entry) ? normalizeSyncMetaEntry(entry) : {};
    return;
  }

  const monthKey = month || Store.getSelectedMonth() || getCurrentMonthKey();
  if (!syncMetadata.months[monthKey]) syncMetadata.months[monthKey] = {};
  if (metaScopeKey === 'monthSettings.commonSnapshot') {
    if (!syncMetadata.months[monthKey].monthSettings) syncMetadata.months[monthKey].monthSettings = {};
    if (hasUsableSyncMetaEntry(entry)) {
      syncMetadata.months[monthKey].monthSettings.commonSnapshot = normalizeSyncMetaEntry(entry);
    } else {
      delete syncMetadata.months[monthKey].monthSettings?.commonSnapshot;
    }
  } else if (metaScopeKey === 'monthSettings') {
    if (hasUsableSyncMetaEntry(entry)) {
      syncMetadata.months[monthKey].monthSettings = {
        ...(syncMetadata.months[monthKey].monthSettings || {}),
        ...normalizeSyncMetaEntry(entry)
      };
    } else if (syncMetadata.months[monthKey].monthSettings) {
      delete syncMetadata.months[monthKey].monthSettings.revision;
      delete syncMetadata.months[monthKey].monthSettings.updatedAt;
    }
  } else if (hasUsableSyncMetaEntry(entry)) {
    syncMetadata.months[monthKey][metaScopeKey] = normalizeSyncMetaEntry(entry);
  } else {
    delete syncMetadata.months[monthKey][metaScopeKey];
  }

  syncMetadata.months[monthKey] = pruneMonthSyncMetadata(syncMetadata.months[monthKey]);
  if (Object.keys(syncMetadata.months[monthKey] || {}).length === 0) {
    delete syncMetadata.months[monthKey];
  }
};

const persistSyncMetadataToLocalStorage = () => {
  localStorage.setItem(WORK_TRACKER_SYNC_META_KEY, JSON.stringify(syncMetadata));
};

const dispatchFirebaseTransferStatsChanged = () => {
  window.dispatchEvent(new Event('firebaseTransferStatsChanged'));
};

const recordFirebaseDownload = (payload = null) => {
  firebaseTransferStats.downloadBytes += estimateSerializedBytes(payload);
  firebaseTransferStats.readCount += 1;
  firebaseTransferStats.lastReadAt = getIsoTimestamp();
  dispatchFirebaseTransferStatsChanged();
};

const recordFirebaseUpload = (payload = null) => {
  firebaseTransferStats.uploadBytes += estimateSerializedBytes(payload);
  firebaseTransferStats.writeCount += 1;
  firebaseTransferStats.lastWriteAt = getIsoTimestamp();
  dispatchFirebaseTransferStatsChanged();
};

const trackFirebaseSnapshotValue = (snapshot) => {
  if (!snapshot || typeof snapshot.val !== 'function') return snapshot;
  const value = snapshot.val();
  recordFirebaseDownload(value);
  return {
    val: () => value
  };
};

const getMonthMetaSnapshot = (month = '') => {
  const monthKey = month || Store.getSelectedMonth() || getCurrentMonthKey();
  return cloneData(syncMetadata.months[monthKey] || {});
};

const getRemoteMetaScopeRevision = (remoteMeta = {}, scopeKey = '', month = '') => {
  const metaScopeKey = getSyncScopeKey(scopeKey);
  if (metaScopeKey === 'common') {
    return normalizeSyncMetaEntry(remoteMeta?.common).revision;
  }
  const monthKey = month || Store.getSelectedMonth() || getCurrentMonthKey();
  const monthMeta = getSyncMetaMonthEntry(normalizeSyncMetadata(remoteMeta), monthKey);

  if (metaScopeKey === 'monthSettings.commonSnapshot') {
    return normalizeSyncMetaEntry(monthMeta?.monthSettings?.commonSnapshot).revision;
  }

  if (metaScopeKey === 'monthSettings') {
    return normalizeSyncMetaEntry(monthMeta?.monthSettings).revision;
  }

  return normalizeSyncMetaEntry(monthMeta?.[metaScopeKey]).revision;
};

const isRemoteMetaNewerForScope = (remoteMeta = {}, scopeKey = '', month = '') => {
  const localEntry = getSyncMetaEntry(scopeKey, month);
  const remoteNormalizedMeta = normalizeSyncMetadata(remoteMeta);
  const toComparableTimestamp = (value = '') => {
    const timestamp = Date.parse((value || '').toString().trim());
    return Number.isFinite(timestamp) ? timestamp : 0;
  };
  const remoteEntry = normalizeSyncMetaEntry(
    getSyncScopeKey(scopeKey) === 'common'
      ? remoteNormalizedMeta.common
      : (getSyncScopeKey(scopeKey) === 'monthSettings.commonSnapshot'
          ? getSyncMetaMonthEntry(remoteNormalizedMeta, month || Store.getSelectedMonth() || getCurrentMonthKey())?.monthSettings?.commonSnapshot
          : (getSyncScopeKey(scopeKey) === 'monthSettings'
              ? getSyncMetaMonthEntry(remoteNormalizedMeta, month || Store.getSelectedMonth() || getCurrentMonthKey())?.monthSettings
              : getSyncMetaMonthEntry(remoteNormalizedMeta, month || Store.getSelectedMonth() || getCurrentMonthKey())?.[getSyncScopeKey(scopeKey)]))
  );

  if (remoteEntry.revision > localEntry.revision) return true;
  if (remoteEntry.revision < localEntry.revision) return false;
  return toComparableTimestamp(remoteEntry.updatedAt) > toComparableTimestamp(localEntry.updatedAt);
};

const shouldFetchRemoteCommon = (remoteMeta = {}) => isRemoteMetaNewerForScope(remoteMeta, 'common');

const getRemoteMonthSyncScopesToFetch = (month = '', remoteMeta = {}) => MONTH_SYNC_SCOPE_KEYS
  .filter(scopeKey => isRemoteMetaNewerForScope(remoteMeta, `month.${scopeKey}`, month));

const shouldFetchRemoteMonth = (month = '', remoteMeta = {}) => getRemoteMonthSyncScopesToFetch(month, remoteMeta).length > 0;

const createHistoryBackupPayload = (type, month, commonSnapshot, monthSnapshot, stateSnapshot = null) => {
  const storesFullStateSnapshot = type === 'monthly' && !!stateSnapshot;

  return {
    id: generateId(),
    type,
    month,
    dayKey: type === 'daily' ? getTodayKey() : '',
    timestamp: getIsoTimestamp(),
    author: getCurrentActorId(),
    commonSnapshot: storesFullStateSnapshot ? null : commonSnapshot,
    monthSnapshot: storesFullStateSnapshot ? null : monthSnapshot,
    stateSnapshot
  };
};

const rebuildSyncMetadataFromState = (state, author = getCurrentActorId(), timestamp = getIsoTimestamp()) => {
  const nextMeta = createDefaultSyncMeta();

  nextMeta.common = {
    revision: 1,
    updatedAt: timestamp
  };

  Object.keys(state?.months || {}).forEach(monthKey => {
    const nextMonthMeta = {};

    ['monthlySheets', 'worksSheets', 'expenses'].forEach(scopeKey => {
      if (hasStateDataForSyncScope(state, scopeKey, monthKey)) {
        nextMonthMeta[scopeKey] = {
          revision: 1,
          updatedAt: timestamp
        };
      }
    });

    if (hasStateDataForSyncScope(state, 'monthSettings', monthKey)) {
      nextMonthMeta.monthSettings = {
        revision: 1,
        updatedAt: timestamp
      };
    }

    if (hasStateDataForSyncScope(state, 'monthSettings.commonSnapshot', monthKey)) {
      if (!nextMonthMeta.monthSettings) nextMonthMeta.monthSettings = {};
      nextMonthMeta.monthSettings.commonSnapshot = {
        revision: 1,
        updatedAt: timestamp
      };
    }

    if (Object.keys(pruneMonthSyncMetadata(nextMonthMeta)).length > 0) {
      nextMeta.months[monthKey] = pruneMonthSyncMetadata(nextMonthMeta);
    }
  });

  return nextMeta;
};

const ensureArchivedMonthHasSnapshot = (monthRecord, fallbackCommon) => {
  if (!monthRecord || typeof monthRecord !== 'object') return monthRecord;
  if (!monthRecord.monthSettings) monthRecord.monthSettings = {};

  if (monthRecord.monthSettings.isArchived === true && !monthRecord.monthSettings.commonSnapshot && fallbackCommon) {
    monthRecord.monthSettings.commonSnapshot = cloneData(fallbackCommon);
  }

  return monthRecord;
};

const normalizeCommonData = (common = {}, baseCommon = createDefaultState().common) => ({
  persons: Array.isArray(common.persons) ? common.persons.map(p => normalizePersonRecord(p)) : cloneData(baseCommon.persons || []),
  clients: Array.isArray(common.clients) ? common.clients.map(c => normalizeClientRecord(c)) : cloneData(baseCommon.clients || []),
  worksCatalog: Array.isArray(common.worksCatalog) && common.worksCatalog.length > 0 ? common.worksCatalog : cloneData(baseCommon.worksCatalog || []),
  config: { ...baseCommon.config, ...(common.config || {}) }
});

const normalizePersonRecord = (person = {}, existingPerson = {}) => {
  const normalizedType = person.type || existingPerson.type || 'EMPLOYEE';
  const normalizedName = (person.name ?? existingPerson.name ?? '').toString().trim();
  const normalizedLastName = (person.lastName ?? existingPerson.lastName ?? '').toString().trim();
  const canHaveEmployer = normalizedType === 'EMPLOYEE' || normalizedType === 'WORKING_PARTNER';
  const employerId = canHaveEmployer ? ((person.employerId ?? existingPerson.employerId) || null) : null;
  const usesContractCharges = normalizedType === 'EMPLOYEE' || normalizedType === 'WORKING_PARTNER';

  let participatesInCosts = false;
  if (normalizedType === 'PARTNER' || normalizedType === 'SEPARATE_COMPANY') {
    participatesInCosts = (person.participatesInCosts ?? existingPerson.participatesInCosts ?? true) !== false;
  } else if (normalizedType === 'WORKING_PARTNER') {
    participatesInCosts = (person.participatesInCosts ?? existingPerson.participatesInCosts ?? false) === true;
  }

  const sharesEmployeeProfits = normalizedType === 'SEPARATE_COMPANY'
    ? (person.sharesEmployeeProfits ?? existingPerson.sharesEmployeeProfits ?? false) === true
    : false;
  const employeeProfitSharePercent = normalizedType === 'SEPARATE_COMPANY'
    ? clampPercent(person.employeeProfitSharePercent ?? existingPerson.employeeProfitSharePercent, DEFAULT_PROFIT_SHARE_PERCENT)
    : DEFAULT_PROFIT_SHARE_PERCENT;
  const receivesCompanyEmployeeProfits = normalizedType === 'SEPARATE_COMPANY'
    ? (person.receivesCompanyEmployeeProfits ?? existingPerson.receivesCompanyEmployeeProfits ?? false) === true
    : false;
  const companyEmployeeProfitSharePercent = normalizedType === 'SEPARATE_COMPANY'
    ? clampPercent(person.companyEmployeeProfitSharePercent ?? existingPerson.companyEmployeeProfitSharePercent, DEFAULT_PROFIT_SHARE_PERCENT)
    : DEFAULT_PROFIT_SHARE_PERCENT;
  const countsEmployeeAccountingRefund = normalizedType === 'SEPARATE_COMPANY'
    ? (person.countsEmployeeAccountingRefund ?? existingPerson.countsEmployeeAccountingRefund ?? true) === true
    : false;

  const contractTaxAmount = usesContractCharges
    ? parseFloat(person.contractTaxAmount ?? existingPerson.contractTaxAmount ?? DEFAULT_CONTRACT_TAX_AMOUNT) || 0
    : 0;
  const contractZusAmount = usesContractCharges
    ? parseFloat(person.contractZusAmount ?? existingPerson.contractZusAmount ?? DEFAULT_CONTRACT_ZUS_AMOUNT) || 0
    : 0;
  const contractChargesPaidByEmployer = usesContractCharges
    ? (person.contractChargesPaidByEmployer ?? existingPerson.contractChargesPaidByEmployer ?? (normalizedType === 'EMPLOYEE')) === true
    : false;

  return {
    ...existingPerson,
    ...person,
    id: person.id || existingPerson.id || generateId(),
    name: normalizedName,
    lastName: normalizedLastName,
    type: normalizedType,
    employerId,
    participatesInCosts,
    sharesEmployeeProfits,
    employeeProfitSharePercent,
    receivesCompanyEmployeeProfits,
    companyEmployeeProfitSharePercent,
    countsEmployeeAccountingRefund,
    contractTaxAmount,
    contractZusAmount,
    contractChargesPaidByEmployer
  };
};

const normalizeExpenseRecord = (expense = {}, existingExpense = {}) => {
  const type = (expense.type ?? existingExpense.type ?? 'COST').toString().trim() || 'COST';
  const amount = parseFloat(expense.amount ?? existingExpense.amount);
  const rawMode = (expense.dietaCalculationMode ?? existingExpense.dietaCalculationMode ?? '').toString().trim().toUpperCase();
  const dietaCalculationMode = type === 'DIETA'
    ? (rawMode === 'ACTIVE_DAYS' || rawMode === 'MANUAL_DAYS' || rawMode === 'FIXED'
        ? rawMode
        : ((expense.dietaByActiveDays ?? existingExpense.dietaByActiveDays) === true ? 'ACTIVE_DAYS' : 'FIXED'))
    : 'FIXED';

  return {
    ...existingExpense,
    ...expense,
    id: expense.id || existingExpense.id || generateId(),
    type,
    date: (expense.date ?? existingExpense.date ?? '').toString().trim(),
    name: (expense.name ?? existingExpense.name ?? '').toString().trim(),
    amount: Number.isFinite(amount) ? amount : 0,
    paidById: (expense.paidById ?? existingExpense.paidById ?? '').toString().trim(),
    advanceForId: (expense.advanceForId ?? existingExpense.advanceForId ?? '').toString().trim(),
    dietaCalculationMode,
    dietaByActiveDays: type === 'DIETA' && dietaCalculationMode === 'ACTIVE_DAYS',
    dietDaysAdjustment: type === 'DIETA' && dietaCalculationMode === 'ACTIVE_DAYS'
      ? (parseInt(expense.dietDaysAdjustment ?? existingExpense.dietDaysAdjustment, 10) || 0)
      : 0,
    dietaDaysCount: type === 'DIETA' && dietaCalculationMode === 'MANUAL_DAYS'
      ? Math.max(0, parseInt(expense.dietaDaysCount ?? existingExpense.dietaDaysCount, 10) || 0)
      : 0
  };
};

const normalizeMonthDays = (days) => {
  if (!days || typeof days !== 'object') return {};

  return Object.entries(days).reduce((acc, [dayKey, dayValue]) => {
    if (!dayValue || typeof dayValue !== 'object') return acc;
    acc[dayKey] = dayValue;
    return acc;
  }, {});
};

const normalizeMonth = (m) => {
  if (!m || typeof m !== 'object') {
    return {
      monthlySheets: [],
      worksSheets: [],
      expenses: [],
      monthSettings: {
        persons: {},
        clients: {},
        settlementConfig: {},
        personContractCharges: {},
        payouts: normalizePayoutSettings({}),
        invoices: normalizeInvoicesConfig({})
      }
    };
  }

  const rawMonthSettings = (m.monthSettings && typeof m.monthSettings === 'object') ? m.monthSettings : {};
  const normalizedMonthSettings = {
    ...rawMonthSettings,
    persons: (rawMonthSettings.persons && typeof rawMonthSettings.persons === 'object') ? rawMonthSettings.persons : {},
    clients: (rawMonthSettings.clients && typeof rawMonthSettings.clients === 'object') ? rawMonthSettings.clients : {},
    settlementConfig: (rawMonthSettings.settlementConfig && typeof rawMonthSettings.settlementConfig === 'object') ? rawMonthSettings.settlementConfig : {},
    personContractCharges: (rawMonthSettings.personContractCharges && typeof rawMonthSettings.personContractCharges === 'object') ? rawMonthSettings.personContractCharges : {},
    payouts: normalizePayoutSettings(rawMonthSettings.payouts || {}),
    invoices: normalizeInvoicesConfig(rawMonthSettings.invoices || {})
  };

  if (rawMonthSettings.commonSnapshot) {
    normalizedMonthSettings.commonSnapshot = normalizeCommonData(rawMonthSettings.commonSnapshot, createDefaultState().common);
  }

  return {
    monthlySheets: normalizeRemoteCollectionArray(m.monthlySheets)
      .filter(sheet => sheet && typeof sheet === 'object')
      .map(sheet => ({ ...sheet, days: normalizeMonthDays(sheet.days) })),
    worksSheets: normalizeRemoteCollectionArray(m.worksSheets)
      .filter(sheet => sheet && typeof sheet === 'object'),
    expenses: normalizeRemoteCollectionArray(m.expenses)
      .filter(expense => expense && typeof expense === 'object')
      .map(expense => normalizeExpenseRecord(expense)),
    monthSettings: normalizedMonthSettings
  };
};

const normalizeState = (rawState) => {
  if (typeof Migration !== 'undefined' && Migration.isV2State(rawState)) {
    const migrated = Migration.v2ToV3(rawState);
    if (migrated.selectedMonth) localSettings.selectedMonth = migrated.selectedMonth;
    return normalizeState({ version: 'v3', common: migrated.common, months: migrated.months });
  }

  const baseState = createDefaultState();
  const candidate = rawState && typeof rawState === 'object' && !Array.isArray(rawState) ? rawState : {};
  
  if (candidate.version !== 'v3') return baseState;

  const common = candidate.common || {};
  const normalizedCommon = normalizeCommonData(common, baseState.common);

  const normalizedMonths = {};
  if (candidate.months && typeof candidate.months === 'object') {
    Object.entries(candidate.months).forEach(([key, m]) => {
      normalizedMonths[key] = normalizeMonth(m);
      normalizedMonths[key].monthlySheets.forEach(sheet => {
         if (!Array.isArray(sheet.activePersons)) {
           sheet.activePersons = getActivePersonIdsForMonthFromState({ common: normalizedCommon, months: normalizedMonths }, key);
         }
      });
    });
  }

  const finalState = { version: 'v3', common: normalizedCommon, months: normalizedMonths };

  Object.values(finalState.months).forEach(monthRecord => {
    ensureArchivedMonthHasSnapshot(monthRecord, finalState.common);
  });

  finalState.common.worksCatalog.forEach(w => {
    if (w.unit === 'm3') w.unit = 'm³';
    if (w.unit === 'm2') w.unit = 'm²';
    if (!w.coreId) {
       const def = DEFAULT_WORKS_CATALOG.find(d => d.name === w.name);
       if (def) w.coreId = def.coreId;
    }
  });

  Object.entries(finalState.months).forEach(([key, m]) => {
    m.worksSheets.forEach(s => {
      if (s.entries) {
        s.entries.forEach(e => {
          if (e.unit === 'm3') e.unit = 'm³';
          if (e.unit === 'm2') e.unit = 'm²';
        });
      }
    });
    if (m.monthSettings) {
      m.monthSettings.payouts = normalizePayoutSettings(m.monthSettings.payouts || {});
      m.monthSettings.invoices = normalizeInvoicesConfig(m.monthSettings.invoices || {});
      if (m.monthSettings.commonSnapshot) {
        m.monthSettings.commonSnapshot = normalizeCommonData(m.monthSettings.commonSnapshot, baseState.common);
      }
    }
  });

  return finalState;
};

const ensureMonthSettings = (month) => {
  const monthKey = month || Store.getSelectedMonth() || getCurrentMonthKey();
  if (!appState.months[monthKey]) {
    appState.months[monthKey] = {
      monthlySheets: [],
      worksSheets: [],
      expenses: [],
      monthSettings: {
        persons: {},
        clients: {},
        settlementConfig: {},
        personContractCharges: {},
        invoices: { issueDate: '', emailIntro: '', clients: {} }
      }
    };
  }
  const m = appState.months[monthKey];
  if (!m.monthSettings) m.monthSettings = {};
  if (!m.monthSettings.persons) m.monthSettings.persons = {};
  if (!m.monthSettings.clients) m.monthSettings.clients = {};
  if (!m.monthSettings.settlementConfig) m.monthSettings.settlementConfig = {};
  if (!m.monthSettings.personContractCharges) m.monthSettings.personContractCharges = {};
  if (!m.monthSettings.payouts) m.monthSettings.payouts = normalizePayoutSettings({});
  if (!m.monthSettings.invoices) m.monthSettings.invoices = { issueDate: '', emailIntro: '', clients: {} };
  ensureArchivedMonthHasSnapshot(m, appState.common);
  return m.monthSettings;
};

const getMonthRecord = (month) => appState.months[month] || null;

const getCommonDataForMonth = (month) => {
  const monthRecord = getMonthRecord(month);
  if (!monthRecord) return appState.common;

  const monthSettings = ensureMonthSettings(month);
  return monthSettings.commonSnapshot || appState.common;
};

const getMutableCommonDataForMonth = (month) => {
  const monthSettings = ensureMonthSettings(month);
  if (monthSettings.commonSnapshot) {
    return monthSettings.commonSnapshot;
  }
  return appState.common;
};

const PERSON_REFERENCE_ID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const normalizePersonReferenceId = (value = '') => (value || '').toString().trim();

const createOrphanedPersonCleanupSummary = (month = '') => ({
  success: true,
  changed: false,
  month,
  removedPersonIds: [],
  removedMonthlySheetReferences: 0,
  removedWorksSheetReferences: 0,
  removedExpenseEntries: 0,
  removedMonthSettingReferences: 0,
  removedInvoiceReferences: 0,
  totalRemoved: 0
});

const cleanupOrphanedPersonEntriesInMonth = (month = Store.getSelectedMonth(), targetPersonIds = []) => {
  const monthRecord = getMonthRecord(month);
  const summary = createOrphanedPersonCleanupSummary(month);
  if (!monthRecord) return summary;

  const validPersonIds = new Set((getCommonDataForMonth(month)?.persons || [])
    .map(person => normalizePersonReferenceId(person?.id))
    .filter(Boolean));
  const targetIds = new Set((Array.isArray(targetPersonIds) ? targetPersonIds : [targetPersonIds])
    .map(personId => normalizePersonReferenceId(personId))
    .filter(Boolean));
  const removedPersonIds = new Set();
  const markRemoved = (personId = '') => {
    const normalizedId = normalizePersonReferenceId(personId);
    if (!normalizedId) return;
    removedPersonIds.add(normalizedId);
    summary.changed = true;
    summary.totalRemoved += 1;
  };
  const shouldRemovePersonOnlyReference = (personId = '') => {
    const normalizedId = normalizePersonReferenceId(personId);
    if (!normalizedId || validPersonIds.has(normalizedId)) return false;
    return targetIds.size === 0 || targetIds.has(normalizedId);
  };
  const shouldRemoveExpensePaidById = (paidById = '') => {
    const normalizedId = normalizePersonReferenceId(paidById);
    if (!normalizedId) return false;
    if (targetIds.size > 0) return targetIds.has(normalizedId) && !validPersonIds.has(normalizedId);
    return PERSON_REFERENCE_ID_PATTERN.test(normalizedId) && !validPersonIds.has(normalizedId);
  };
  const removeObjectReferences = (targetObject = null, counterKey = '') => {
    if (!targetObject || typeof targetObject !== 'object') return;
    Object.keys(targetObject).forEach(personId => {
      if (!shouldRemovePersonOnlyReference(personId)) return;
      delete targetObject[personId];
      summary[counterKey] += 1;
      markRemoved(personId);
    });
  };

  (monthRecord.monthlySheets || []).forEach(sheet => {
    if (Array.isArray(sheet.activePersons)) {
      const nextActivePersons = sheet.activePersons.filter(personId => {
        if (!shouldRemovePersonOnlyReference(personId)) return true;
        summary.removedMonthlySheetReferences += 1;
        markRemoved(personId);
        return false;
      });
      if (nextActivePersons.length !== sheet.activePersons.length) {
        sheet.activePersons = nextActivePersons;
      }
    }

    Object.values(sheet.days || {}).forEach(day => {
      removeObjectReferences(day?.hours, 'removedMonthlySheetReferences');
      removeObjectReferences(day?.manual, 'removedMonthlySheetReferences');
      removeObjectReferences(day?.activityOverrides, 'removedMonthlySheetReferences');
    });
  });

  (monthRecord.worksSheets || []).forEach(sheet => {
    if (Array.isArray(sheet.activePersons)) {
      const nextActivePersons = sheet.activePersons.filter(personId => {
        if (!shouldRemovePersonOnlyReference(personId)) return true;
        summary.removedWorksSheetReferences += 1;
        markRemoved(personId);
        return false;
      });
      if (nextActivePersons.length !== sheet.activePersons.length) {
        sheet.activePersons = nextActivePersons;
      }
    }

    (sheet.entries || []).forEach(entry => {
      removeObjectReferences(entry?.hours, 'removedWorksSheetReferences');
    });
  });

  monthRecord.expenses = (monthRecord.expenses || []).filter(expense => {
    const shouldRemoveExpense = shouldRemovePersonOnlyReference(expense?.advanceForId)
      || shouldRemoveExpensePaidById(expense?.paidById);
    if (!shouldRemoveExpense) return true;

    summary.removedExpenseEntries += 1;
    markRemoved(expense?.advanceForId || expense?.paidById || '');
    return false;
  });

  const monthSettings = ensureMonthSettings(month);
  removeObjectReferences(monthSettings.persons, 'removedMonthSettingReferences');
  removeObjectReferences(monthSettings.personContractCharges, 'removedMonthSettingReferences');

  const payoutEmployees = monthSettings.payouts?.employees;
  if (payoutEmployees && typeof payoutEmployees === 'object') {
    Object.keys(payoutEmployees).forEach(personId => {
      if (!shouldRemovePersonOnlyReference(personId)) return;
      delete payoutEmployees[personId];
      summary.removedMonthSettingReferences += 1;
      markRemoved(personId);
    });
    monthSettings.payouts = normalizePayoutSettings(monthSettings.payouts || {});
  }

  monthSettings.invoices = normalizeInvoicesConfig({
    ...(monthSettings.invoices || {}),
    clients: Object.fromEntries(Object.entries(monthSettings.invoices?.clients || {}).map(([clientKey, clientConfig]) => {
      const normalizedConfig = { ...(clientConfig || {}) };

      if (Array.isArray(normalizedConfig.issuerIds)) {
        normalizedConfig.issuerIds = normalizedConfig.issuerIds.filter(personId => {
          if (!shouldRemovePersonOnlyReference(personId)) return true;
          summary.removedInvoiceReferences += 1;
          markRemoved(personId);
          return false;
        });
      }

      ['percentageAllocations', 'manualAmounts', 'separateCompanyWithEmployees'].forEach(key => {
        if (!normalizedConfig[key] || typeof normalizedConfig[key] !== 'object') return;
        Object.keys(normalizedConfig[key]).forEach(personId => {
          if (!shouldRemovePersonOnlyReference(personId)) return;
          delete normalizedConfig[key][personId];
          summary.removedInvoiceReferences += 1;
          markRemoved(personId);
        });
      });

      if (Array.isArray(normalizedConfig.percentageTouchedIssuerIds)) {
        normalizedConfig.percentageTouchedIssuerIds = normalizedConfig.percentageTouchedIssuerIds.filter(personId => {
          if (!shouldRemovePersonOnlyReference(personId)) return true;
          summary.removedInvoiceReferences += 1;
          markRemoved(personId);
          return false;
        });
      }

      if (shouldRemovePersonOnlyReference(normalizedConfig.lastEditedPercentageIssuerId)) {
        normalizedConfig.lastEditedPercentageIssuerId = '';
        summary.removedInvoiceReferences += 1;
        markRemoved(clientConfig?.lastEditedPercentageIssuerId || '');
      }

      return [clientKey, normalizedConfig];
    })),
    extraInvoices: (monthSettings.invoices?.extraInvoices || []).filter(invoice => {
      if (!shouldRemovePersonOnlyReference(invoice?.issuerId)) return true;
      summary.removedInvoiceReferences += 1;
      markRemoved(invoice?.issuerId || '');
      return false;
    })
  });

  summary.removedPersonIds = [...removedPersonIds].sort((a, b) => a.localeCompare(b, 'pl-PL'));
  return summary;
};

const buildOrphanedPersonCleanupTrackedChanges = (month = Store.getSelectedMonth()) => ([
  { scope: 'month.monthlySheets', month, label: 'Usunięto osierocone wpisy osób z arkuszy godzin' },
  { scope: 'month.worksSheets', month, label: 'Usunięto osierocone wpisy osób z arkuszy prac' },
  { scope: 'month.expenses', month, label: 'Usunięto osierocone wpisy osób z kosztów i zaliczek' },
  { scope: 'month.monthSettings.persons', month, label: 'Usunięto osierocone statusy osób z miesiąca' },
  { scope: 'month.monthSettings.personContractCharges', month, label: 'Usunięto osierocone stawki UZ z miesiąca' },
  { scope: 'month.monthSettings.payouts', month, label: 'Usunięto osierocone ustawienia wypłat z miesiąca' },
  { scope: 'month.monthSettings.invoices', month, label: 'Usunięto osierocone wpisy osób z faktur miesiąca' }
]);

const resolveCommonMutationScope = (baseScope, month = Store.getSelectedMonth()) => {
  const monthSettings = ensureMonthSettings(month);
  if (monthSettings.commonSnapshot) {
    return `month.monthSettings.commonSnapshot.${baseScope.replace(/^common\./, '')}`;
  }
  return baseScope;
};

const getScopeSnapshot = (scopeKey, month = Store.getSelectedMonth()) => {
  const monthKey = month || Store.getSelectedMonth() || getCurrentMonthKey();
  const monthRecord = appState.months[monthKey] || normalizeMonth(null);
  const monthSettings = ensureMonthSettings(monthKey);

  switch (scopeKey) {
    case 'common.persons':
      return cloneData(appState.common.persons || []);
    case 'common.clients':
      return cloneData(appState.common.clients || []);
    case 'common.worksCatalog':
      return cloneData(appState.common.worksCatalog || []);
    case 'common.config':
      return cloneData(appState.common.config || {});
    case 'month.monthlySheets':
      return cloneData(monthRecord.monthlySheets || []);
    case 'month.worksSheets':
      return cloneData(monthRecord.worksSheets || []);
    case 'month.expenses':
      return cloneData(monthRecord.expenses || []);
    case 'month.monthSettings.persons':
      return cloneData(monthSettings.persons || {});
    case 'month.monthSettings.clients':
      return cloneData(monthSettings.clients || {});
    case 'month.monthSettings.settlementConfig':
      return cloneData(monthSettings.settlementConfig || {});
    case 'month.monthSettings.personContractCharges':
      return cloneData(monthSettings.personContractCharges || {});
    case 'month.monthSettings.payouts':
      return cloneData(monthSettings.payouts || normalizePayoutSettings({}));
    case 'month.monthSettings.invoices':
      return cloneData(monthSettings.invoices || normalizeInvoicesConfig({}));
    case 'month.monthSettings.archive':
      return cloneData({
        isArchived: monthSettings.isArchived === true,
        commonSnapshot: monthSettings.commonSnapshot || null
      });
    case 'month.monthSettings.commonSnapshot.persons':
      return cloneData(monthSettings.commonSnapshot?.persons || []);
    case 'month.monthSettings.commonSnapshot.clients':
      return cloneData(monthSettings.commonSnapshot?.clients || []);
    case 'month.monthSettings.commonSnapshot.worksCatalog':
      return cloneData(monthSettings.commonSnapshot?.worksCatalog || []);
    case 'month.monthSettings.commonSnapshot.config':
      return cloneData(monthSettings.commonSnapshot?.config || {});
    default:
      return null;
  }
};

const normalizeScopeSnapshot = (scopeKey, snapshot) => {
  switch (scopeKey) {
    case 'common.persons':
    case 'month.monthSettings.commonSnapshot.persons':
      return Array.isArray(snapshot) ? snapshot.map(person => normalizePersonRecord(person)) : [];
    case 'common.clients':
    case 'month.monthSettings.commonSnapshot.clients':
      return Array.isArray(snapshot) ? snapshot.map(client => normalizeClientRecord(client)) : [];
    case 'common.worksCatalog':
      return Array.isArray(snapshot) && snapshot.length > 0 ? snapshot : createDefaultState().common.worksCatalog;
    case 'month.monthSettings.commonSnapshot.worksCatalog':
      return Array.isArray(snapshot) ? snapshot : [];
    case 'common.config':
    case 'month.monthSettings.commonSnapshot.config':
      return { ...createDefaultState().common.config, ...(snapshot || {}) };
    case 'month.monthlySheets':
      return normalizeMonth({ monthlySheets: Array.isArray(snapshot) ? snapshot : [] }).monthlySheets;
    case 'month.worksSheets':
      return normalizeMonth({ worksSheets: Array.isArray(snapshot) ? snapshot : [] }).worksSheets;
    case 'month.expenses':
      return normalizeMonth({ expenses: Array.isArray(snapshot) ? snapshot : [] }).expenses;
    case 'month.monthSettings.persons':
    case 'month.monthSettings.clients':
    case 'month.monthSettings.settlementConfig':
    case 'month.monthSettings.personContractCharges':
      return { ...(snapshot || {}) };
    case 'month.monthSettings.payouts':
      return normalizePayoutSettings(snapshot || {});
    case 'month.monthSettings.invoices':
      return normalizeInvoicesConfig(snapshot || {});
    case 'month.monthSettings.archive':
      return {
        isArchived: snapshot?.isArchived === true,
        commonSnapshot: snapshot?.commonSnapshot ? normalizeCommonData(snapshot.commonSnapshot, createDefaultState().common) : null
      };
    default:
      return cloneData(snapshot);
  }
};

const setScopeSnapshot = (scopeKey, month = Store.getSelectedMonth(), snapshot) => {
  const monthKey = month || Store.getSelectedMonth() || getCurrentMonthKey();
  const normalizedSnapshot = normalizeScopeSnapshot(scopeKey, snapshot);

  switch (scopeKey) {
    case 'common.persons':
      appState.common.persons = normalizedSnapshot;
      return;
    case 'common.clients':
      appState.common.clients = normalizedSnapshot;
      return;
    case 'common.worksCatalog':
      appState.common.worksCatalog = normalizedSnapshot;
      return;
    case 'common.config':
      appState.common.config = normalizedSnapshot;
      return;
    case 'month.monthlySheets':
      ensureMonthSettings(monthKey);
      appState.months[monthKey].monthlySheets = normalizedSnapshot;
      return;
    case 'month.worksSheets':
      ensureMonthSettings(monthKey);
      appState.months[monthKey].worksSheets = normalizedSnapshot;
      return;
    case 'month.expenses':
      ensureMonthSettings(monthKey);
      appState.months[monthKey].expenses = normalizedSnapshot;
      return;
    case 'month.monthSettings.persons':
      ensureMonthSettings(monthKey).persons = normalizedSnapshot;
      return;
    case 'month.monthSettings.clients':
      ensureMonthSettings(monthKey).clients = normalizedSnapshot;
      return;
    case 'month.monthSettings.settlementConfig':
      ensureMonthSettings(monthKey).settlementConfig = normalizedSnapshot;
      return;
    case 'month.monthSettings.personContractCharges':
      ensureMonthSettings(monthKey).personContractCharges = normalizedSnapshot;
      return;
    case 'month.monthSettings.payouts':
      ensureMonthSettings(monthKey).payouts = normalizedSnapshot;
      return;
    case 'month.monthSettings.invoices':
      ensureMonthSettings(monthKey).invoices = normalizedSnapshot;
      return;
    case 'month.monthSettings.archive': {
      const monthSettings = ensureMonthSettings(monthKey);
      monthSettings.isArchived = normalizedSnapshot.isArchived === true;
      if (normalizedSnapshot.commonSnapshot) {
        monthSettings.commonSnapshot = normalizedSnapshot.commonSnapshot;
      } else {
        delete monthSettings.commonSnapshot;
      }
      ensureArchivedMonthHasSnapshot(appState.months[monthKey], appState.common);
      return;
    }
    case 'month.monthSettings.commonSnapshot.persons':
    case 'month.monthSettings.commonSnapshot.clients':
    case 'month.monthSettings.commonSnapshot.worksCatalog':
    case 'month.monthSettings.commonSnapshot.config': {
      const snapshotKey = scopeKey.replace('month.monthSettings.commonSnapshot.', '');
      const monthSettings = ensureMonthSettings(monthKey);
      if (!monthSettings.commonSnapshot) {
        monthSettings.commonSnapshot = cloneData(appState.common);
      }
      monthSettings.commonSnapshot[snapshotKey] = normalizedSnapshot;
      return;
    }
    default:
      return;
  }
};

const getCollectionItemsWithId = (snapshot = []) => Array.isArray(snapshot)
  ? snapshot.filter(item => item && typeof item === 'object' && item.id)
  : [];

const getCollectionAddRemoveDiff = (beforeSnapshot = [], afterSnapshot = []) => {
  const beforeItems = getCollectionItemsWithId(beforeSnapshot);
  const afterItems = getCollectionItemsWithId(afterSnapshot);
  const beforeById = new Map(beforeItems.map(item => [item.id, item]));
  const afterById = new Map(afterItems.map(item => [item.id, item]));

  return {
    addedItems: afterItems.filter(item => !beforeById.has(item.id)),
    removedItems: beforeItems.filter(item => !afterById.has(item.id))
  };
};

const buildCollectionMutationDiff = (beforeSnapshot = [], afterSnapshot = []) => {
  const beforeItems = getCollectionItemsWithId(beforeSnapshot);
  const afterItems = getCollectionItemsWithId(afterSnapshot);
  const beforeById = new Map(beforeItems.map(item => [item.id, item]));
  const afterById = new Map(afterItems.map(item => [item.id, item]));
  const beforeOrder = beforeItems.map(item => item.id);
  const afterOrder = afterItems.map(item => item.id);
  const sortOrderChangedItems = afterItems.filter((item, index) => {
    if (!item?.id) return false;
    return beforeOrder[index] !== item.id;
  });
  const touchedItemsById = new Map();

  [
    ...afterItems.filter(item => !beforeById.has(item.id)),
    ...afterItems.filter(item => beforeById.has(item.id) && !deepEqual(beforeById.get(item.id), item)),
    ...sortOrderChangedItems
  ].forEach(item => {
    if (item?.id) touchedItemsById.set(item.id, item);
  });

  return {
    addedItems: afterItems.filter(item => !beforeById.has(item.id)),
    removedItems: beforeItems.filter(item => !afterById.has(item.id)),
    updatedItems: afterItems.filter(item => beforeById.has(item.id) && !deepEqual(beforeById.get(item.id), item)),
    touchedItems: [...touchedItemsById.values()],
    sortOrderChangedItemIds: sortOrderChangedItems.map(item => item.id).filter(Boolean),
    orderChanged: beforeOrder.length !== afterOrder.length || beforeOrder.some((itemId, index) => itemId !== afterOrder[index]),
    sortOrderById: afterItems.reduce((acc, item, index) => {
      acc[item.id] = index;
      return acc;
    }, {})
  };
};

const buildMonthlySheetDayMutations = (beforeSnapshot = [], afterSnapshot = []) => {
  const beforeItems = getCollectionItemsWithId(beforeSnapshot);
  const afterItems = getCollectionItemsWithId(afterSnapshot);
  const beforeById = new Map(beforeItems.map(item => [item.id, item]));

  return afterItems.reduce((acc, afterItem) => {
    const beforeItem = beforeById.get(afterItem.id);
    if (!beforeItem) return acc;

    const beforeComparable = cloneData(beforeItem);
    const afterComparable = cloneData(afterItem);
    delete beforeComparable.days;
    delete afterComparable.days;

    const beforeDays = normalizeMonthDays(beforeItem.days);
    const afterDays = normalizeMonthDays(afterItem.days);
    if (!deepEqual(beforeComparable, afterComparable) || deepEqual(beforeDays, afterDays)) {
      return acc;
    }

    const dayKeys = [...new Set([...Object.keys(beforeDays), ...Object.keys(afterDays)])];
    const dayUpdates = dayKeys.reduce((dayAcc, dayKey) => {
      const beforeDay = beforeDays[dayKey];
      const afterDay = afterDays[dayKey];
      if (deepEqual(beforeDay, afterDay)) return dayAcc;
      dayAcc[dayKey] = afterDay === undefined ? null : afterDay;
      return dayAcc;
    }, {});

    if (Object.keys(dayUpdates).length > 0) {
      acc[afterItem.id] = { dayUpdates };
    }

    return acc;
  }, {});
};

const buildMergedCollectionSnapshot = (currentSnapshot = [], itemsToMerge = [], preferredOrderSnapshot = []) => {
  const currentItems = Array.isArray(currentSnapshot) ? currentSnapshot : [];
  const currentItemsWithId = getCollectionItemsWithId(currentItems);
  const currentById = new Map(currentItemsWithId.map(item => [item.id, item]));

  itemsToMerge.forEach(item => {
    if (item?.id && !currentById.has(item.id)) {
      currentById.set(item.id, item);
    }
  });

  const orderedIds = [...new Set([
    ...getCollectionItemsWithId(preferredOrderSnapshot).map(item => item.id),
    ...currentItemsWithId.map(item => item.id),
    ...getCollectionItemsWithId(itemsToMerge).map(item => item.id)
  ])];

  const mergedItems = orderedIds
    .map(itemId => currentById.get(itemId))
    .filter(Boolean);

  const itemsWithoutId = currentItems.filter(item => !(item && typeof item === 'object' && item.id));
  return [...mergedItems, ...itemsWithoutId];
};

const getInvoiceExtraInvoicesSnapshot = (snapshot = {}) => Array.isArray(snapshot?.extraInvoices)
  ? snapshot.extraInvoices
  : [];

const buildSharedHistoryCollectionActionChanges = (entry = null, actionMode = 'merge-before') => {
  if (!entry?.changes || !Array.isArray(entry.changes)) return [];

  return entry.changes.map(change => {
    const month = change.month || entry.month || Store.getSelectedMonth();

    if (change.scope === 'month.monthSettings.invoices') {
      const beforeInvoices = normalizeInvoicesConfig(change.beforeSnapshot || {});
      const afterInvoices = normalizeInvoicesConfig(change.afterSnapshot || {});
      const currentInvoices = normalizeInvoicesConfig(getScopeSnapshot(change.scope, month) || {});
      const { addedItems, removedItems } = getCollectionAddRemoveDiff(
        getInvoiceExtraInvoicesSnapshot(beforeInvoices),
        getInvoiceExtraInvoicesSnapshot(afterInvoices)
      );

      if (actionMode === 'merge-before' && removedItems.length > 0) {
        const nextExtraInvoices = buildMergedCollectionSnapshot(
          currentInvoices.extraInvoices || [],
          removedItems,
          beforeInvoices.extraInvoices || []
        );
        const nextSnapshot = { ...currentInvoices, extraInvoices: nextExtraInvoices };
        return deepEqual(currentInvoices, nextSnapshot) ? null : { scope: change.scope, month, nextSnapshot };
      }

      if (actionMode === 'remove-added' && addedItems.length > 0) {
        const addedIds = new Set(addedItems.map(item => item.id));
        const nextExtraInvoices = (currentInvoices.extraInvoices || []).filter(item => !addedIds.has(item?.id));
        const nextSnapshot = { ...currentInvoices, extraInvoices: nextExtraInvoices };
        return deepEqual(currentInvoices, nextSnapshot) ? null : { scope: change.scope, month, nextSnapshot };
      }

      return null;
    }

    if (!Array.isArray(change.beforeSnapshot) || !Array.isArray(change.afterSnapshot)) {
      return null;
    }

    const currentSnapshot = getScopeSnapshot(change.scope, month);
    if (!Array.isArray(currentSnapshot)) {
      return null;
    }

    const { addedItems, removedItems } = getCollectionAddRemoveDiff(change.beforeSnapshot, change.afterSnapshot);

    if (actionMode === 'merge-before' && removedItems.length > 0) {
      const nextSnapshot = buildMergedCollectionSnapshot(currentSnapshot, removedItems, change.beforeSnapshot);
      return deepEqual(currentSnapshot, nextSnapshot) ? null : { scope: change.scope, month, nextSnapshot };
    }

    if (actionMode === 'remove-added' && addedItems.length > 0) {
      const addedIds = new Set(addedItems.map(item => item.id));
      const nextSnapshot = currentSnapshot.filter(item => !addedIds.has(item?.id));
      return deepEqual(currentSnapshot, nextSnapshot) ? null : { scope: change.scope, month, nextSnapshot };
    }

    return null;
  }).filter(Boolean);
};

const applySharedHistoryCollectionAction = (entryId = '', actionMode = 'merge-before') => {
  const entry = sharedHistoryEntries.find(item => item.id === entryId);
  if (!entry) return { success: false, message: 'Nie znaleziono wpisu historii zmian.' };
  if (!historyEntryHasLoadedDetails(entry)) {
    return { success: false, message: 'Najpierw wczytaj szczegóły tego wpisu historii.' };
  }

  const applicableChanges = buildSharedHistoryCollectionActionChanges(entry, actionMode);
  if (applicableChanges.length === 0) {
    return {
      success: false,
      message: actionMode === 'merge-before'
        ? 'Brak usuniętych wpisów do przywrócenia i połączenia z aktualnym stanem.'
        : 'Brak dodanych wpisów do usunięcia z aktualnego stanu.'
    };
  }

  const contextLabel = actionMode === 'merge-before'
    ? `Przywrócono i połączono stan sprzed zmiany: ${entry.label}`
    : `Usunięto wpisy dodane przez zmianę: ${entry.label}`;

  const context = createTrackedChangeContext(
    applicableChanges.map(change => ({ scope: change.scope, month: change.month })),
    contextLabel,
    {
      selectedMonth: entry.month || Store.getSelectedMonth(),
      skipLocalHistory: true,
      forceImmediateRemoteSync: true
    }
  );

  applicableChanges.forEach(change => {
    setScopeSnapshot(change.scope, change.month, change.nextSnapshot);
  });

  activeMutationTransaction = context;
  finalizeTrackedChangeContext(entry.month || Store.getSelectedMonth());
  localRedoStack = [];
  return { success: true };
};

const getScopeMonthKey = (change = {}) => change.month || Store.getSelectedMonth() || getCurrentMonthKey();

let activeMutationTransaction = null;
let localUndoStack = [];
let localRedoStack = [];
let sharedHistoryEntries = [];
let dailyBackupEntries = [];
let monthlyBackupEntries = [];

const recordedDailyBackupKeys = new Set();
const recordedMonthlyBackupKeys = new Set();

const rebuildRecordedBackupKeys = () => {
  recordedDailyBackupKeys.clear();
  recordedMonthlyBackupKeys.clear();

  dailyBackupEntries.forEach(entry => {
    const dayKey = (entry?.dayKey || '').toString().trim();
    if (dayKey) recordedDailyBackupKeys.add(dayKey);
  });

  monthlyBackupEntries.forEach(entry => {
    const monthKey = (entry?.month || '').toString().trim();
    if (monthKey) recordedMonthlyBackupKeys.add(monthKey);
  });
};

const ensurePeriodicBackupsForLogin = () => {
  const currentMonthKey = getCurrentMonthKey();
  const todayKey = getTodayKey();
  let changed = false;
  let dailyRemoteUpdates = {};
  let monthlyRemoteUpdates = {};

  ensureMonthSettings(currentMonthKey);

  if (!recordedDailyBackupKeys.has(todayKey)) {
    const nextDailyEntry = createHistoryBackupPayload('daily', currentMonthKey, cloneData(appState.common), cloneData(appState.months[currentMonthKey] || normalizeMonth(null)));
    const dailyEntriesResult = prependNormalizedEntriesWithLimit(dailyBackupEntries, nextDailyEntry, SHARED_DAILY_BACKUPS_LIMIT, normalizeBackupEntries);
    dailyBackupEntries = dailyEntriesResult.entries;
    dailyRemoteUpdates = buildIndexedRemoteUpdatesForEntries(
      SHARED_DAILY_BACKUPS_INDEX_ROOT,
      SHARED_DAILY_BACKUPS_ENTRIES_ROOT,
      [nextDailyEntry],
      dailyEntriesResult.removedEntries,
      buildBackupIndexEntry
    );
    changed = true;
  }

  if (!recordedMonthlyBackupKeys.has(currentMonthKey)) {
    const nextMonthlyEntry = createHistoryBackupPayload('monthly', currentMonthKey, cloneData(appState.common), cloneData(appState.months[currentMonthKey] || normalizeMonth(null)), cloneData(appState));
    const monthlyEntriesResult = prependNormalizedEntriesWithLimit(monthlyBackupEntries, nextMonthlyEntry, SHARED_MONTHLY_BACKUPS_LIMIT, normalizeBackupEntries);
    monthlyBackupEntries = monthlyEntriesResult.entries;
    monthlyRemoteUpdates = buildIndexedRemoteUpdatesForEntries(
      SHARED_MONTHLY_BACKUPS_INDEX_ROOT,
      SHARED_MONTHLY_BACKUPS_ENTRIES_ROOT,
      [nextMonthlyEntry],
      monthlyEntriesResult.removedEntries,
      buildBackupIndexEntry
    );
    changed = true;
  }

  if (!changed) return false;

  rebuildRecordedBackupKeys();
  applyTrackedChangePersistence([currentMonthKey], false, {
    ...dailyRemoteUpdates,
    ...monthlyRemoteUpdates
  }, {
    immediateRemoteSync: true
  });
  return true;
};

const replaceHistoryBackupEntry = (type = 'daily') => {
  const currentMonthKey = getCurrentMonthKey();
  ensureMonthSettings(currentMonthKey);

  const nextEntry = type === 'monthly'
    ? createHistoryBackupPayload('monthly', currentMonthKey, cloneData(appState.common), cloneData(appState.months[currentMonthKey] || normalizeMonth(null)), cloneData(appState))
    : createHistoryBackupPayload('daily', currentMonthKey, cloneData(appState.common), cloneData(appState.months[currentMonthKey] || normalizeMonth(null)));
  let remoteUpdates = {};

  if (type === 'monthly') {
    const removedEntries = monthlyBackupEntries.filter(entry => (entry?.month || '') === currentMonthKey);
    const nextEntriesResult = prependNormalizedEntriesWithLimit(
      monthlyBackupEntries.filter(entry => (entry?.month || '') !== currentMonthKey),
      nextEntry,
      SHARED_MONTHLY_BACKUPS_LIMIT,
      normalizeBackupEntries
    );
    monthlyBackupEntries = nextEntriesResult.entries;
    remoteUpdates = buildIndexedRemoteUpdatesForEntries(
      SHARED_MONTHLY_BACKUPS_INDEX_ROOT,
      SHARED_MONTHLY_BACKUPS_ENTRIES_ROOT,
      [nextEntry],
      [...removedEntries, ...nextEntriesResult.removedEntries],
      buildBackupIndexEntry
    );
  } else {
    const todayKey = getTodayKey();
    const removedEntries = dailyBackupEntries.filter(entry => (entry?.dayKey || '') === todayKey);
    const nextEntriesResult = prependNormalizedEntriesWithLimit(
      dailyBackupEntries.filter(entry => (entry?.dayKey || '') !== todayKey),
      nextEntry,
      SHARED_DAILY_BACKUPS_LIMIT,
      normalizeBackupEntries
    );
    dailyBackupEntries = nextEntriesResult.entries;
    remoteUpdates = buildIndexedRemoteUpdatesForEntries(
      SHARED_DAILY_BACKUPS_INDEX_ROOT,
      SHARED_DAILY_BACKUPS_ENTRIES_ROOT,
      [nextEntry],
      [...removedEntries, ...nextEntriesResult.removedEntries],
      buildBackupIndexEntry
    );
  }

  rebuildRecordedBackupKeys();
  applyTrackedChangePersistence([currentMonthKey], false, remoteUpdates);

  return { success: true, entry: cloneData(nextEntry) };
};

const clearPersistedStateFromLocalStorage = (preservedLocalSettings = null) => {
  Object.keys(localStorage).forEach(key => {
    if (
      key === WORK_TRACKER_COMMON_KEY
      || key === WORK_TRACKER_SYNC_META_KEY
      || key === LEGACY_WORK_TRACKER_INVOICE_TOTALS_KEY
      || key === WORK_TRACKER_LEGACY_KEY
      || key.startsWith(WORK_TRACKER_MONTH_PREFIX)
    ) {
      localStorage.removeItem(key);
    }
  });

  if (preservedLocalSettings !== null) {
    localStorage.setItem(WORK_TRACKER_LOCAL_SETTINGS_KEY, preservedLocalSettings);
  }
};

const persistCurrentStateToLocalStorage = (monthKeys = [], persistCommon = true) => {
  if (persistCommon) {
    localStorage.setItem(WORK_TRACKER_COMMON_KEY, JSON.stringify(appState.common));
  }

  [...new Set(monthKeys.filter(Boolean))].forEach(monthKey => {
    if (appState.months[monthKey]) {
      localStorage.setItem(WORK_TRACKER_MONTH_PREFIX + monthKey, JSON.stringify(appState.months[monthKey]));
    }
  });

  persistSyncMetadataToLocalStorage();
};

const buildMonthKeysToPersist = (changes = [], fallbackMonth = '') => {
  const monthKeys = changes
    .filter(change => isMonthScope(change.scope))
    .map(change => getScopeMonthKey(change));

  if (fallbackMonth) monthKeys.push(fallbackMonth);
  return [...new Set(monthKeys.filter(Boolean))];
};

const createTrackedChangeContext = (changes = [], label = 'Zmiana', options = {}) => {
  const selectedMonth = options.selectedMonth || Store.getSelectedMonth() || getCurrentMonthKey();
  const normalizedChanges = changes.map(change => {
    const month = change.month || selectedMonth;
    const scope = change.scope;
    return {
      scope,
      month,
      scopeLabel: change.scopeLabel || getScopeDisplayLabel(scope),
      label: change.label || label,
      beforeSnapshot: cloneData(getScopeSnapshot(scope, month)),
      beforeRevision: getSyncMetaEntry(scope, month).revision
    };
  });

  const primaryMonth = normalizedChanges.find(change => isMonthScope(change.scope))?.month || selectedMonth;

  return {
    id: generateId(),
    label,
    timestamp: getIsoTimestamp(),
    author: getCurrentActorId(),
    selectedMonth,
    primaryMonth,
    persistCommon: normalizedChanges.some(change => isCommonScope(change.scope)),
    changes: normalizedChanges,
    dailyBackup: null,
    monthlyBackup: null,
    skipLocalHistory: options.skipLocalHistory === true,
    skipSharedHistory: options.skipSharedHistory === true,
    forceImmediateRemoteSync: options.forceImmediateRemoteSync === true
  };
};

const buildSharedHistoryEntryFromStateTransition = (beforeState = {}, afterState = {}, beforeMeta = {}, afterMeta = {}, label = 'Zmiana synchronizacji') => {
  const normalizedBeforeState = normalizeState(beforeState || {});
  const normalizedAfterState = normalizeState(afterState || {});
  const normalizedBeforeMeta = normalizeSyncMetadata(beforeMeta || {});
  const normalizedAfterMeta = normalizeSyncMetadata(afterMeta || {});
  const selectedMonth = Store.getSelectedMonth() || getCurrentMonthKey();
  const knownMonths = [...new Set([
    ...Object.keys(normalizedBeforeState.months || {}),
    ...Object.keys(normalizedAfterState.months || {})
  ])].sort((a, b) => b.localeCompare(a, 'pl-PL'));
  const trackedScopes = [
    ...COMMON_SCOPE_KEYS.map(scopeKey => ({ scope: `common.${scopeKey}`, month: selectedMonth })),
    ...knownMonths.flatMap(monthKey => MONTH_SCOPE_KEYS.map(scopeKey => ({ scope: `month.${scopeKey}`, month: monthKey })))
  ];

  const changes = trackedScopes.map(change => {
    const beforeSnapshot = getScopeSnapshotFromState(normalizedBeforeState, change.scope, change.month);
    const afterSnapshot = getScopeSnapshotFromState(normalizedAfterState, change.scope, change.month);
    if (deepEqual(beforeSnapshot, afterSnapshot)) return null;

    return {
      scope: change.scope,
      scopeLabel: getScopeDisplayLabel(change.scope),
      month: change.month,
      beforeSnapshot,
      afterSnapshot,
      beforeRevision: getRemoteMetaScopeRevision(normalizedBeforeMeta, change.scope, change.month),
      afterRevision: getRemoteMetaScopeRevision(normalizedAfterMeta, change.scope, change.month)
    };
  }).filter(Boolean);

  if (changes.length === 0) return null;

  const primaryMonth = changes.find(change => isMonthScope(change.scope))?.month || selectedMonth;
  return {
    id: generateId(),
    label,
    timestamp: getIsoTimestamp(),
    author: getCurrentActorId(),
    month: primaryMonth,
    changes
  };
};

const dispatchStateEvents = () => {
  window.dispatchEvent(new Event('appStateChanged'));
  window.dispatchEvent(new Event('historyStateChanged'));
};

const removeSharedHistoryEntryById = (entryId = '') => {
  sharedHistoryEntries = sharedHistoryEntries.filter(item => item.id !== entryId);
};

const restoreSharedHistoryEntryToTop = (entry = null) => {
  if (!entry?.id) return;
  removeSharedHistoryEntryById(entry.id);
  sharedHistoryEntries = limitSharedHistoryEntries([cloneData(entry), ...sharedHistoryEntries]).entries;
};

const prependSharedHistoryEntry = (entry = null) => {
  if (!entry?.id) {
    return { success: false, removedEntries: [] };
  }

  const result = limitSharedHistoryEntries([cloneData(entry), ...sharedHistoryEntries]);
  sharedHistoryEntries = result.entries;
  return {
    success: true,
    removedEntries: result.removedEntries
  };
};

const restoreSharedHistoryEntrySnapshot = (entryId, snapshotKey = 'beforeSnapshot') => {
  const entry = sharedHistoryEntries.find(item => item.id === entryId);
  if (!entry) return { success: false, message: 'Nie znaleziono wpisu historii zmian.' };
  if (!historyEntryHasLoadedDetails(entry)) {
    return { success: false, message: 'Najpierw wczytaj szczegóły tego wpisu historii.' };
  }

  const missingSnapshots = (entry.changes || []).filter(change => change?.[snapshotKey] === undefined);
  if (missingSnapshots.length > 0) {
    return { success: false, message: snapshotKey === 'afterSnapshot' ? 'Ten wpis historii nie zawiera stanu po zmianie.' : 'Ten wpis historii nie zawiera stanu sprzed zmiany.' };
  }

  const contextLabel = snapshotKey === 'afterSnapshot'
    ? `Przywrócono stan po zmianie: ${entry.label}`
    : `Przywrócono z historii: ${entry.label}`;

  const context = createTrackedChangeContext(entry.changes.map(change => ({ scope: change.scope, month: change.month })), contextLabel, {
    selectedMonth: entry.month || Store.getSelectedMonth(),
      skipLocalHistory: true,
      forceImmediateRemoteSync: true
  });

  context.skipLocalHistory = true;

  entry.changes.forEach(change => {
    setScopeSnapshot(change.scope, change.month, change[snapshotKey]);
  });

  activeMutationTransaction = context;
  finalizeTrackedChangeContext(entry.month || Store.getSelectedMonth());
  localRedoStack = [];
  dispatchStateEvents();
  return { success: true };
};

const buildMonthSettingsDataUpdates = (monthKey = '', monthSettingsPayload = null) => {
  const updates = {};
  const basePath = `/shared_data/months/${monthKey}/monthSettings`;
  const payload = monthSettingsPayload && typeof monthSettingsPayload === 'object' ? monthSettingsPayload : {};

  ['persons', 'clients', 'settlementConfig', 'personContractCharges', 'invoices'].forEach(key => {
    updates[`${basePath}/${key}`] = Object.prototype.hasOwnProperty.call(payload, key) ? payload[key] : null;
  });

  updates[`${basePath}/isArchived`] = payload.isArchived === true ? true : null;
  return updates;
};

const getRemoteCollectionMutationKey = (monthKey = '', scopeKey = '') => `${monthKey}::${scopeKey}`;

const COLLECTION_SYNC_SCOPE_KEYS = ['monthlySheets', 'worksSheets', 'expenses'];

const buildRemoteCollectionScopeMutations = (changes = []) => {
  return changes.reduce((acc, change) => {
    const monthKey = getScopeMonthKey(change);
    const syncScopeKey = getSyncScopeKey(change.scope);
    if (!COLLECTION_SYNC_SCOPE_KEYS.includes(syncScopeKey) || !Array.isArray(change.beforeSnapshot) || !Array.isArray(change.afterSnapshot)) {
      return acc;
    }

    acc[getRemoteCollectionMutationKey(monthKey, syncScopeKey)] = {
      ...buildCollectionMutationDiff(change.beforeSnapshot, change.afterSnapshot),
      ...(syncScopeKey === 'monthlySheets'
        ? { dayMutationsById: buildMonthlySheetDayMutations(change.beforeSnapshot, change.afterSnapshot) }
        : {})
    };
    return acc;
  }, {});
};

const buildCollectionScopeUpdatesFromMutation = (monthKey = '', scopeKey = '', mutation = null) => {
  const updates = {};
  if (!monthKey || !scopeKey || !mutation) return updates;

  const changedItems = mutation.touchedItems || [];
  const sortOrderChangedItemIds = new Set(mutation.sortOrderChangedItemIds || []);

  changedItems.forEach(item => {
    if (!item?.id) return;

    if (scopeKey === 'monthlySheets' && !sortOrderChangedItemIds.has(item.id) && mutation.dayMutationsById?.[item.id]) {
      Object.entries(mutation.dayMutationsById[item.id].dayUpdates || {}).forEach(([dayKey, dayValue]) => {
        updates[`/shared_data/months/${monthKey}/${scopeKey}/${item.id}/days/${dayKey}`] = dayValue;
      });
      return;
    }

    updates[`/shared_data/months/${monthKey}/${scopeKey}/${item.id}`] = serializeRemoteCollectionItem(item, mutation.sortOrderById?.[item.id] ?? 0);
  });

  (mutation.removedItems || []).forEach(item => {
    if (!item?.id) return;
    updates[`/shared_data/months/${monthKey}/${scopeKey}/${item.id}`] = null;
  });

  return updates;
};

const collectRemoteSyncTargetsFromChanges = (changes = [], persistCommon = false) => {
  const targets = new Map();
  const registerTarget = (scopeKey = '', monthKey = '') => {
    const syncScopeKey = getSyncScopeKey(scopeKey);
    const mapKey = syncScopeKey === 'common' ? 'common' : `${monthKey}:${syncScopeKey}`;
    if (!syncScopeKey || targets.has(mapKey)) return;
    targets.set(mapKey, { scopeKey: syncScopeKey, monthKey });
  };

  if (persistCommon) {
    registerTarget('common');
  }

  changes.forEach(change => {
    registerTarget(change.scope, getScopeMonthKey(change));
    if (change.scope === 'month.monthSettings.archive') {
      registerTarget('month.monthSettings.commonSnapshot', getScopeMonthKey(change));
    }
  });

  return [...targets.values()];
};

const buildRemoteUpdatesForSyncTargets = (syncTargets = [], remoteSyncMetadata = serializeSyncMetadataForRemote(syncMetadata), stateData = appState, collectionScopeMutations = {}) => {
  return syncTargets.reduce((updates, target) => {
    if (target.scopeKey === 'common') {
      updates['/shared_data/common'] = buildRemoteCommonPayload(stateData?.common || {});
      updates[`/${SHARED_META_ROOT}/common`] = hasUsableSyncMetaEntry(remoteSyncMetadata.common)
        ? normalizeSyncMetaEntry(remoteSyncMetadata.common)
        : null;
      return updates;
    }

    const monthKey = target.monthKey;
    const monthMeta = getSyncMetaMonthEntry(remoteSyncMetadata, monthKey);

    if (target.scopeKey === 'monthSettings') {
      Object.assign(updates, buildMonthSettingsDataUpdates(monthKey, getRemoteMonthScopeData(stateData?.months?.[monthKey] || {}, 'monthSettings')));
      const monthSettingsMetaEntry = normalizeSyncMetaEntry(monthMeta?.monthSettings);
      updates[`/${SHARED_META_ROOT}/months/${monthKey}/monthSettings/revision`] = hasUsableSyncMetaEntry(monthSettingsMetaEntry) ? monthSettingsMetaEntry.revision : null;
      updates[`/${SHARED_META_ROOT}/months/${monthKey}/monthSettings/updatedAt`] = hasUsableSyncMetaEntry(monthSettingsMetaEntry) ? monthSettingsMetaEntry.updatedAt : null;
      return updates;
    }

    if (target.scopeKey === 'monthSettings.commonSnapshot') {
      updates[`/shared_data/months/${monthKey}/monthSettings/commonSnapshot`] = getRemoteMonthScopeData(stateData?.months?.[monthKey] || {}, 'monthSettings.commonSnapshot');
      updates[`/${SHARED_META_ROOT}/months/${monthKey}/monthSettings/commonSnapshot`] = hasUsableSyncMetaEntry(monthMeta?.monthSettings?.commonSnapshot)
        ? normalizeSyncMetaEntry(monthMeta.monthSettings.commonSnapshot)
        : null;
      return updates;
    }

    const collectionMutation = collectionScopeMutations[getRemoteCollectionMutationKey(monthKey, target.scopeKey)] || null;
    if (COLLECTION_SYNC_SCOPE_KEYS.includes(target.scopeKey) && collectionMutation) {
      Object.assign(updates, buildCollectionScopeUpdatesFromMutation(monthKey, target.scopeKey, collectionMutation));
      updates[`/${SHARED_META_ROOT}/months/${monthKey}/${target.scopeKey}`] = hasUsableSyncMetaEntry(monthMeta?.[target.scopeKey])
        ? normalizeSyncMetaEntry(monthMeta[target.scopeKey])
        : null;
      return updates;
    }

    updates[`/shared_data/months/${monthKey}/${target.scopeKey}`] = getRemoteMonthScopeData(stateData?.months?.[monthKey] || {}, target.scopeKey);
    updates[`/${SHARED_META_ROOT}/months/${monthKey}/${target.scopeKey}`] = hasUsableSyncMetaEntry(monthMeta?.[target.scopeKey])
      ? normalizeSyncMetaEntry(monthMeta[target.scopeKey])
      : null;
    return updates;
  }, {});
};

const buildFirebaseRootUpdatesFromStateBundle = (stateData = {}, metaData = {}) => {
  const normalizedState = normalizeState(stateData || {});
  const normalizedMeta = serializeSyncMetadataForRemote(metaData || createDefaultSyncMeta());
  const syncTargets = [
    { scopeKey: 'common', monthKey: '' },
    ...Object.keys(normalizedState.months || {}).flatMap(monthKey => MONTH_SYNC_SCOPE_KEYS.map(scopeKey => ({ scopeKey, monthKey })))
  ];

  return {
    [`/${LEGACY_SHARED_INVOICE_TOTALS_ROOT}`]: null,
    ...buildRemoteUpdatesForSyncTargets(syncTargets, normalizedMeta, normalizedState, {})
  };
};

const applyTrackedChangePersistence = (monthKeysToSync = [], persistCommon = true, extraRemoteUpdates = {}, options = {}) => {
  persistCurrentStateToLocalStorage(monthKeysToSync, persistCommon);

  if (!window.isOfflineMode && typeof firebase !== 'undefined' && firebase.auth && firebase.auth().currentUser && !window.isImportingFromFirebase) {
    const remoteSyncMetadata = serializeSyncMetadataForRemote(syncMetadata);
    const syncTargets = Array.isArray(options.remoteSyncTargets) ? options.remoteSyncTargets : [
      ...(persistCommon ? [{ scopeKey: 'common', monthKey: '' }] : []),
      ...[...new Set(monthKeysToSync.filter(Boolean))].flatMap(monthKey => MONTH_SYNC_SCOPE_KEYS.map(scopeKey => ({ scopeKey, monthKey })))
    ];
    const updates = {
      [`/${LEGACY_SHARED_INVOICE_TOTALS_ROOT}`]: null,
      ...buildRemoteUpdatesForSyncTargets(syncTargets, remoteSyncMetadata, appState, options.collectionScopeMutations || {}),
      ...extraRemoteUpdates
    };

    if (Object.keys(updates).length > 0) {
      scheduleFirebaseRootUpdates(updates, { immediate: options.immediateRemoteSync === true });
    }
  }

  dispatchStateEvents();
};

const finalizeTrackedChangeContext = (monthToSync = '', saveOptions = {}) => {
  const transaction = activeMutationTransaction;
  activeMutationTransaction = null;
  if (!transaction) return false;

  const completedChanges = transaction.changes.map(change => {
    const afterSnapshot = cloneData(getScopeSnapshot(change.scope, change.month));
    return {
      ...change,
      afterSnapshot
    };
  }).filter(change => !deepEqual(change.beforeSnapshot, change.afterSnapshot));

  if (completedChanges.length === 0) {
    return false;
  }

  const timestamp = getIsoTimestamp();
  const author = getCurrentActorId();
  const updatedSyncScopes = new Map();
  completedChanges.forEach(change => {
    const syncScopeKey = getSyncScopeKey(change.scope);
    const scopeMonthKey = syncScopeKey === 'common' ? '' : getScopeMonthKey(change);
    const mapKey = syncScopeKey === 'common' ? 'common' : `${scopeMonthKey}:${syncScopeKey}`;
    if (updatedSyncScopes.has(mapKey)) {
      change.afterRevision = updatedSyncScopes.get(mapKey).revision;
      return;
    }

    const currentMeta = getSyncMetaEntry(change.scope, change.month);
    const nextRevision = hasStateDataForSyncScope(appState, syncScopeKey, scopeMonthKey)
      ? (currentMeta.revision + 1)
      : 0;
    const nextMeta = {
      revision: nextRevision,
      updatedAt: nextRevision > 0 ? timestamp : ''
    };
    setSyncMetaEntry(change.scope, change.month, nextMeta);
    const normalizedNextMeta = getSyncMetaEntry(change.scope, change.month);
    updatedSyncScopes.set(mapKey, normalizedNextMeta);
    change.afterRevision = normalizedNextMeta.revision;
  });

  const localHistoryEntry = {
    id: transaction.id,
    label: transaction.label,
    timestamp,
    author,
    month: transaction.primaryMonth,
    changes: completedChanges.map(change => ({
      scope: change.scope,
      scopeLabel: change.scopeLabel,
      month: change.month,
      beforeSnapshot: change.beforeSnapshot,
      afterSnapshot: change.afterSnapshot,
      beforeRevision: change.beforeRevision,
      afterRevision: change.afterRevision
    }))
  };

  if (!transaction.skipLocalHistory) {
    localUndoStack.push(localHistoryEntry);
    localUndoStack = localUndoStack.slice(-LOCAL_HISTORY_LIMIT);
    localRedoStack = [];
  }

  if (!transaction.skipSharedHistory) {
    const sharedHistoryResult = prependNormalizedEntriesWithLimit(sharedHistoryEntries, localHistoryEntry, SHARED_HISTORY_LIMIT, normalizeSharedHistoryEntries);
    sharedHistoryEntries = sharedHistoryResult.entries;
    saveOptions.removedSharedHistoryEntries = sharedHistoryResult.removedEntries;
  }

  if (transaction.dailyBackup) {
    const dailyBackupResult = prependNormalizedEntriesWithLimit(dailyBackupEntries, transaction.dailyBackup, SHARED_DAILY_BACKUPS_LIMIT, normalizeBackupEntries);
    dailyBackupEntries = dailyBackupResult.entries;
    saveOptions.removedDailyBackupEntries = dailyBackupResult.removedEntries;
    rebuildRecordedBackupKeys();
  }

  if (transaction.monthlyBackup) {
    const monthlyBackupResult = prependNormalizedEntriesWithLimit(monthlyBackupEntries, transaction.monthlyBackup, SHARED_MONTHLY_BACKUPS_LIMIT, normalizeBackupEntries);
    monthlyBackupEntries = monthlyBackupResult.entries;
    saveOptions.removedMonthlyBackupEntries = monthlyBackupResult.removedEntries;
    rebuildRecordedBackupKeys();
  }

  const monthKeysToSync = buildMonthKeysToPersist(completedChanges, monthToSync || transaction.primaryMonth);
  const persistCommon = transaction.persistCommon || saveOptions.persistCommon === true;
  const remoteSyncTargets = collectRemoteSyncTargetsFromChanges(completedChanges, persistCommon);
  const collectionScopeMutations = buildRemoteCollectionScopeMutations(completedChanges);
  const remoteUpdates = {
    ...(!transaction.skipSharedHistory
      ? buildIndexedRemoteUpdatesForEntries(
          SHARED_HISTORY_INDEX_ROOT,
          SHARED_HISTORY_ENTRIES_ROOT,
          [localHistoryEntry],
          saveOptions.removedSharedHistoryEntries || [],
          buildSharedHistoryIndexEntry
        )
      : {}),
    ...(transaction.dailyBackup
      ? buildIndexedRemoteUpdatesForEntries(
          SHARED_DAILY_BACKUPS_INDEX_ROOT,
          SHARED_DAILY_BACKUPS_ENTRIES_ROOT,
          [transaction.dailyBackup],
          saveOptions.removedDailyBackupEntries || [],
          buildBackupIndexEntry
        )
      : {}),
    ...(transaction.monthlyBackup
      ? buildIndexedRemoteUpdatesForEntries(
          SHARED_MONTHLY_BACKUPS_INDEX_ROOT,
          SHARED_MONTHLY_BACKUPS_ENTRIES_ROOT,
          [transaction.monthlyBackup],
          saveOptions.removedMonthlyBackupEntries || [],
          buildBackupIndexEntry
        )
      : {})
  };

  applyTrackedChangePersistence(monthKeysToSync, persistCommon, remoteUpdates, {
    collectionScopeMutations,
    remoteSyncTargets,
    immediateRemoteSync: transaction.forceImmediateRemoteSync === true || saveOptions.immediateRemoteSync === true
  });
  return true;
};

const persistAllMonthsToLocalStorage = () => {
  Object.keys(localStorage)
    .filter(key => key.startsWith(WORK_TRACKER_MONTH_PREFIX))
    .forEach(key => localStorage.removeItem(key));

  Object.entries(appState.months || {}).forEach(([monthKey, monthRecord]) => {
    localStorage.setItem(WORK_TRACKER_MONTH_PREFIX + monthKey, JSON.stringify(monthRecord));
  });
};

const loadPersistedMonthsFromLocalStorage = () => {
  const months = {};

  Object.keys(localStorage)
    .filter(key => key.startsWith(WORK_TRACKER_MONTH_PREFIX))
    .forEach(key => {
      const monthKey = key.slice(WORK_TRACKER_MONTH_PREFIX.length);
      if (!/^\d{4}-\d{2}$/.test(monthKey)) return;

      const persistedMonth = safeParseStorage(key);
      if (persistedMonth) {
        months[monthKey] = persistedMonth;
      }
    });

  return months;
};

const queueFirebaseRootUpdateFlush = () => {
  if (pendingFirebaseRootUpdateTimer) {
    window.clearTimeout(pendingFirebaseRootUpdateTimer);
  }

  pendingFirebaseRootUpdateTimer = window.setTimeout(() => {
    pendingFirebaseRootUpdateTimer = null;
    flushPendingFirebaseRootUpdates();
  }, FIREBASE_WRITE_DEBOUNCE_MS);
};

function performFirebaseRootUpdate(updates = {}) {
  if (Object.keys(updates).length === 0) return Promise.resolve();

  if (window.isOfflineMode || typeof firebase === 'undefined' || !firebase.auth || !firebase.auth().currentUser || window.isImportingFromFirebase) {
    return Promise.resolve();
  }

  recordFirebaseUpload(updates);
  return firebase.database().ref().update(updates).catch(error => {
    if (!isFirebasePermissionDeniedError(error)) {
      console.error('firebase: Błąd zapisu danych aplikacji:', error);
      return;
    }

    const legacyUpdates = buildSharedDataOnlyUpdates(updates);
    if (Object.keys(legacyUpdates).length === 0) {
      console.warn('firebase: Brak uprawnień do nowych gałęzi historii/meta i brak danych zapasowych do zapisania.', error);
      return;
    }

    console.warn('firebase: Brak uprawnień do nowych gałęzi historii/meta. Zapis przełączony na zgodność tylko z /shared_data.', error);
    recordFirebaseUpload(legacyUpdates);
    return firebase.database().ref().update(legacyUpdates).catch(legacyError => {
      console.error('firebase: Błąd zapisu danych zgodności do /shared_data:', legacyError);
    });
  });
}

function flushPendingFirebaseRootUpdates() {
  if (pendingFirebaseRootUpdateTimer) {
    window.clearTimeout(pendingFirebaseRootUpdateTimer);
    pendingFirebaseRootUpdateTimer = null;
  }

  const updates = { ...pendingFirebaseRootUpdates };
  pendingFirebaseRootUpdates = {};

  if (Object.keys(updates).length === 0) {
    return pendingFirebaseRootFlushPromise;
  }

  pendingFirebaseRootFlushPromise = pendingFirebaseRootFlushPromise
    .catch(() => {})
    .then(() => performFirebaseRootUpdate(updates));

  return pendingFirebaseRootFlushPromise;
}

const scheduleFirebaseRootUpdates = (updates = {}, options = {}) => {
  if (Object.keys(updates).length === 0) return Promise.resolve();

  pendingFirebaseRootUpdates = {
    ...pendingFirebaseRootUpdates,
    ...updates
  };

  if (options.immediate === true) {
    return flushPendingFirebaseRootUpdates();
  }

  queueFirebaseRootUpdateFlush();
  return Promise.resolve();
};

const isMonthArchived = (month) => getMonthRecord(month)?.monthSettings?.isArchived === true;

const canMutateMonth = (month) => !isMonthArchived(month);

// Inicjalizacja stanu
const persistedCommon = safeParseStorage(WORK_TRACKER_COMMON_KEY);
const persistedLocalSettings = safeParseStorage(WORK_TRACKER_LOCAL_SETTINGS_KEY);
const persistedSyncMeta = safeParseStorage(WORK_TRACKER_SYNC_META_KEY);
const persistedLegacy = safeParseStorage(WORK_TRACKER_LEGACY_KEY);

let localSettings = normalizeLocalSettings(persistedLocalSettings || createDefaultLocalSettings());
let appState;
let syncMetadata = normalizeSyncMetadata(persistedSyncMeta || createDefaultSyncMeta());
let firebaseTransferStats = createDefaultFirebaseTransferStats();
let pendingFirebaseRootUpdates = {};
let pendingFirebaseRootUpdateTimer = null;
let pendingFirebaseRootFlushPromise = Promise.resolve();

if (!persistedCommon && persistedLegacy) {
    appState = normalizeState(persistedLegacy);
} else {
    appState = normalizeState({
        version: 'v3',
        common: persistedCommon || {},
        months: loadPersistedMonthsFromLocalStorage()
    });
}

ensureMonthSettings(localSettings.selectedMonth);

const initialSyncMetadataBootstrap = mergeSyncMetadataWithState(appState, syncMetadata);
syncMetadata = initialSyncMetadataBootstrap.meta;

localStorage.setItem(WORK_TRACKER_COMMON_KEY, JSON.stringify(appState.common));
localStorage.setItem(WORK_TRACKER_LOCAL_SETTINGS_KEY, JSON.stringify(localSettings));
persistSyncMetadataToLocalStorage();
localStorage.removeItem(LEGACY_WORK_TRACKER_INVOICE_TOTALS_KEY);
if (appState.months[localSettings.selectedMonth]) {
    localStorage.setItem(WORK_TRACKER_MONTH_PREFIX + localSettings.selectedMonth, JSON.stringify(appState.months[localSettings.selectedMonth]));
}

const applyRemoteMonthSettingsScope = (monthKey = '', rawMonthSettings = {}) => {
  const monthSettings = ensureMonthSettings(monthKey);
  const normalizedMonthSettings = normalizeMonth({ monthSettings: rawMonthSettings || {} }).monthSettings;

  monthSettings.persons = cloneData(normalizedMonthSettings.persons || {});
  monthSettings.clients = cloneData(normalizedMonthSettings.clients || {});
  monthSettings.settlementConfig = cloneData(normalizedMonthSettings.settlementConfig || {});
  monthSettings.personContractCharges = cloneData(normalizedMonthSettings.personContractCharges || {});
  monthSettings.payouts = normalizePayoutSettings(normalizedMonthSettings.payouts || {});
  monthSettings.invoices = normalizeInvoicesConfig(normalizedMonthSettings.invoices || {});

  if (normalizedMonthSettings.isArchived === true) {
    monthSettings.isArchived = true;
  } else {
    delete monthSettings.isArchived;
  }

  ensureArchivedMonthHasSnapshot(appState.months[monthKey], appState.common);
  return monthSettings;
};

const applyRemoteMonthCommonSnapshotScope = (monthKey = '', rawCommonSnapshot = null) => {
  const monthSettings = ensureMonthSettings(monthKey);

  if (rawCommonSnapshot && typeof rawCommonSnapshot === 'object') {
    monthSettings.commonSnapshot = normalizeCommonData(rawCommonSnapshot, createDefaultState().common);
  } else {
    delete monthSettings.commonSnapshot;
  }

  ensureArchivedMonthHasSnapshot(appState.months[monthKey], appState.common);
  return monthSettings;
};

const applyRemoteMonthSyncScope = (monthKey = '', scopeKey = '', rawValue = null) => {
  ensureMonthSettings(monthKey);

  switch (scopeKey) {
    case 'monthlySheets':
      appState.months[monthKey].monthlySheets = normalizeMonth({ monthlySheets: rawValue }).monthlySheets;
      return true;
    case 'worksSheets':
      appState.months[monthKey].worksSheets = normalizeMonth({ worksSheets: rawValue }).worksSheets;
      return true;
    case 'expenses':
      appState.months[monthKey].expenses = normalizeMonth({ expenses: rawValue }).expenses;
      return true;
    case 'monthSettings':
      applyRemoteMonthSettingsScope(monthKey, rawValue || {});
      return true;
    case 'monthSettings.commonSnapshot':
      applyRemoteMonthCommonSnapshotScope(monthKey, rawValue);
      return true;
    default:
      return false;
  }
};

const remoteMonthDataNeedsMigration = (rawMonthData = {}) => Array.isArray(rawMonthData?.monthlySheets)
  || Array.isArray(rawMonthData?.worksSheets)
  || Array.isArray(rawMonthData?.expenses);

const Store = {
  getState: () => {
    const monthKey = Store.getSelectedMonth();
    const m = appState.months[monthKey] || { monthlySheets: [], worksSheets: [], expenses: [], monthSettings: {} };
    const isArchived = m.monthSettings?.isArchived === true;
    const common = getCommonDataForMonth(monthKey);

    return {
      selectedMonth: monthKey,
      isArchived: isArchived,
      hasCommonSnapshot: !!m.monthSettings?.commonSnapshot,
      persons: common.persons,
      clients: common.clients,
      worksCatalog: common.worksCatalog,
      config: common.config,
      settings: {},
      monthlySheets: m.monthlySheets,
      worksSheets: m.worksSheets,
      expenses: m.expenses,
      monthSettings: { [monthKey]: m.monthSettings }
    };
  },
  getStateForMonth: (monthKey) => {
    const m = appState.months[monthKey] || { monthlySheets: [], worksSheets: [], expenses: [], monthSettings: {} };
    const isArchived = m.monthSettings?.isArchived === true;
    const common = getCommonDataForMonth(monthKey);

    return {
      selectedMonth: monthKey,
      isArchived: isArchived,
      hasCommonSnapshot: !!m.monthSettings?.commonSnapshot,
      persons: common.persons,
      clients: common.clients,
      worksCatalog: common.worksCatalog,
      config: common.config,
      settings: {},
      monthlySheets: m.monthlySheets,
      worksSheets: m.worksSheets,
      expenses: m.expenses,
      monthSettings: { [monthKey]: m.monthSettings }
    };
  },
  getSettings: () => ({ ...localSettings }),
  getAppearanceSettings: () => ({ ...localSettings }),
  getSelectedMonth: () => localSettings.selectedMonth,
  getExportData: () => JSON.parse(JSON.stringify(appState)),
  getDefaultCatalogItems: () => DEFAULT_WORKS_CATALOG.map(item => ({ ...item })),
  getDefaultCatalogItem: (coreId) => {
    const item = DEFAULT_WORKS_CATALOG.find(entry => entry.coreId === coreId);
    return item ? { ...item } : null;
  },
  getMonthArchiveStatus: (month = Store.getSelectedMonth()) => {
    const monthRecord = getMonthRecord(month);
    const monthSettings = monthRecord?.monthSettings || {};
    return {
      month,
      isArchived: monthSettings.isArchived === true,
      hasSnapshot: !!monthSettings.commonSnapshot,
      hasData: !!monthRecord && (
        (monthRecord.monthlySheets || []).length > 0
        || (monthRecord.worksSheets || []).length > 0
        || (monthRecord.expenses || []).length > 0
        || Object.keys(monthSettings.persons || {}).length > 0
        || Object.keys(monthSettings.clients || {}).length > 0
        || Object.keys(monthSettings.settlementConfig || {}).length > 0
        || Object.keys(monthSettings.personContractCharges || {}).length > 0
        || Object.keys(monthSettings.invoices?.clients || {}).length > 0
        || (monthSettings.invoices?.extraInvoices || []).length > 0
      )
    };
  },
  getKnownMonths: () => Object.keys(appState.months || {}).sort((a, b) => b.localeCompare(a, 'pl-PL')),
  getSyncMetadata: () => cloneData(syncMetadata),
  getFirebaseTransferStats: () => ({ ...firebaseTransferStats }),
  recordFirebaseDownload: (payload = null) => {
    recordFirebaseDownload(payload);
    return true;
  },
  recordFirebaseUpload: (payload = null) => {
    recordFirebaseUpload(payload);
    return true;
  },
  resetFirebaseTransferStats: () => {
    firebaseTransferStats = createDefaultFirebaseTransferStats();
    dispatchFirebaseTransferStatsChanged();
    return true;
  },
  flushPendingRemoteSync: () => flushPendingFirebaseRootUpdates(),
  normalizeSyncMetadata: (meta = {}) => normalizeSyncMetadata(meta),
  serializeSyncMetadataForRemote: (meta = {}) => serializeSyncMetadataForRemote(meta),
  serializeRemoteCommon: (common = {}) => buildRemoteCommonPayload(common),
  serializeRemoteMonth: (monthRecord = {}) => serializeRemoteMonthRecord(monthRecord),
  buildFirebaseRootUpdatesFromStateBundle: (state = {}, meta = {}) => buildFirebaseRootUpdatesFromStateBundle(state, meta),
  remoteMonthDataNeedsMigration: (rawMonthData = {}) => remoteMonthDataNeedsMigration(rawMonthData),
  buildSyncMetadataFromState: (state) => rebuildSyncMetadataFromState(state),
  ensureSyncMetadataBootstrap: (state = appState, author = getCurrentActorId(), timestamp = getIsoTimestamp()) => {
    const normalizedState = normalizeState(state);
    const result = mergeSyncMetadataWithState(normalizedState, syncMetadata, author, timestamp);
    syncMetadata = result.meta;
    persistSyncMetadataToLocalStorage();
    return {
      changed: result.changed,
      meta: cloneData(syncMetadata)
    };
  },
  buildSyncConflictInfo: (remoteMeta = {}) => {
    const conflicts = collectSyncConflictEntries(syncMetadata, normalizeSyncMetadata(remoteMeta));
    return {
      hasLocalNewerChanges: conflicts.localNewer.length > 0,
      hasRemoteNewerChanges: conflicts.remoteNewer.length > 0,
      localNewer: cloneData(conflicts.localNewer),
      remoteNewer: cloneData(conflicts.remoteNewer)
    };
  },
  recordSharedSyncOverwriteHistory: (beforeState = {}, afterState = {}, beforeMeta = {}, afterMeta = {}, label = 'Nadpisano bazę Firebase bazą lokalną') => {
    const entry = buildSharedHistoryEntryFromStateTransition(beforeState, afterState, beforeMeta, afterMeta, label);
    if (!entry) {
      return { success: false, history: cloneData(sharedHistoryEntries), entry: null, removedEntries: [] };
    }

    const prependResult = prependSharedHistoryEntry(entry);
    dispatchStateEvents();
    return {
      success: true,
      history: cloneData(sharedHistoryEntries),
      entry: cloneData(entry),
      removedEntries: cloneData(prependResult.removedEntries || [])
    };
  },
  shouldFetchRemoteCommon: (remoteMeta = {}) => shouldFetchRemoteCommon(normalizeSyncMetadata(remoteMeta)),
  shouldFetchRemoteMonth: (month = Store.getSelectedMonth(), remoteMeta = {}) => shouldFetchRemoteMonth(month, normalizeSyncMetadata(remoteMeta)),
  getRemoteMonthSyncScopesToFetch: (month = Store.getSelectedMonth(), remoteMeta = {}) => getRemoteMonthSyncScopesToFetch(month, normalizeSyncMetadata(remoteMeta)),
  applyRemoteCommon: (remoteCommon = {}, remoteMeta = {}) => {
    appState.common = normalizeCommonData(remoteCommon, createDefaultState().common);
    const nextMeta = cloneData(syncMetadata);
    nextMeta.common = normalizeSyncMetadata({ common: remoteMeta.common || {} }).common;
    syncMetadata = mergeSyncMetadataWithState(appState, nextMeta).meta;
    persistCurrentStateToLocalStorage([], true);
    dispatchStateEvents();
    return true;
  },
  applyRemoteMonthScope: (month = Store.getSelectedMonth(), scopeKey = '', remoteValue = null, remoteMeta = {}) => {
    applyRemoteMonthSyncScope(month, scopeKey, remoteValue);

    const normalizedRemoteMeta = normalizeSyncMetadata({ months: { [month]: remoteMeta?.months?.[month] || remoteMeta || {} } });
    const remoteMonthMeta = getSyncMetaMonthEntry(normalizedRemoteMeta, month);

    if (scopeKey === 'monthSettings.commonSnapshot') {
      setSyncMetaEntry(`month.${scopeKey}`, month, remoteMonthMeta?.monthSettings?.commonSnapshot || {});
    } else if (scopeKey === 'monthSettings') {
      setSyncMetaEntry(`month.${scopeKey}`, month, remoteMonthMeta?.monthSettings || {});
    } else {
      setSyncMetaEntry(`month.${scopeKey}`, month, remoteMonthMeta?.[scopeKey] || {});
    }

    syncMetadata = mergeSyncMetadataWithState(appState, syncMetadata).meta;
    persistCurrentStateToLocalStorage([month], false);
    dispatchStateEvents();
    return true;
  },
  applyRemoteMonth: (month = Store.getSelectedMonth(), remoteMonth = {}, remoteMeta = {}) => {
    appState.months[month] = normalizeMonth(remoteMonth);
    ensureArchivedMonthHasSnapshot(appState.months[month], appState.common);
    const nextMeta = cloneData(syncMetadata);
    if (!nextMeta.months[month]) nextMeta.months[month] = {};
    nextMeta.months[month] = normalizeSyncMetadata({ months: { [month]: remoteMeta?.months?.[month] || remoteMeta || {} } }).months[month] || {};
    syncMetadata = mergeSyncMetadataWithState(appState, nextMeta).meta;
    persistCurrentStateToLocalStorage([month], false);
    dispatchStateEvents();
    return true;
  },
  applyRemoteStateBundle: (remoteState = {}, remoteMeta = {}) => {
    appState = normalizeState(remoteState);
    const mergedMeta = mergeSyncMetadataWithState(appState, normalizeSyncMetadata(remoteMeta));
    syncMetadata = mergedMeta.meta;
    localUndoStack = [];
    localRedoStack = [];
    persistAllMonthsToLocalStorage();
    localStorage.setItem(WORK_TRACKER_COMMON_KEY, JSON.stringify(appState.common));
    persistSyncMetadataToLocalStorage();
    dispatchStateEvents();
    return true;
  },
  clearLocalDatabaseStorage: (options = {}) => {
    const preserveLocalSettings = options.preserveLocalSettings !== false;
    const preservedLocalSettings = preserveLocalSettings
      ? localStorage.getItem(WORK_TRACKER_LOCAL_SETTINGS_KEY)
      : null;

    clearPersistedStateFromLocalStorage(preservedLocalSettings);

    appState = createDefaultState();
    syncMetadata = createDefaultSyncMeta();
    localUndoStack = [];
    localRedoStack = [];
    sharedHistoryEntries = [];
    dailyBackupEntries = [];
    monthlyBackupEntries = [];
    rebuildRecordedBackupKeys();
    localStorage.removeItem(LEGACY_WORK_TRACKER_INVOICE_TOTALS_KEY);

    if (options.dispatchStateChanged === true) {
      dispatchStateEvents();
    }

    return true;
  },
  setSharedHistoryCaches: ({ history, dailyBackups, monthlyBackups } = {}, options = {}) => {
    if (history !== undefined) {
      sharedHistoryEntries = normalizeSharedHistoryEntries(history).slice(0, SHARED_HISTORY_LIMIT);
    }
    if (dailyBackups !== undefined) {
      dailyBackupEntries = normalizeBackupEntries(dailyBackups);
    }
    if (monthlyBackups !== undefined) {
      monthlyBackupEntries = normalizeBackupEntries(monthlyBackups);
    }
    rebuildRecordedBackupKeys();
    if (options.skipPeriodicBackups !== true) {
      ensurePeriodicBackupsForLogin();
    }
    dispatchStateEvents();
  },
  getSharedHistoryEntry: (entryId = '') => {
    const entry = sharedHistoryEntries.find(item => item.id === entryId) || null;
    return entry ? cloneData(entry) : null;
  },
  getBackupEntry: (type = 'daily', entryId = '') => {
    const sourceEntries = type === 'monthly' ? monthlyBackupEntries : dailyBackupEntries;
    const entry = sourceEntries.find(item => item.id === entryId) || null;
    return entry ? cloneData(entry) : null;
  },
  cacheSharedHistoryEntryDetail: (entry = {}) => {
    if (!entry?.id) return false;

    const nextEntry = normalizeSharedHistoryEntries([entry])[0];
    if (!nextEntry) return false;

    sharedHistoryEntries = prependNormalizedEntriesWithLimit(
      sharedHistoryEntries.filter(item => item.id !== nextEntry.id),
      nextEntry,
      SHARED_HISTORY_LIMIT,
      normalizeSharedHistoryEntries
    ).entries;
    dispatchStateEvents();
    return true;
  },
  cacheBackupEntryDetail: (type = 'daily', entry = {}) => {
    if (!entry?.id) return false;

    const nextEntry = normalizeBackupEntries([entry])[0];
    if (!nextEntry) return false;

    if (type === 'monthly') {
      monthlyBackupEntries = normalizeBackupEntries([
        nextEntry,
        ...monthlyBackupEntries.filter(item => item.id !== nextEntry.id)
      ]);
    } else {
      dailyBackupEntries = normalizeBackupEntries([
        nextEntry,
        ...dailyBackupEntries.filter(item => item.id !== nextEntry.id)
      ]);
    }

    rebuildRecordedBackupKeys();
    dispatchStateEvents();
    return true;
  },
  buildSharedHistoryMigrationUpdates: (entries = [], removedEntries = []) => {
    const normalizedEntries = normalizeSharedHistoryEntries(entries);
    const explicitRemovedEntries = normalizeSharedHistoryEntries(removedEntries);

    return buildIndexedRemoteUpdatesForEntries(
      SHARED_HISTORY_INDEX_ROOT,
      SHARED_HISTORY_ENTRIES_ROOT,
      normalizedEntries.slice(0, SHARED_HISTORY_LIMIT),
      [...normalizedEntries.slice(SHARED_HISTORY_LIMIT), ...explicitRemovedEntries],
      buildSharedHistoryIndexEntry
    );
  },
  buildBackupMigrationUpdates: (type = 'daily', entries = []) => buildIndexedRemoteUpsertUpdatesFromCollection(
    entries,
    type === 'monthly' ? SHARED_MONTHLY_BACKUPS_INDEX_ROOT : SHARED_DAILY_BACKUPS_INDEX_ROOT,
    type === 'monthly' ? SHARED_MONTHLY_BACKUPS_ENTRIES_ROOT : SHARED_DAILY_BACKUPS_ENTRIES_ROOT,
    buildBackupIndexEntry,
    normalizeBackupEntries
  ),
  ensurePeriodicBackupsForLogin: () => ensurePeriodicBackupsForLogin(),
  createOrReplaceDailyBackup: () => replaceHistoryBackupEntry('daily'),
  createOrReplaceMonthlyBackup: () => replaceHistoryBackupEntry('monthly'),
  getHistoryState: () => ({
    canUndo: localUndoStack.length > 0,
    canRedo: localRedoStack.length > 0,
    undoCount: localUndoStack.length,
    redoCount: localRedoStack.length,
    localUndoEntries: cloneData([...localUndoStack].reverse()),
    localRedoEntries: cloneData([...localRedoStack].reverse()),
    sharedEntries: cloneData(sharedHistoryEntries),
    dailyBackups: cloneData(dailyBackupEntries),
    monthlyBackups: cloneData(monthlyBackupEntries)
  }),

  setSelectedMonth: (month) => {
    flushPendingFirebaseRootUpdates();
    localSettings.selectedMonth = month || getCurrentMonthKey();
    ensureMonthSettings(localSettings.selectedMonth);
    Store.saveLocalSettings();
    window.dispatchEvent(new CustomEvent('monthChanged', { detail: { month: localSettings.selectedMonth } }));
  },

  getMonthSettings: (month = Store.getSelectedMonth()) => ensureMonthSettings(month),
  isMonthArchived: (month = Store.getSelectedMonth()) => isMonthArchived(month),

  isPersonActiveInMonth: (personId, month = Store.getSelectedMonth()) => {
    const person = (getCommonDataForMonth(month).persons || []).find(p => p.id === personId);
    if (!person) return false;
    const monthSettings = ensureMonthSettings(month);
    if (monthSettings.persons[personId] !== undefined) return monthSettings.persons[personId] !== false;
    return person.isActive !== false;
  },

  isClientActiveInMonth: (clientId, month = Store.getSelectedMonth()) => {
    const client = (getCommonDataForMonth(month).clients || []).find(c => c.id === clientId);
    if (!client) return false;
    const monthSettings = ensureMonthSettings(month);
    if (monthSettings.clients[clientId] !== undefined) return monthSettings.clients[clientId] !== false;
    return client.isActive !== false;
  },
  
  save: (monthToSync = Store.getSelectedMonth()) => {
    if (activeMutationTransaction) {
      return finalizeTrackedChangeContext(monthToSync);
    }

    const selectedMonthKey = Store.getSelectedMonth();
    const monthKeysToSync = [...new Set([selectedMonthKey, monthToSync].filter(Boolean))];
    applyTrackedChangePersistence(monthKeysToSync, true);
    return true;
  },

  saveLocalSettings: () => {
    localStorage.setItem(WORK_TRACKER_LOCAL_SETTINGS_KEY, JSON.stringify(localSettings));
    window.dispatchEvent(new Event('appStateChanged'));
  },

  updateAppearanceSettings: (newSettings) => {
    Object.assign(localSettings, newSettings);
    Store.saveLocalSettings();
  },
  resetAppearanceSettings: (nextSettings = createDefaultLocalSettings()) => {
    localSettings = normalizeLocalSettings({ ...createDefaultLocalSettings(), ...nextSettings });
    ensureMonthSettings(localSettings.selectedMonth);
    Store.saveLocalSettings();
    return true;
  },

  // Persons
  addPerson: (person) => {
    if (!canMutateMonth(Store.getSelectedMonth())) return false;
    const common = getMutableCommonDataForMonth(Store.getSelectedMonth());
    common.persons.push(normalizePersonRecord({ ...person, isActive: true }));
    Store.save();
    return true;
  },
  updatePerson: (id, updates) => {
    if (!canMutateMonth(Store.getSelectedMonth())) return false;
    const common = getMutableCommonDataForMonth(Store.getSelectedMonth());
    const idx = common.persons.findIndex(p => p.id === id);
    if (idx !== -1) {
      common.persons[idx] = normalizePersonRecord(updates, common.persons[idx]);
      Store.save();
      return true;
    }
    return false;
  },
  deletePerson: (id) => {
    if (!canMutateMonth(Store.getSelectedMonth())) return false;
    const month = Store.getSelectedMonth();
    const common = getMutableCommonDataForMonth(Store.getSelectedMonth());
    common.persons = common.persons.filter(p => p.id !== id);
    const cleanupResult = cleanupOrphanedPersonEntriesInMonth(month, [id]);
    Store.save();
    return { success: true, cleanupResult };
  },
  reorderPersons: (newOrderedIds) => {
    if (!canMutateMonth(Store.getSelectedMonth())) return false;
    const common = getMutableCommonDataForMonth(Store.getSelectedMonth());
    const reordered = newOrderedIds.map(id => common.persons.find(p => p.id === id)).filter(Boolean);
    const missing = common.persons.filter(p => !newOrderedIds.includes(p.id));
    common.persons = [...reordered, ...missing];
    Store.save();
    return true;
  },
  togglePersonStatus: (id, month = Store.getSelectedMonth()) => {
    if (!canMutateMonth(month)) return false;
    ensureMonthSettings(month).persons[id] = !Store.isPersonActiveInMonth(id, month);
    Store.save(month);
    return true;
  },
  cleanupOrphanedPersonEntries: (month = Store.getSelectedMonth()) => {
    if (!canMutateMonth(month)) {
      return { success: false, changed: false, month, message: 'Nie można czyścić osieroconych wpisów w zarchiwizowanym miesiącu.' };
    }

    const cleanupResult = cleanupOrphanedPersonEntriesInMonth(month);
    if (!cleanupResult.changed) return cleanupResult;

    Store.save(month);
    return cleanupResult;
  },

  // Sheets
  addMonthlySheet: (sheet) => {
    const month = sheet?.month || Store.getSelectedMonth();
    if (!canMutateMonth(month)) return false;
    const activePersons = Array.isArray(sheet?.activePersons) ? sheet.activePersons : getActivePersonIdsForMonthFromState(appState, month);
    ensureMonthSettings(month);
    appState.months[month].monthlySheets.push({ ...sheet, id: generateId(), activePersons });
    Store.save(month);
    return true;
  },
  updateMonthlySheet: (id, updates) => {
    const month = Store.getSelectedMonth();
    if (!canMutateMonth(month)) return false;
    const m = appState.months[month];
    const idx = m.monthlySheets.findIndex(s => s.id === id);
    if (idx !== -1) { m.monthlySheets[idx] = { ...m.monthlySheets[idx], ...updates }; Store.save(month); return true; }
    return false;
  },
  deleteMonthlySheet: (id) => {
    const month = Store.getSelectedMonth();
    if (!canMutateMonth(month)) return false;
    const m = appState.months[month];
    m.monthlySheets = m.monthlySheets.filter(s => s.id !== id);
    Store.save(month);
    return true;
  },
  reorderMonthlySheets: (newOrderedIds, month = Store.getSelectedMonth()) => {
    if (!canMutateMonth(month)) return false;
    const m = appState.months[month];
    if (!m) return false;

    const reordered = newOrderedIds.map(id => m.monthlySheets.find(sheet => sheet.id === id)).filter(Boolean);
    const missing = m.monthlySheets.filter(sheet => !newOrderedIds.includes(sheet.id));
    m.monthlySheets = [...reordered, ...missing];
    Store.save(month);
    return true;
  },
  getMonthlySheet: (id) => {
    const sheet = appState.months[Store.getSelectedMonth()]?.monthlySheets.find(s => s.id === id) || null;
    return sheet ? cloneData(sheet) : null;
  },

  // Works
  addWorksSheet: (sheet) => {
    const month = sheet?.month || Store.getSelectedMonth();
    if (!canMutateMonth(month)) return false;
    ensureMonthSettings(month);
    appState.months[month].worksSheets.push({ ...sheet, id: generateId(), entries: [] });
    Store.save(month);
    return true;
  },
  updateWorksSheet: (id, updates) => {
    const month = Store.getSelectedMonth();
    if (!canMutateMonth(month)) return false;
    const m = appState.months[month];
    const idx = m.worksSheets.findIndex(s => s.id === id);
    if (idx !== -1) { m.worksSheets[idx] = { ...m.worksSheets[idx], ...updates }; Store.save(month); return true; }
    return false;
  },
  deleteWorksSheet: (id) => {
    const month = Store.getSelectedMonth();
    if (!canMutateMonth(month)) return false;
    const m = appState.months[month];
    m.worksSheets = m.worksSheets.filter(s => s.id !== id);
    Store.save(month);
    return true;
  },
  reorderWorksSheets: (newOrderedIds, month = Store.getSelectedMonth()) => {
    if (!canMutateMonth(month)) return false;
    const m = appState.months[month];
    if (!m) return false;

    const reordered = newOrderedIds.map(id => m.worksSheets.find(sheet => sheet.id === id)).filter(Boolean);
    const missing = m.worksSheets.filter(sheet => !newOrderedIds.includes(sheet.id));
    m.worksSheets = [...reordered, ...missing];
    Store.save(month);
    return true;
  },
  getWorksSheet: (id) => {
    const sheet = appState.months[Store.getSelectedMonth()]?.worksSheets.find(s => s.id === id) || null;
    return sheet ? cloneData(sheet) : null;
  },
  reorderWorksSheetEntriesForDate: (sheetId, date, newOrderedIds, month = Store.getSelectedMonth()) => {
    if (!canMutateMonth(month) || !sheetId || !date) return false;
    const m = appState.months[month];
    if (!m) return false;

    const sheet = m.worksSheets.find(item => item.id === sheetId);
    if (!sheet || !Array.isArray(sheet.entries)) return false;

    const sameDateEntries = sheet.entries.filter(entry => entry?.date === date);
    const otherEntries = sheet.entries.filter(entry => entry?.date !== date);
    const reordered = newOrderedIds.map(id => sameDateEntries.find(entry => entry.id === id)).filter(Boolean);
    const missing = sameDateEntries.filter(entry => !newOrderedIds.includes(entry.id));
    sheet.entries = [...otherEntries, ...reordered, ...missing];
    Store.save(month);
    return true;
  },

  // Catalog
  addWorkToCatalog: (work) => {
    if (!canMutateMonth(Store.getSelectedMonth())) return false;
    getMutableCommonDataForMonth(Store.getSelectedMonth()).worksCatalog.push({ ...work, id: generateId() });
    Store.save();
    return true;
  },
  updateWorkInCatalog: (id, updates) => {
    if (!canMutateMonth(Store.getSelectedMonth())) return false;
    const common = getMutableCommonDataForMonth(Store.getSelectedMonth());
    const idx = common.worksCatalog.findIndex(w => w.id === id);
    if (idx !== -1) { common.worksCatalog[idx] = { ...common.worksCatalog[idx], ...updates }; Store.save(); return true; }
    return false;
  },
  deleteWorkFromCatalog: (id) => {
    if (!canMutateMonth(Store.getSelectedMonth())) return false;
    const common = getMutableCommonDataForMonth(Store.getSelectedMonth());
    common.worksCatalog = common.worksCatalog.filter(w => w.id !== id);
    Store.save();
    return true;
  },
  restoreWorkInCatalogToDefault: (id) => {
    if (!canMutateMonth(Store.getSelectedMonth())) return false;
    const common = getMutableCommonDataForMonth(Store.getSelectedMonth());
    const idx = common.worksCatalog.findIndex(w => w.id === id);
    if (idx === -1) return false;

    const current = common.worksCatalog[idx];
    if (!current?.coreId) return false;

    const defaultItem = Store.getDefaultCatalogItem(current.coreId);
    if (!defaultItem) return false;

    common.worksCatalog[idx] = {
      ...current,
      name: defaultItem.name,
      unit: defaultItem.unit,
      defaultPrice: defaultItem.defaultPrice,
      coreId: defaultItem.coreId
    };
    Store.save();
    return true;
  },
  restoreDeletedDefaultToCatalog: (coreId) => {
    if (!canMutateMonth(Store.getSelectedMonth()) || !coreId) return false;
    const common = getMutableCommonDataForMonth(Store.getSelectedMonth());
    if (common.worksCatalog.some(w => w.coreId === coreId)) return false;

    const defaultItem = Store.getDefaultCatalogItem(coreId);
    if (!defaultItem) return false;

    common.worksCatalog.push({ ...defaultItem, id: generateId() });
    Store.save();
    return true;
  },

  // Expenses
  addExpense: (expense) => {
    const month = Store.getSelectedMonth();
    if (!canMutateMonth(month)) return false;
    ensureMonthSettings(month);
    appState.months[month].expenses.push(normalizeExpenseRecord(expense));
    Store.save(month);
    return true;
  },
  updateExpense: (id, updates) => {
    const month = Store.getSelectedMonth();
    if (!canMutateMonth(month)) return false;
    const m = appState.months[month];
    const idx = m.expenses.findIndex(e => e.id === id);
    if (idx !== -1) { m.expenses[idx] = normalizeExpenseRecord(updates, m.expenses[idx]); Store.save(month); return true; }
    return false;
  },
  deleteExpense: (id) => {
    const month = Store.getSelectedMonth();
    if (!canMutateMonth(month)) return false;
    const m = appState.months[month];
    m.expenses = m.expenses.filter(e => e.id !== id);
    Store.save(month);
    return true;
  },
  reorderExpensesForDate: (date, newOrderedIds, month = Store.getSelectedMonth()) => {
    if (!canMutateMonth(month) || !date) return false;
    const m = appState.months[month];
    if (!m) return false;

    const sameDateExpenses = m.expenses.filter(expense => expense?.date === date);
    const otherExpenses = m.expenses.filter(expense => expense?.date !== date);
    const reordered = newOrderedIds.map(id => sameDateExpenses.find(expense => expense.id === id)).filter(Boolean);
    const missing = sameDateExpenses.filter(expense => !newOrderedIds.includes(expense.id));
    m.expenses = [...otherExpenses, ...reordered, ...missing];
    Store.save(month);
    return true;
  },

  // Clients
  addClient: (client) => {
    if (!canMutateMonth(Store.getSelectedMonth())) return false;
    const common = getMutableCommonDataForMonth(Store.getSelectedMonth());
    common.clients.push(normalizeClientRecord({ ...client, isActive: true }));
    Store.save();
    return true;
  },
  updateClient: (id, updates) => {
    if (!canMutateMonth(Store.getSelectedMonth())) return false;
    const common = getMutableCommonDataForMonth(Store.getSelectedMonth());
    const idx = common.clients.findIndex(c => c.id === id);
    if (idx !== -1) { common.clients[idx] = normalizeClientRecord(updates, common.clients[idx]); Store.save(); return true; }
    return false;
  },
  deleteClient: (id) => {
    if (!canMutateMonth(Store.getSelectedMonth())) return false;
    const common = getMutableCommonDataForMonth(Store.getSelectedMonth());
    common.clients = common.clients.filter(c => c.id !== id);
    Store.save();
    return true;
  },
  toggleClientStatus: (id, month = Store.getSelectedMonth()) => {
    if (!canMutateMonth(month)) return false;
    ensureMonthSettings(month).clients[id] = !Store.isClientActiveInMonth(id, month);
    Store.save(month);
    return true;
  },
  reorderClients: (newOrderedIds) => {
    if (!canMutateMonth(Store.getSelectedMonth())) return false;
    const common = getMutableCommonDataForMonth(Store.getSelectedMonth());
    const reordered = newOrderedIds.map(id => common.clients.find(c => c.id === id)).filter(Boolean);
    const missing = common.clients.filter(c => !newOrderedIds.includes(c.id));
    common.clients = [...reordered, ...missing];
    Store.save();
    return true;
  },

  // Config
  updateConfig: (newConfig) => {
    if (!canMutateMonth(Store.getSelectedMonth())) return false;
    const common = getMutableCommonDataForMonth(Store.getSelectedMonth());
    common.config = { ...common.config, ...newConfig };
    Store.save();
    return true;
  },
  updateSettlementMonthConfig: (monthConfig = {}, personContractCharges = {}, month = Store.getSelectedMonth()) => {
    if (!canMutateMonth(month)) return false;
    const ms = ensureMonthSettings(month);
    if (monthConfig.taxRate !== undefined) ms.settlementConfig.taxRate = parseFloat(monthConfig.taxRate);
    if (monthConfig.zusFixedAmount !== undefined) ms.settlementConfig.zusFixedAmount = parseFloat(monthConfig.zusFixedAmount);
    ms.personContractCharges = personContractCharges;
    Store.save(month);
    return true;
  },
  getPayoutSettings: (month = Store.getSelectedMonth()) => normalizePayoutSettings(ensureMonthSettings(month).payouts || {}),
  updatePayoutMonthConfig: (payoutConfig = {}, month = Store.getSelectedMonth()) => {
    if (!canMutateMonth(month)) return false;
    const ms = ensureMonthSettings(month);
    ms.payouts = normalizePayoutSettings({ ...(ms.payouts || {}), ...payoutConfig });
    Store.save(month);
    return true;
  },
  updatePayoutEmployeeConfig: (personId, employeeConfig = {}, month = Store.getSelectedMonth()) => {
    const normalizedPersonId = (personId || '').toString().trim();
    if (!canMutateMonth(month) || !normalizedPersonId) return false;

    const ms = ensureMonthSettings(month);
    const existingRecord = normalizePayoutEmployeeRecord(ms.payouts?.employees?.[normalizedPersonId] || {});
    ms.payouts = normalizePayoutSettings({
      ...(ms.payouts || {}),
      employees: {
        ...(ms.payouts?.employees || {}),
        [normalizedPersonId]: {
          ...existingRecord,
          ...employeeConfig
        }
      }
    });
    Store.save(month);
    return true;
  },
  settleEmployeePayout: (personId, settlement = {}, month = Store.getSelectedMonth()) => {
    const normalizedPersonId = (personId || '').toString().trim();
    if (!canMutateMonth(month) || !normalizedPersonId) return false;

    const ms = ensureMonthSettings(month);
    const existingRecord = normalizePayoutEmployeeRecord(ms.payouts?.employees?.[normalizedPersonId] || {});
    const monthRecord = appState.months[month] || (appState.months[month] = normalizeMonth(null));

    const selectedAdvanceExpenseIds = Array.isArray(settlement.advanceExpenseIds)
      ? [...new Set(settlement.advanceExpenseIds.map(v => (v || '').toString().trim()).filter(Boolean))]
      : [];
    const alreadyRemovedIds = new Set(existingRecord.removedAdvanceExpenseIds || []);

    // Advances to deduct — archive them before deletion
    const advancesToDeduct = (monthRecord.expenses || []).filter(expense =>
      expense?.type === 'ADVANCE' &&
      selectedAdvanceExpenseIds.includes(expense.id) &&
      !alreadyRemovedIds.has(expense.id)
    );

    // Default payer = person's employer if not specified
    const allPersons = appState.common?.persons || appState.persons || [];
    const personRecord = allPersons.find(p => p.id === normalizedPersonId);
    const employerPartnerId = (personRecord?.employerId || '').toString().trim();
    const paidByPartnerId = (settlement.paidByPartnerId || '').toString().trim() || employerPartnerId;

    const deductedAdvances = advancesToDeduct.map(expense => ({
      id: expense.id,
      date: expense.date || '',
      name: expense.name || '',
      amount: Math.max(0, parseFloat(expense.amount) || 0),
      paidById: expense.paidById || '',
      restoredToCosts: false,
      restoredAt: ''
    }));

    // Refund obligations: if payer differs from advance giver
    const advanceRefunds = deductedAdvances
      .filter(adv => adv.paidById && paidByPartnerId && adv.paidById !== paidByPartnerId)
      .map(adv => ({
        advanceId: adv.id,
        toPartnerId: adv.paidById,
        amount: adv.amount,
        returned: false,
        returnedAt: ''
      }));

    // Remove deducted advances from expenses
    const deductedIds = new Set(advancesToDeduct.map(e => e.id));
    if (deductedIds.size > 0) {
      monthRecord.expenses = (monthRecord.expenses || []).filter(e => !deductedIds.has(e.id));
    }

    const cashAmount = Math.max(0, parseFloat(settlement.cashAmount) || 0);

    // Cross-partner salary refund: if payer differs from employer, employer owes payer
    const salaryRefund = (paidByPartnerId && employerPartnerId && paidByPartnerId !== employerPartnerId && cashAmount > 0.005)
      ? { toPartnerId: paidByPartnerId, amount: cashAmount, returned: false, returnedAt: '' }
      : null;

    const entryId = generateId();
    const payoutEntry = normalizePayoutEntry({
      id: entryId,
      type: settlement.payoutType || 'monthly',
      label: settlement.payoutLabel || '',
      sourceMonth: (settlement.sourceMonth || '').toString().trim(),
      payoutDate: settlement.payoutDate || '',
      cashAmount,
      paidByPartnerId,
      deductedAdvances,
      advanceRefunds,
      salaryRefund,
      createdAt: new Date().toISOString()
    });

    // Set snapshots only on first payout
    const isFirstPayout = existingRecord.payouts.length === 0;
    const updatedRecord = normalizePayoutEmployeeRecord({
      ...existingRecord,
      payouts: [...existingRecord.payouts, payoutEntry],
      sourceMonth: settlement.sourceMonth || existingRecord.sourceMonth || '',
      baseAmountSnapshot: isFirstPayout
        ? Math.max(0, parseFloat(settlement.baseAmountSnapshot) || 0)
        : existingRecord.baseAmountSnapshot,
      carryoverAmountSnapshot: isFirstPayout
        ? Math.max(0, parseFloat(settlement.carryoverAmountSnapshot) || 0)
        : existingRecord.carryoverAmountSnapshot,
      plannedAmountSnapshot: isFirstPayout
        ? Math.max(0, parseFloat(settlement.plannedAmountSnapshot) || 0)
        : existingRecord.plannedAmountSnapshot
    });

    ms.payouts = normalizePayoutSettings({
      ...(ms.payouts || {}),
      employees: { ...(ms.payouts?.employees || {}), [normalizedPersonId]: updatedRecord }
    });
    Store.save(month);
    return true;
  },

  restoreAdvanceToCosts: (personId, payoutEntryId, advanceId, month = Store.getSelectedMonth()) => {
    const normalizedPersonId = (personId || '').toString().trim();
    if (!canMutateMonth(month) || !normalizedPersonId) return false;

    const ms = ensureMonthSettings(month);
    const existingRecord = normalizePayoutEmployeeRecord(ms.payouts?.employees?.[normalizedPersonId] || {});
    const monthRecord = appState.months[month] || (appState.months[month] = normalizeMonth(null));

    const payoutIndex = existingRecord.payouts.findIndex(p => p.id === payoutEntryId);
    if (payoutIndex === -1) return false;

    const entry = existingRecord.payouts[payoutIndex];
    const advIndex = entry.deductedAdvances.findIndex(a => a.id === advanceId && !a.restoredToCosts);
    if (advIndex === -1) return false;

    const adv = entry.deductedAdvances[advIndex];
    // Re-add as a COST expense
    monthRecord.expenses.push(normalizeExpenseRecord({
      id: generateId(),
      type: 'COST',
      date: adv.date || `${month}-01`,
      name: `[z zaliczki] ${adv.name}`.trim(),
      amount: adv.amount,
      paidById: adv.paidById
    }));

    const updatedAdvances = [...entry.deductedAdvances];
    updatedAdvances[advIndex] = { ...adv, restoredToCosts: true, restoredAt: new Date().toISOString() };

    const updatedPayouts = [...existingRecord.payouts];
    updatedPayouts[payoutIndex] = { ...entry, deductedAdvances: updatedAdvances };

    const updatedRecord = normalizePayoutEmployeeRecord({ ...existingRecord, payouts: updatedPayouts });
    ms.payouts = normalizePayoutSettings({
      ...(ms.payouts || {}),
      employees: { ...(ms.payouts?.employees || {}), [normalizedPersonId]: updatedRecord }
    });
    Store.save(month);
    return true;
  },

  markAdvanceRefundReturned: (personId, payoutEntryId, advanceId, month = Store.getSelectedMonth()) => {
    const normalizedPersonId = (personId || '').toString().trim();
    if (!canMutateMonth(month) || !normalizedPersonId) return false;

    const ms = ensureMonthSettings(month);
    const existingRecord = normalizePayoutEmployeeRecord(ms.payouts?.employees?.[normalizedPersonId] || {});

    const payoutIndex = existingRecord.payouts.findIndex(p => p.id === payoutEntryId);
    if (payoutIndex === -1) return false;

    const entry = existingRecord.payouts[payoutIndex];
    const refundIndex = entry.advanceRefunds.findIndex(r => r.advanceId === advanceId && !r.returned);
    if (refundIndex === -1) return false;

    const updatedRefunds = [...entry.advanceRefunds];
    updatedRefunds[refundIndex] = { ...updatedRefunds[refundIndex], returned: true, returnedAt: new Date().toISOString() };

    const updatedPayouts = [...existingRecord.payouts];
    updatedPayouts[payoutIndex] = { ...entry, advanceRefunds: updatedRefunds };

    const updatedRecord = normalizePayoutEmployeeRecord({ ...existingRecord, payouts: updatedPayouts });
    ms.payouts = normalizePayoutSettings({
      ...(ms.payouts || {}),
      employees: { ...(ms.payouts?.employees || {}), [normalizedPersonId]: updatedRecord }
    });
    Store.save(month);
    return true;
  },

  markSalaryRefundReturned: (personId, payoutEntryId, month = Store.getSelectedMonth()) => {
    const normalizedPersonId = (personId || '').toString().trim();
    if (!canMutateMonth(month) || !normalizedPersonId) return false;

    const ms = ensureMonthSettings(month);
    const existingRecord = normalizePayoutEmployeeRecord(ms.payouts?.employees?.[normalizedPersonId] || {});

    const payoutIndex = existingRecord.payouts.findIndex(p => p.id === payoutEntryId);
    if (payoutIndex === -1) return false;
    const entry = existingRecord.payouts[payoutIndex];
    if (!entry.salaryRefund || entry.salaryRefund.returned) return false;

    const updatedPayouts = [...existingRecord.payouts];
    updatedPayouts[payoutIndex] = {
      ...entry,
      salaryRefund: { ...entry.salaryRefund, returned: true, returnedAt: new Date().toISOString() }
    };

    const updatedRecord = normalizePayoutEmployeeRecord({ ...existingRecord, payouts: updatedPayouts });
    ms.payouts = normalizePayoutSettings({
      ...(ms.payouts || {}),
      employees: { ...(ms.payouts?.employees || {}), [normalizedPersonId]: updatedRecord }
    });
    Store.save(month);
    return true;
  },

  deletePayoutEntry: (personId, entryId, payoutMonth = Store.getSelectedMonth()) => {
    const normalizedPersonId = (personId || '').toString().trim();
    if (!canMutateMonth(payoutMonth) || !normalizedPersonId || !entryId) return false;

    const ms = ensureMonthSettings(payoutMonth);
    const existingRecord = normalizePayoutEmployeeRecord(ms.payouts?.employees?.[normalizedPersonId] || {});
    const monthRecord = appState.months[payoutMonth] || (appState.months[payoutMonth] = normalizeMonth(null));

    const entryIndex = existingRecord.payouts.findIndex(p => p.id === entryId);
    if (entryIndex === -1) return false;

    const entry = existingRecord.payouts[entryIndex];

    // Restore non-restored advances back as ADVANCE expenses
    entry.deductedAdvances
      .filter(adv => !adv.restoredToCosts)
      .forEach(adv => {
        monthRecord.expenses.push(normalizeExpenseRecord({
          id: generateId(),
          type: 'ADVANCE',
          date: adv.date || `${payoutMonth}-01`,
          name: adv.name || 'Zaliczka',
          amount: adv.amount,
          paidById: adv.paidById,
          advanceForId: normalizedPersonId
        }));
      });

    const updatedPayouts = existingRecord.payouts.filter(p => p.id !== entryId);
    const updatedRecord = normalizePayoutEmployeeRecord({
      ...existingRecord,
      payouts: updatedPayouts,
      ...(updatedPayouts.length === 0 ? {
        plannedAmountSnapshot: 0,
        baseAmountSnapshot: 0,
        carryoverAmountSnapshot: 0
      } : {})
    });

    ms.payouts = normalizePayoutSettings({
      ...(ms.payouts || {}),
      employees: { ...(ms.payouts?.employees || {}), [normalizedPersonId]: updatedRecord }
    });
    Store.save(payoutMonth);
    return true;
  },

  updateSeparateCompanyPayoutsConfig: (companyId, config = {}, month = Store.getSelectedMonth()) => {
    const normalizedId = (companyId || '').toString().trim();
    if (!canMutateMonth(month) || !normalizedId) return false;

    const ms = ensureMonthSettings(month);
    const existingSettings = ms.payouts || {};
    ms.payouts = normalizePayoutSettings({
      ...existingSettings,
      separateCompanies: {
        ...(existingSettings.separateCompanies || {}),
        [normalizedId]: {
          ...(existingSettings.separateCompanies?.[normalizedId] || {}),
          ...config
        }
      }
    });
    Store.save(month);
    return true;
  },

  updateInvoiceMonthConfig: (invoiceConfig = {}, month = Store.getSelectedMonth()) => {
    if (!canMutateMonth(month)) return false;
    const ms = ensureMonthSettings(month);
    if (normalizeInvoicesConfig(ms.invoices || {}).issued === true) return false;
    ms.invoices = normalizeInvoicesConfig({ ...ms.invoices, ...invoiceConfig });
    Store.save(month);
    return true;
  },
  updateInvoiceClientConfig: (clientId, clientConfig = {}, month = Store.getSelectedMonth()) => {
    if (!canMutateMonth(month) || !clientId) return false;
    const ms = ensureMonthSettings(month);
    if (normalizeInvoicesConfig(ms.invoices || {}).issued === true) return false;
    const existingClientConfig = ms.invoices?.clients?.[clientId] || {};

    ms.invoices = normalizeInvoicesConfig({
      ...ms.invoices,
      clients: {
        ...(ms.invoices?.clients || {}),
        [clientId]: {
          ...existingClientConfig,
          ...clientConfig
        }
      }
    });

    Store.save(month);
    return true;
  },
  setInvoicesIssued: (issuedSnapshot, month = Store.getSelectedMonth()) => {
    if (!canMutateMonth(month)) return false;
    const ms = ensureMonthSettings(month);
    ms.invoices = normalizeInvoicesConfig({
      ...ms.invoices,
      issued: true,
      issuedAt: getIsoTimestamp(),
      issuedSnapshot
    });
    Store.save(month);
    return true;
  },
  restoreInvoiceEditing: (month = Store.getSelectedMonth()) => {
    if (!canMutateMonth(month)) return false;
    const ms = ensureMonthSettings(month);
    ms.invoices = normalizeInvoicesConfig({
      ...ms.invoices,
      issued: false,
      issuedAt: '',
      issuedSnapshot: null
    });
    Store.save(month);
    return true;
  },
  setMonthArchived: (isArchived, month = Store.getSelectedMonth()) => {
    const ms = ensureMonthSettings(month);
    ms.isArchived = isArchived;
    if (isArchived && !ms.commonSnapshot) {
      ms.commonSnapshot = cloneData(appState.common);
    }
    Store.save(month);
  },
  deleteMonthCommonSnapshot: (month = Store.getSelectedMonth()) => {
    const ms = ensureMonthSettings(month);
    if (ms.isArchived === true) return false;
    delete ms.commonSnapshot;
    Store.save(month);
    return true;
  },
  normalizeMonth,
  updateSettings: () => false,
  undo: () => {
    const entry = localUndoStack[localUndoStack.length - 1];
    if (!entry) return { success: false, message: 'Brak zmian do cofnięcia.' };

    const conflicts = entry.changes.filter(change => !deepEqual(getScopeSnapshot(change.scope, change.month), change.afterSnapshot));
    if (conflicts.length > 0) {
      return { success: false, message: 'Nie można cofnąć tej zmiany, ponieważ dane zostały zmienione lokalnie lub zdalnie.' };
    }

    const context = createTrackedChangeContext(entry.changes.map(change => ({ scope: change.scope, month: change.month })), `Cofnięto: ${entry.label}`, {
      selectedMonth: entry.month || Store.getSelectedMonth(),
      skipLocalHistory: true,
      skipSharedHistory: true,
      forceImmediateRemoteSync: true
    });

    context.skipLocalHistory = true;
    context.skipSharedHistory = true;

    removeSharedHistoryEntryById(entry.id);

    entry.changes.forEach(change => {
      setScopeSnapshot(change.scope, change.month, change.beforeSnapshot);
    });

    activeMutationTransaction = context;
    const primaryMonth = entry.month || Store.getSelectedMonth();
    finalizeTrackedChangeContext(primaryMonth);

    localUndoStack.pop();
    localRedoStack.push(entry);
    dispatchStateEvents();
    return { success: true };
  },
  redo: () => {
    const entry = localRedoStack[localRedoStack.length - 1];
    if (!entry) return { success: false, message: 'Brak zmian do ponowienia.' };

    const conflicts = entry.changes.filter(change => !deepEqual(getScopeSnapshot(change.scope, change.month), change.beforeSnapshot));
    if (conflicts.length > 0) {
      return { success: false, message: 'Nie można ponowić tej zmiany, ponieważ dane zostały zmienione lokalnie lub zdalnie.' };
    }

    const context = createTrackedChangeContext(entry.changes.map(change => ({ scope: change.scope, month: change.month })), `Ponowiono: ${entry.label}`, {
      selectedMonth: entry.month || Store.getSelectedMonth(),
      skipLocalHistory: true,
      skipSharedHistory: true,
      forceImmediateRemoteSync: true
    });

    context.skipLocalHistory = true;
    context.skipSharedHistory = true;

    restoreSharedHistoryEntryToTop(entry);

    entry.changes.forEach(change => {
      setScopeSnapshot(change.scope, change.month, change.afterSnapshot);
    });

    activeMutationTransaction = context;
    const primaryMonth = entry.month || Store.getSelectedMonth();
    finalizeTrackedChangeContext(primaryMonth);

    localRedoStack.pop();
    localUndoStack.push(entry);
    dispatchStateEvents();
    return { success: true };
  },
  restoreSharedHistoryEntry: (entryId) => {
    return restoreSharedHistoryEntrySnapshot(entryId, 'beforeSnapshot');
  },
  restoreSharedHistoryEntryAfter: (entryId) => {
    return restoreSharedHistoryEntrySnapshot(entryId, 'afterSnapshot');
  },
  restoreSharedHistoryEntryMergeBefore: (entryId) => {
    return applySharedHistoryCollectionAction(entryId, 'merge-before');
  },
  restoreSharedHistoryEntryRemoveAdded: (entryId) => {
    return applySharedHistoryCollectionAction(entryId, 'remove-added');
  },
  restoreBackupEntry: (type, entryId) => {
    const sourceEntries = type === 'monthly' ? monthlyBackupEntries : dailyBackupEntries;
    const entry = sourceEntries.find(item => item.id === entryId);
    if (!entry) return { success: false, message: 'Nie znaleziono snapshotu.' };
    if (!backupEntryHasLoadedDetails(entry)) {
      return { success: false, message: 'Najpierw wczytaj pełne dane tego snapshotu.' };
    }

    const primaryMonth = entry.month || Store.getSelectedMonth();
    const changes = type === 'monthly' && entry.stateSnapshot
      ? buildMonthlySnapshotRestoreChanges(entry.stateSnapshot)
      : [
      { scope: 'common.persons', month: primaryMonth },
      { scope: 'common.clients', month: primaryMonth },
      { scope: 'common.worksCatalog', month: primaryMonth },
      { scope: 'common.config', month: primaryMonth },
      { scope: 'month.monthlySheets', month: primaryMonth },
      { scope: 'month.worksSheets', month: primaryMonth },
      { scope: 'month.expenses', month: primaryMonth },
      { scope: 'month.monthSettings.persons', month: primaryMonth },
      { scope: 'month.monthSettings.clients', month: primaryMonth },
      { scope: 'month.monthSettings.settlementConfig', month: primaryMonth },
      { scope: 'month.monthSettings.personContractCharges', month: primaryMonth },
      { scope: 'month.monthSettings.invoices', month: primaryMonth },
      { scope: 'month.monthSettings.archive', month: primaryMonth }
    ];

    const context = createTrackedChangeContext(changes, `Przywrócono snapshot ${type === 'monthly' ? 'miesięczny' : 'dzienny'}: ${entry.month}`, {
      selectedMonth: primaryMonth,
      skipLocalHistory: true,
      forceImmediateRemoteSync: true
    });

    context.skipLocalHistory = true;

    if (type === 'monthly' && entry.stateSnapshot) {
      appState = normalizeState(entry.stateSnapshot);
      Object.values(appState.months || {}).forEach(monthRecord => {
        ensureArchivedMonthHasSnapshot(monthRecord, appState.common);
      });
    } else {
      appState.common = normalizeCommonData(entry.commonSnapshot || {}, createDefaultState().common);
      appState.months[primaryMonth] = normalizeMonth(entry.monthSnapshot || {});
      ensureArchivedMonthHasSnapshot(appState.months[primaryMonth], appState.common);
    }

    activeMutationTransaction = context;
    finalizeTrackedChangeContext(primaryMonth);
    localRedoStack = [];
    dispatchStateEvents();
    return { success: true };
  },

  // Invoice Extra Invoices
  addInvoiceExtraInvoice: (invoice, month = Store.getSelectedMonth()) => {
    if (!canMutateMonth(month)) return false;
    const ms = ensureMonthSettings(month);
    if (normalizeInvoicesConfig(ms.invoices || {}).issued === true) return false;
    if (!ms.invoices) ms.invoices = { issueDate: '', emailIntro: '', clients: {} };
    if (!Array.isArray(ms.invoices.extraInvoices)) ms.invoices.extraInvoices = [];
    ms.invoices.extraInvoices.push({ ...invoice, id: generateId() });
    ms.invoices = normalizeInvoicesConfig(ms.invoices);
    Store.save(month);
    return true;
  },
  updateInvoiceExtraInvoice: (id, updates, month = Store.getSelectedMonth()) => {
    if (!canMutateMonth(month)) return false;
    const ms = ensureMonthSettings(month);
    if (normalizeInvoicesConfig(ms.invoices || {}).issued === true) return false;
    if (!ms.invoices?.extraInvoices) return;
    const idx = ms.invoices.extraInvoices.findIndex(inv => inv.id === id);
    if (idx !== -1) {
      ms.invoices.extraInvoices[idx] = { ...ms.invoices.extraInvoices[idx], ...updates };
      ms.invoices = normalizeInvoicesConfig(ms.invoices);
      Store.save(month);
      return true;
    }
    return false;
  },
  deleteInvoiceExtraInvoice: (id, month = Store.getSelectedMonth()) => {
    if (!canMutateMonth(month)) return false;
    const ms = ensureMonthSettings(month);
    if (normalizeInvoicesConfig(ms.invoices || {}).issued === true) return false;
    if (!ms.invoices?.extraInvoices) return;
    ms.invoices.extraInvoices = ms.invoices.extraInvoices.filter(inv => inv.id !== id);
    ms.invoices = normalizeInvoicesConfig(ms.invoices);
    Store.save(month);
    return true;
  },

  importState: (rawState, returnOnly = false) => {
    const ns = normalizeState(rawState);
    if (returnOnly) return ns;
    appState = ns;
    syncMetadata = rebuildSyncMetadataFromState(appState);
    localUndoStack = [];
    localRedoStack = [];
    persistAllMonthsToLocalStorage();
    applyTrackedChangePersistence(Object.keys(appState.months || {}), true, {
      ...buildIndexedRemoteUpsertUpdatesFromCollection(sharedHistoryEntries, SHARED_HISTORY_INDEX_ROOT, SHARED_HISTORY_ENTRIES_ROOT, buildSharedHistoryIndexEntry, normalizeSharedHistoryEntries),
      ...buildIndexedRemoteUpsertUpdatesFromCollection(dailyBackupEntries, SHARED_DAILY_BACKUPS_INDEX_ROOT, SHARED_DAILY_BACKUPS_ENTRIES_ROOT, buildBackupIndexEntry, normalizeBackupEntries),
      ...buildIndexedRemoteUpsertUpdatesFromCollection(monthlyBackupEntries, SHARED_MONTHLY_BACKUPS_INDEX_ROOT, SHARED_MONTHLY_BACKUPS_ENTRIES_ROOT, buildBackupIndexEntry, normalizeBackupEntries)
    });
    return ns;
  },

  loadYearData: (year, upToMonth = 12) => {
    if (window.isOfflineMode) return Promise.resolve();
    if (typeof firebase === 'undefined' || !firebase.auth || !firebase.auth().currentUser) return Promise.resolve();

    const db = firebase.database();
    const promises = [];

    console.log(`firebase: Rozpoczynanie dociągania danych historycznych dla roku ${year}...`);

    for (let m = 1; m < upToMonth; m++) {
      const monthKey = `${year}-${String(m).padStart(2, '0')}`;
      const p = db.ref(`shared_meta/months/${monthKey}`).once('value').then(metaSnapshot => {
        const trackedMetaSnapshot = trackFirebaseSnapshotValue(metaSnapshot);
        const remoteMeta = { months: { [monthKey]: trackedMetaSnapshot.val() || {} } };
        if (appState.months[monthKey] && !Store.shouldFetchRemoteMonth(monthKey, remoteMeta)) {
          return null;
        }

        return db.ref(`shared_data/months/${monthKey}`).once('value').then(snapshot => {
          const trackedSnapshot = trackFirebaseSnapshotValue(snapshot);
          const val = trackedSnapshot.val();
          if (val) {
            appState.months[monthKey] = normalizeMonth(val);
            syncMetadata.months[monthKey] = normalizeSyncMetadata(remoteMeta).months[monthKey] || syncMetadata.months[monthKey] || {};
          }
        });
      }).catch(error => {
        if (!isFirebasePermissionDeniedError(error)) {
          throw error;
        }

        return db.ref(`shared_data/months/${monthKey}`).once('value').then(snapshot => {
          const trackedSnapshot = trackFirebaseSnapshotValue(snapshot);
          const val = trackedSnapshot.val();
          if (val) {
            appState.months[monthKey] = normalizeMonth(val);
          }
        });
      });
      promises.push(p);
    }

    return Promise.all(promises).then(() => {
      syncMetadata = mergeSyncMetadataWithState(appState, syncMetadata).meta;
      persistAllMonthsToLocalStorage();
      persistSyncMetadataToLocalStorage();
      console.log(`firebase: Zakończono dociąganie danych historycznych (${promises.length} miesięcy).`);
      window.dispatchEvent(new Event('appStateChanged'));
    });
  }
};

const wrapTrackedStoreMethod = (methodName, getTrackedChanges) => {
  const originalMethod = Store[methodName];
  if (typeof originalMethod !== 'function') return;

  Store[methodName] = (...args) => {
    const trackedChanges = (typeof getTrackedChanges === 'function' ? getTrackedChanges(...args) : null) || null;
    if (!trackedChanges || !Array.isArray(trackedChanges.changes) || trackedChanges.changes.length === 0) {
      return originalMethod(...args);
    }

    activeMutationTransaction = createTrackedChangeContext(trackedChanges.changes, trackedChanges.label, {
      selectedMonth: trackedChanges.selectedMonth,
      skipLocalHistory: trackedChanges.skipLocalHistory === true,
      skipSharedHistory: trackedChanges.skipSharedHistory === true
    });

    const result = originalMethod(...args);
    if (activeMutationTransaction) {
      activeMutationTransaction = null;
    }
    return result;
  };
};

const buildCommonTrackedChange = (baseScope, label, month = Store.getSelectedMonth()) => ({
  scope: resolveCommonMutationScope(baseScope, month),
  month,
  scopeLabel: getScopeDisplayLabel(resolveCommonMutationScope(baseScope, month)),
  label
});

wrapTrackedStoreMethod('addPerson', () => ({
  label: 'Dodano osobę',
  changes: [buildCommonTrackedChange('common.persons', 'Dodano osobę')]
}));
wrapTrackedStoreMethod('updatePerson', () => ({
  label: 'Zaktualizowano osobę',
  changes: [buildCommonTrackedChange('common.persons', 'Zaktualizowano osobę')]
}));
wrapTrackedStoreMethod('deletePerson', () => ({
  label: 'Usunięto osobę',
  changes: [
    buildCommonTrackedChange('common.persons', 'Usunięto osobę'),
    ...buildOrphanedPersonCleanupTrackedChanges(Store.getSelectedMonth())
  ]
}));
wrapTrackedStoreMethod('cleanupOrphanedPersonEntries', (month = Store.getSelectedMonth()) => ({
  label: 'Usunięto osierocone wpisy osób z miesiąca',
  selectedMonth: month,
  changes: buildOrphanedPersonCleanupTrackedChanges(month)
}));
wrapTrackedStoreMethod('reorderPersons', () => ({
  label: 'Zmieniono kolejność osób',
  changes: [buildCommonTrackedChange('common.persons', 'Zmieniono kolejność osób')]
}));
wrapTrackedStoreMethod('togglePersonStatus', (id, month = Store.getSelectedMonth()) => ({
  label: 'Zmieniono status osoby w miesiącu',
  selectedMonth: month,
  changes: [{ scope: 'month.monthSettings.persons', month, label: 'Zmieniono status osoby w miesiącu' }]
}));

wrapTrackedStoreMethod('addMonthlySheet', (sheet = {}) => ({
  label: 'Dodano arkusz godzin',
  selectedMonth: sheet?.month || Store.getSelectedMonth(),
  changes: [{ scope: 'month.monthlySheets', month: sheet?.month || Store.getSelectedMonth(), label: 'Dodano arkusz godzin' }]
}));
wrapTrackedStoreMethod('updateMonthlySheet', () => ({
  label: 'Zaktualizowano arkusz godzin',
  changes: [{ scope: 'month.monthlySheets', month: Store.getSelectedMonth(), label: 'Zaktualizowano arkusz godzin' }]
}));
wrapTrackedStoreMethod('deleteMonthlySheet', () => ({
  label: 'Usunięto arkusz godzin',
  changes: [{ scope: 'month.monthlySheets', month: Store.getSelectedMonth(), label: 'Usunięto arkusz godzin' }]
}));
wrapTrackedStoreMethod('reorderMonthlySheets', (newOrderedIds, month = Store.getSelectedMonth()) => ({
  label: 'Zmieniono kolejność arkuszy godzin',
  selectedMonth: month,
  changes: [{ scope: 'month.monthlySheets', month, label: 'Zmieniono kolejność arkuszy godzin' }]
}));

wrapTrackedStoreMethod('addWorksSheet', (sheet = {}) => ({
  label: 'Dodano arkusz prac',
  selectedMonth: sheet?.month || Store.getSelectedMonth(),
  changes: [{ scope: 'month.worksSheets', month: sheet?.month || Store.getSelectedMonth(), label: 'Dodano arkusz prac' }]
}));
wrapTrackedStoreMethod('updateWorksSheet', () => ({
  label: 'Zaktualizowano arkusz prac',
  changes: [{ scope: 'month.worksSheets', month: Store.getSelectedMonth(), label: 'Zaktualizowano arkusz prac' }]
}));
wrapTrackedStoreMethod('deleteWorksSheet', () => ({
  label: 'Usunięto arkusz prac',
  changes: [{ scope: 'month.worksSheets', month: Store.getSelectedMonth(), label: 'Usunięto arkusz prac' }]
}));
wrapTrackedStoreMethod('reorderWorksSheets', (newOrderedIds, month = Store.getSelectedMonth()) => ({
  label: 'Zmieniono kolejność arkuszy prac',
  selectedMonth: month,
  changes: [{ scope: 'month.worksSheets', month, label: 'Zmieniono kolejność arkuszy prac' }]
}));
wrapTrackedStoreMethod('reorderWorksSheetEntriesForDate', (sheetId, date, newOrderedIds, month = Store.getSelectedMonth()) => ({
  label: 'Zmieniono kolejność pozycji prac dla tej samej daty',
  selectedMonth: month,
  changes: [{ scope: 'month.worksSheets', month, label: 'Zmieniono kolejność pozycji prac dla tej samej daty' }]
}));

wrapTrackedStoreMethod('addWorkToCatalog', () => ({
  label: 'Dodano pozycję katalogu prac',
  changes: [buildCommonTrackedChange('common.worksCatalog', 'Dodano pozycję katalogu prac')]
}));
wrapTrackedStoreMethod('updateWorkInCatalog', () => ({
  label: 'Zaktualizowano pozycję katalogu prac',
  changes: [buildCommonTrackedChange('common.worksCatalog', 'Zaktualizowano pozycję katalogu prac')]
}));
wrapTrackedStoreMethod('deleteWorkFromCatalog', () => ({
  label: 'Usunięto pozycję katalogu prac',
  changes: [buildCommonTrackedChange('common.worksCatalog', 'Usunięto pozycję katalogu prac')]
}));
wrapTrackedStoreMethod('restoreWorkInCatalogToDefault', () => ({
  label: 'Przywrócono domyślną pozycję katalogu prac',
  changes: [buildCommonTrackedChange('common.worksCatalog', 'Przywrócono domyślną pozycję katalogu prac')]
}));
wrapTrackedStoreMethod('restoreDeletedDefaultToCatalog', () => ({
  label: 'Przywrócono usuniętą pozycję katalogu prac',
  changes: [buildCommonTrackedChange('common.worksCatalog', 'Przywrócono usuniętą pozycję katalogu prac')]
}));

wrapTrackedStoreMethod('addExpense', () => ({
  label: 'Dodano koszt lub zaliczkę',
  changes: [{ scope: 'month.expenses', month: Store.getSelectedMonth(), label: 'Dodano koszt lub zaliczkę' }]
}));
wrapTrackedStoreMethod('updateExpense', () => ({
  label: 'Zaktualizowano koszt lub zaliczkę',
  changes: [{ scope: 'month.expenses', month: Store.getSelectedMonth(), label: 'Zaktualizowano koszt lub zaliczkę' }]
}));
wrapTrackedStoreMethod('deleteExpense', () => ({
  label: 'Usunięto koszt lub zaliczkę',
  changes: [{ scope: 'month.expenses', month: Store.getSelectedMonth(), label: 'Usunięto koszt lub zaliczkę' }]
}));
wrapTrackedStoreMethod('reorderExpensesForDate', (date, newOrderedIds, month = Store.getSelectedMonth()) => ({
  label: 'Zmieniono kolejność kosztów dla tej samej daty',
  selectedMonth: month,
  changes: [{ scope: 'month.expenses', month, label: 'Zmieniono kolejność kosztów dla tej samej daty' }]
}));

wrapTrackedStoreMethod('addClient', () => ({
  label: 'Dodano klienta',
  changes: [buildCommonTrackedChange('common.clients', 'Dodano klienta')]
}));
wrapTrackedStoreMethod('updateClient', () => ({
  label: 'Zaktualizowano klienta',
  changes: [buildCommonTrackedChange('common.clients', 'Zaktualizowano klienta')]
}));
wrapTrackedStoreMethod('deleteClient', () => ({
  label: 'Usunięto klienta',
  changes: [buildCommonTrackedChange('common.clients', 'Usunięto klienta')]
}));
wrapTrackedStoreMethod('toggleClientStatus', (id, month = Store.getSelectedMonth()) => ({
  label: 'Zmieniono status klienta w miesiącu',
  selectedMonth: month,
  changes: [{ scope: 'month.monthSettings.clients', month, label: 'Zmieniono status klienta w miesiącu' }]
}));
wrapTrackedStoreMethod('reorderClients', () => ({
  label: 'Zmieniono kolejność klientów',
  changes: [buildCommonTrackedChange('common.clients', 'Zmieniono kolejność klientów')]
}));

wrapTrackedStoreMethod('updateConfig', () => ({
  label: 'Zaktualizowano konfigurację wspólną',
  changes: [buildCommonTrackedChange('common.config', 'Zaktualizowano konfigurację wspólną')]
}));
wrapTrackedStoreMethod('updateSettlementMonthConfig', (monthConfig = {}, personContractCharges = {}, month = Store.getSelectedMonth()) => ({
  label: 'Zaktualizowano konfigurację rozliczenia miesiąca',
  selectedMonth: month,
  changes: [
    { scope: 'month.monthSettings.settlementConfig', month, label: 'Zaktualizowano konfigurację rozliczenia miesiąca' },
    { scope: 'month.monthSettings.personContractCharges', month, label: 'Zaktualizowano podatek i ZUS UZ miesiąca' }
  ]
}));
wrapTrackedStoreMethod('updateInvoiceMonthConfig', (invoiceConfig = {}, month = Store.getSelectedMonth()) => ({
  label: 'Zaktualizowano konfigurację faktur miesiąca',
  selectedMonth: month,
  changes: [{ scope: 'month.monthSettings.invoices', month, label: 'Zaktualizowano konfigurację faktur miesiąca' }]
}));
wrapTrackedStoreMethod('updateInvoiceClientConfig', (clientId, clientConfig = {}, month = Store.getSelectedMonth()) => ({
  label: 'Zaktualizowano ustawienia faktur klienta',
  selectedMonth: month,
  changes: [{ scope: 'month.monthSettings.invoices', month, label: 'Zaktualizowano ustawienia faktur klienta' }]
}));
wrapTrackedStoreMethod('setInvoicesIssued', (issuedSnapshot, month = Store.getSelectedMonth()) => ({
  label: 'Oznaczono faktury jako wystawione',
  selectedMonth: month,
  changes: [{ scope: 'month.monthSettings.invoices', month, label: 'Oznaczono faktury jako wystawione' }]
}));
wrapTrackedStoreMethod('restoreInvoiceEditing', (month = Store.getSelectedMonth()) => ({
  label: 'Przywrócono edycję faktur',
  selectedMonth: month,
  changes: [{ scope: 'month.monthSettings.invoices', month, label: 'Przywrócono edycję faktur' }]
}));
wrapTrackedStoreMethod('setMonthArchived', (isArchived, month = Store.getSelectedMonth()) => ({
  label: isArchived ? 'Zarchiwizowano miesiąc' : 'Przywrócono edycję miesiąca',
  selectedMonth: month,
  changes: [{ scope: 'month.monthSettings.archive', month, label: isArchived ? 'Zarchiwizowano miesiąc' : 'Przywrócono edycję miesiąca' }]
}));
wrapTrackedStoreMethod('deleteMonthCommonSnapshot', (month = Store.getSelectedMonth()) => ({
  label: 'Usunięto snapshot miesiąca',
  selectedMonth: month,
  changes: [{ scope: 'month.monthSettings.archive', month, label: 'Usunięto snapshot miesiąca' }]
}));
wrapTrackedStoreMethod('addInvoiceExtraInvoice', (invoice, month = Store.getSelectedMonth()) => ({
  label: 'Dodano dodatkową fakturę',
  selectedMonth: month,
  changes: [{ scope: 'month.monthSettings.invoices', month, label: 'Dodano dodatkową fakturę' }]
}));
wrapTrackedStoreMethod('updateInvoiceExtraInvoice', (id, updates, month = Store.getSelectedMonth()) => ({
  label: 'Zaktualizowano dodatkową fakturę',
  selectedMonth: month,
  changes: [{ scope: 'month.monthSettings.invoices', month, label: 'Zaktualizowano dodatkową fakturę' }]
}));
wrapTrackedStoreMethod('deleteInvoiceExtraInvoice', (id, month = Store.getSelectedMonth()) => ({
  label: 'Usunięto dodatkową fakturę',
  selectedMonth: month,
  changes: [{ scope: 'month.monthSettings.invoices', month, label: 'Usunięto dodatkową fakturę' }]
}));
