// Helper for unique IDs
const generateId = () => crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).substring(2, 9);
const getCurrentMonthKey = () => {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
};

const WORK_TRACKER_STATE_KEY = 'workTrackerState';
const WORK_TRACKER_LOCAL_SETTINGS_KEY = 'workTrackerLocalSettings';

const safeParseStorage = (key) => {
  try {
    const rawValue = localStorage.getItem(key);
    return rawValue ? JSON.parse(rawValue) : null;
  } catch {
    return null;
  }
};

const ensureMonthSettings = (month) => {
  const monthKey = month || appState.selectedMonth || getCurrentMonthKey();
  if (!appState.monthSettings) appState.monthSettings = {};
  if (!appState.monthSettings[monthKey]) {
    appState.monthSettings[monthKey] = {
      persons: {},
      clients: {},
      settlementConfig: {},
      personContractCharges: {},
      invoices: {
        issueDate: '',
        emailIntro: '',
        clients: {}
      }
    };
  }
  if (!appState.monthSettings[monthKey].persons) appState.monthSettings[monthKey].persons = {};
  if (!appState.monthSettings[monthKey].clients) appState.monthSettings[monthKey].clients = {};
  if (!appState.monthSettings[monthKey].settlementConfig) appState.monthSettings[monthKey].settlementConfig = {};
  if (!appState.monthSettings[monthKey].personContractCharges) appState.monthSettings[monthKey].personContractCharges = {};
  if (!appState.monthSettings[monthKey].invoices) {
    appState.monthSettings[monthKey].invoices = {
      issueDate: '',
      emailIntro: '',
      clients: {}
    };
  }
  if (typeof appState.monthSettings[monthKey].invoices.issueDate !== 'string') appState.monthSettings[monthKey].invoices.issueDate = '';
  if (typeof appState.monthSettings[monthKey].invoices.emailIntro !== 'string') appState.monthSettings[monthKey].invoices.emailIntro = '';
  if (!appState.monthSettings[monthKey].invoices.clients || typeof appState.monthSettings[monthKey].invoices.clients !== 'object' || Array.isArray(appState.monthSettings[monthKey].invoices.clients)) {
    appState.monthSettings[monthKey].invoices.clients = {};
  }
  return appState.monthSettings[monthKey];
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

const DEFAULT_GOOGLE_DRIVE_FILE_ID = '1QRpJ99tTAakrQyn8yAFdZC6cuzsPdG-U';
const DEFAULT_CONTRACT_TAX_AMOUNT = 104;
const DEFAULT_CONTRACT_ZUS_AMOUNT = 490.27;
const DEFAULT_PROFIT_SHARE_PERCENT = 100;

const createDefaultSharedSettings = () => ({
  googleDriveFileId: DEFAULT_GOOGLE_DRIVE_FILE_ID,
  googleDriveAutoSync: false
});

const createDefaultLocalSettings = () => ({
  theme: 'system',
  scaleLarge: 100,
  scaleVertical: 100
});

const clampPercent = (value, fallback = DEFAULT_PROFIT_SHARE_PERCENT) => {
  const parsed = parseFloat(value);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.max(0, Math.min(100, parsed));
};

const clampScale = (value, fallback = 100) => {
  const parsed = parseInt(value, 10);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.max(20, Math.min(300, parsed));
};

const normalizeSharedSettings = (settings = {}) => ({
  googleDriveFileId: typeof settings.googleDriveFileId === 'string' && settings.googleDriveFileId.trim() !== ''
    ? settings.googleDriveFileId.trim()
    : DEFAULT_GOOGLE_DRIVE_FILE_ID,
  googleDriveAutoSync: settings.googleDriveAutoSync === true
});

const normalizeLocalSettings = (settings = {}) => {
  const theme = typeof settings.theme === 'string' && ['system', 'light', 'dark'].includes(settings.theme)
    ? settings.theme
    : 'system';

  return {
    theme,
    scaleLarge: clampScale(settings.scaleLarge, 100),
    scaleVertical: clampScale(settings.scaleVertical, 100)
  };
};

const normalizeClientRecord = (client = {}, existingClient = {}) => {
  const normalizeString = (value, fallback = '') => (value ?? fallback).toString().trim();
  const hourlyRate = parseFloat(client.hourlyRate ?? existingClient.hourlyRate);

  return {
    ...existingClient,
    ...client,
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
  const monthKey = month || state?.selectedMonth || getCurrentMonthKey();
  const monthSettings = state?.monthSettings?.[monthKey];
  const monthPersonStatuses = monthSettings?.persons || {};

  return (state?.persons || [])
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

const normalizeInvoicesConfig = (config = {}) => {
  const clients = {};
  Object.entries(config.clients || {}).forEach(([clientId, clientConfig]) => {
    clients[clientId] = normalizeInvoiceClientConfig(clientConfig);
  });

  return {
    issueDate: (config.issueDate || '').toString().trim(),
    emailIntro: (config.emailIntro || '').toString().trim(),
    clients,
    extraInvoices: Array.isArray(config.extraInvoices)
      ? config.extraInvoices.map(invoice => normalizeExtraInvoiceRecord(invoice)).filter(invoice => invoice.issuerId && invoice.amount > 0 && (invoice.clientId || invoice.clientName))
      : []
  };
};

const createDefaultState = () => ({
  selectedMonth: getCurrentMonthKey(),
  monthSettings: {},
  persons: [],
  dayRecords: [],
  monthlySheets: [],
  worksSheets: [],
  worksCatalog: DEFAULT_WORKS_CATALOG.map(w => ({ ...w, id: generateId() })),
  expenses: [],
  clients: [],
  config: {
    taxRate: 0.055,
    zusFixedAmount: 1600.27
  },
  settings: createDefaultSharedSettings()
});

const defaultState = createDefaultState();
const APP_STATE_KEYS = ['selectedMonth', 'monthSettings', 'persons', 'dayRecords', 'monthlySheets', 'worksSheets', 'worksCatalog', 'expenses', 'clients', 'config', 'settings'];

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

const normalizeState = (rawState) => {
  const baseState = createDefaultState();
  const candidate = rawState && typeof rawState === 'object' && !Array.isArray(rawState) ? rawState : {};

  const normalizedState = {
    ...baseState,
    ...candidate,
    selectedMonth: typeof candidate.selectedMonth === 'string' && /^\d{4}-\d{2}$/.test(candidate.selectedMonth)
      ? candidate.selectedMonth
      : getCurrentMonthKey(),
    monthSettings: candidate.monthSettings && typeof candidate.monthSettings === 'object' && !Array.isArray(candidate.monthSettings)
      ? candidate.monthSettings
      : {},
    persons: Array.isArray(candidate.persons)
      ? candidate.persons.map(person => normalizePersonRecord(person))
      : [],
    dayRecords: Array.isArray(candidate.dayRecords) ? candidate.dayRecords : [],
    monthlySheets: Array.isArray(candidate.monthlySheets) ? candidate.monthlySheets : [],
    worksSheets: Array.isArray(candidate.worksSheets) ? candidate.worksSheets : [],
    worksCatalog: Array.isArray(candidate.worksCatalog) && candidate.worksCatalog.length > 0 ? candidate.worksCatalog : baseState.worksCatalog,
    expenses: Array.isArray(candidate.expenses) ? candidate.expenses : [],
    clients: Array.isArray(candidate.clients)
      ? candidate.clients.map(client => normalizeClientRecord(client))
      : [],
    config: {
      ...baseState.config,
      ...(candidate.config && typeof candidate.config === 'object' && !Array.isArray(candidate.config) ? candidate.config : {})
    },
    settings: normalizeSharedSettings(candidate.settings && typeof candidate.settings === 'object' && !Array.isArray(candidate.settings) ? candidate.settings : {})
  };

  normalizedState.worksCatalog.forEach(w => {
    if (w.unit === 'm3') w.unit = 'm³';
    if (w.unit === 'm2') w.unit = 'm²';
    if (!w.coreId) {
      const def = DEFAULT_WORKS_CATALOG.find(d => d.name === w.name);
      if (def) w.coreId = def.coreId;
    }
  });

  normalizedState.worksSheets.forEach(s => {
    if (s.entries) {
      s.entries.forEach(e => {
        if (e.unit === 'm3') e.unit = 'm³';
        if (e.unit === 'm2') e.unit = 'm²';
      });
    }
  });

  Object.values(normalizedState.monthSettings).forEach(monthSettings => {
    if (!monthSettings || typeof monthSettings !== 'object' || Array.isArray(monthSettings)) return;
    monthSettings.persons = monthSettings.persons && typeof monthSettings.persons === 'object' && !Array.isArray(monthSettings.persons) ? monthSettings.persons : {};
    monthSettings.clients = monthSettings.clients && typeof monthSettings.clients === 'object' && !Array.isArray(monthSettings.clients) ? monthSettings.clients : {};
    monthSettings.settlementConfig = monthSettings.settlementConfig && typeof monthSettings.settlementConfig === 'object' && !Array.isArray(monthSettings.settlementConfig) ? monthSettings.settlementConfig : {};
    monthSettings.personContractCharges = monthSettings.personContractCharges && typeof monthSettings.personContractCharges === 'object' && !Array.isArray(monthSettings.personContractCharges) ? monthSettings.personContractCharges : {};
    monthSettings.invoices = normalizeInvoicesConfig(monthSettings.invoices || {});
  });

  normalizedState.monthlySheets.forEach(sheet => {
    if (!Array.isArray(sheet.activePersons)) {
      sheet.activePersons = getActivePersonIdsForMonthFromState(normalizedState, sheet.month);
    } else {
      sheet.activePersons = [...new Set(sheet.activePersons.filter(value => typeof value === 'string' && value.trim() !== ''))];
    }
  });

  return normalizedState;
};

const persistedState = safeParseStorage(WORK_TRACKER_STATE_KEY);
const persistedLocalSettings = safeParseStorage(WORK_TRACKER_LOCAL_SETTINGS_KEY);

let appState = normalizeState(persistedState);
let localSettings = normalizeLocalSettings(persistedLocalSettings || persistedState?.settings || createDefaultLocalSettings());

ensureMonthSettings(appState.selectedMonth);

// Store fixed data so next render pulls correct superscript
localStorage.setItem(WORK_TRACKER_STATE_KEY, JSON.stringify(appState));
localStorage.setItem(WORK_TRACKER_LOCAL_SETTINGS_KEY, JSON.stringify(localSettings));

const Store = {
  getState: () => appState,
  getSettings: () => ({ ...appState.settings, ...localSettings }),
  getAppearanceSettings: () => ({ ...localSettings }),
  getSelectedMonth: () => appState.selectedMonth,
  getExportData: () => JSON.parse(JSON.stringify(appState)),

  setSelectedMonth: (month) => {
    appState.selectedMonth = month || getCurrentMonthKey();
    ensureMonthSettings(appState.selectedMonth);
    Store.save();
  },

  getMonthSettings: (month = appState.selectedMonth) => {
    return ensureMonthSettings(month);
  },

  isPersonActiveInMonth: (personId, month = appState.selectedMonth) => {
    const person = appState.persons.find(p => p.id === personId);
    if (!person) return false;

    const monthSettings = ensureMonthSettings(month);
    if (monthSettings.persons[personId] !== undefined) {
      return monthSettings.persons[personId] !== false;
    }

    return person.isActive !== false;
  },

  isClientActiveInMonth: (clientId, month = appState.selectedMonth) => {
    const client = appState.clients.find(c => c.id === clientId);
    if (!client) return false;

    const monthSettings = ensureMonthSettings(month);
    if (monthSettings.clients[clientId] !== undefined) {
      return monthSettings.clients[clientId] !== false;
    }

    return client.isActive !== false;
  },
  
  save: () => {
    localStorage.setItem(WORK_TRACKER_STATE_KEY, JSON.stringify(appState));
    if (typeof firebase !== 'undefined' && firebase.auth && firebase.auth().currentUser && !window.isImportingFromFirebase) {
      firebase.database().ref('shared_data').set(appState);
    }
    window.dispatchEvent(new Event('appStateChanged'));
  },


  saveLocalSettings: () => {
    localStorage.setItem(WORK_TRACKER_LOCAL_SETTINGS_KEY, JSON.stringify(localSettings));
    window.dispatchEvent(new Event('appStateChanged'));
  },

  // Persons
  addPerson: (person) => {
    appState.persons.push(normalizePersonRecord({ ...person, id: generateId(), isActive: true }));
    Store.save();
  },
  
  updatePerson: (id, updates) => {
    const idx = appState.persons.findIndex(p => p.id === id);
    if (idx !== -1) {
      appState.persons[idx] = normalizePersonRecord(updates, appState.persons[idx]);
      Store.save();
    }
  },

  deletePerson: (id) => {
    appState.persons = appState.persons.filter(p => p.id !== id);
    Store.save();
  },

  reorderPersons: (newOrderedIds) => {
    const reordered = newOrderedIds.map(id => appState.persons.find(p => p.id === id)).filter(Boolean);
    // Add any missing ones if they exist just in case
    const missing = appState.persons.filter(p => !newOrderedIds.includes(p.id));
    appState.persons = [...reordered, ...missing];
    Store.save();
  },

  togglePersonStatus: (id, month = appState.selectedMonth) => {
    const person = appState.persons.find(p => p.id === id);
    if (person) {
      const monthSettings = ensureMonthSettings(month);
      const currentValue = Store.isPersonActiveInMonth(id, month);
      monthSettings.persons[id] = !currentValue;
      Store.save();
    }
  },

  setPersonStatusForMonth: (id, isActive, month = appState.selectedMonth) => {
    const person = appState.persons.find(p => p.id === id);
    if (person) {
      const monthSettings = ensureMonthSettings(month);
      monthSettings.persons[id] = isActive !== false;
      Store.save();
    }
  },

  // Day Records
  saveDayRecord: (record) => {
    const idx = appState.dayRecords.findIndex(r => r.date === record.date);
    if (idx !== -1) {
      appState.dayRecords[idx] = record;
    } else {
      if (!record.id) record.id = generateId();
      appState.dayRecords.push(record);
    }
    Store.save();
  },

  deleteDayRecord: (id) => {
    appState.dayRecords = appState.dayRecords.filter(r => r.id !== id);
    Store.save();
  },
  
  getDayRecord: (date) => {
    return appState.dayRecords.find(r => r.date === date) || null;
  },

  // Monthly Sheets
  addMonthlySheet: (sheet) => {
    const month = sheet?.month || appState.selectedMonth;
    const activePersons = Array.isArray(sheet?.activePersons)
      ? [...new Set(sheet.activePersons.filter(value => typeof value === 'string' && value.trim() !== ''))]
      : getActivePersonIdsForMonthFromState(appState, month);
    appState.monthlySheets.push({ ...sheet, id: generateId(), activePersons });
    Store.save();
  },

  updateMonthlySheet: (id, updates) => {
    const idx = appState.monthlySheets.findIndex(s => s.id === id);
    if (idx !== -1) {
      appState.monthlySheets[idx] = { ...appState.monthlySheets[idx], ...updates };
      Store.save();
    }
  },

  deleteMonthlySheet: (id) => {
    appState.monthlySheets = appState.monthlySheets.filter(s => s.id !== id);
    Store.save();
  },

  getMonthlySheet: (id) => {
    return appState.monthlySheets.find(s => s.id === id) || null;
  },

  // Works Sheets
  addWorksSheet: (sheet) => {
    appState.worksSheets.push({ ...sheet, id: generateId(), entries: [] });
    Store.save();
  },

  updateWorksSheet: (id, updates) => {
    const idx = appState.worksSheets.findIndex(s => s.id === id);
    if (idx !== -1) {
      appState.worksSheets[idx] = { ...appState.worksSheets[idx], ...updates };
      Store.save();
    }
  },

  deleteWorksSheet: (id) => {
    appState.worksSheets = appState.worksSheets.filter(s => s.id !== id);
    Store.save();
  },

  getWorksSheet: (id) => {
    return appState.worksSheets.find(s => s.id === id) || null;
  },

  // Works Catalog
  addWorkToCatalog: (work) => {
    appState.worksCatalog.push({ ...work, id: generateId() });
    Store.save();
  },

  updateWorkInCatalog: (id, updates) => {
    const idx = appState.worksCatalog.findIndex(w => w.id === id);
    if (idx !== -1) {
      appState.worksCatalog[idx] = { ...appState.worksCatalog[idx], ...updates };
      Store.save();
    }
  },

  restoreWorkInCatalogToDefault: (id) => {
    const idx = appState.worksCatalog.findIndex(w => w.id === id);
    if (idx !== -1) {
      const w = appState.worksCatalog[idx];
      if (w.coreId) {
        const def = DEFAULT_WORKS_CATALOG.find(d => d.coreId === w.coreId);
        if (def) {
          appState.worksCatalog[idx] = { ...w, name: def.name, unit: def.unit, defaultPrice: def.defaultPrice };
          Store.save();
        }
      }
    }
  },

  getDefaultCatalogItem: (coreId) => {
    return DEFAULT_WORKS_CATALOG.find(d => d.coreId === coreId);
  },

  getDefaultCatalogItems: () => {
    return DEFAULT_WORKS_CATALOG;
  },

  restoreDeletedDefaultToCatalog: (coreId) => {
    const def = DEFAULT_WORKS_CATALOG.find(d => d.coreId === coreId);
    if (def) {
      appState.worksCatalog.push({ ...def, id: generateId() });
      Store.save();
    }
  },

  deleteWorkFromCatalog: (id) => {
    appState.worksCatalog = appState.worksCatalog.filter(w => w.id !== id);
    Store.save();
  },

  // Expenses
  addExpense: (expense) => {
    appState.expenses.push({ ...expense, id: generateId() });
    Store.save();
  },

  updateExpense: (id, updates) => {
    const idx = appState.expenses.findIndex(e => e.id === id);
    if (idx !== -1) {
      appState.expenses[idx] = { ...appState.expenses[idx], ...updates };
      Store.save();
    }
  },

  deleteExpense: (id) => {
    appState.expenses = appState.expenses.filter(e => e.id !== id);
    Store.save();
  },

  // Clients
  addClient: (client) => {
    appState.clients.push(normalizeClientRecord({ ...client, id: generateId(), isActive: true }));
    Store.save();
  },

  updateClient: (id, updates) => {
    const idx = appState.clients.findIndex(c => c.id === id);
    if (idx !== -1) {
      appState.clients[idx] = normalizeClientRecord(updates, appState.clients[idx]);
      Store.save();
    }
  },

  deleteClient: (id) => {
    appState.clients = appState.clients.filter(c => c.id !== id);
    Store.save();
  },

  reorderClients: (newOrderedIds) => {
    const reordered = newOrderedIds.map(id => appState.clients.find(c => c.id === id)).filter(Boolean);
    const missing = appState.clients.filter(c => !newOrderedIds.includes(c.id));
    appState.clients = [...reordered, ...missing];
    Store.save();
  },

  toggleClientStatus: (id, month = appState.selectedMonth) => {
    const client = appState.clients.find(c => c.id === id);
    if (client) {
      const monthSettings = ensureMonthSettings(month);
      const currentValue = Store.isClientActiveInMonth(id, month);
      monthSettings.clients[id] = !currentValue;
      Store.save();
    }
  },

  setClientStatusForMonth: (id, isActive, month = appState.selectedMonth) => {
    const client = appState.clients.find(c => c.id === id);
    if (client) {
      const monthSettings = ensureMonthSettings(month);
      monthSettings.clients[id] = isActive !== false;
      Store.save();
    }
  },

  // Config
  updateConfig: (newConfig) => {
    appState.config = { ...appState.config, ...newConfig };
    Store.save();
  },

  updateSettlementMonthConfig: (monthConfig = {}, personContractCharges = {}, month = appState.selectedMonth) => {
    const monthSettings = ensureMonthSettings(month);
    const nextSettlementConfig = {};

    if (monthConfig && Object.prototype.hasOwnProperty.call(monthConfig, 'taxRate')) {
      const taxRate = parseFloat(monthConfig.taxRate);
      if (Number.isFinite(taxRate)) nextSettlementConfig.taxRate = taxRate;
    }

    if (monthConfig && Object.prototype.hasOwnProperty.call(monthConfig, 'zusFixedAmount')) {
      const zusFixedAmount = parseFloat(monthConfig.zusFixedAmount);
      if (Number.isFinite(zusFixedAmount)) nextSettlementConfig.zusFixedAmount = zusFixedAmount;
    }

    monthSettings.settlementConfig = nextSettlementConfig;

    const nextPersonContractCharges = {};
    Object.entries(personContractCharges || {}).forEach(([personId, values]) => {
      const contractTaxAmount = parseFloat(values?.contractTaxAmount);
      const contractZusAmount = parseFloat(values?.contractZusAmount);
      const entry = {};

      if (Number.isFinite(contractTaxAmount)) entry.contractTaxAmount = contractTaxAmount;
      if (Number.isFinite(contractZusAmount)) entry.contractZusAmount = contractZusAmount;

      if (Object.keys(entry).length > 0) {
        nextPersonContractCharges[personId] = entry;
      }
    });

    monthSettings.personContractCharges = nextPersonContractCharges;
    Store.save();
  },

  updateInvoiceMonthConfig: (invoiceConfig = {}, month = appState.selectedMonth) => {
    const monthSettings = ensureMonthSettings(month);
    monthSettings.invoices = normalizeInvoicesConfig({
      ...monthSettings.invoices,
      ...invoiceConfig,
      clients: {
        ...(monthSettings.invoices?.clients || {}),
        ...(invoiceConfig.clients || {})
      }
    });
    Store.save();
  },

  updateInvoiceClientConfig: (clientId, clientConfig = {}, month = appState.selectedMonth) => {
    if (!clientId) return;
    const monthSettings = ensureMonthSettings(month);
    monthSettings.invoices = normalizeInvoicesConfig(monthSettings.invoices || {});
    monthSettings.invoices.clients[clientId] = normalizeInvoiceClientConfig({
      ...(monthSettings.invoices.clients[clientId] || {}),
      ...clientConfig
    });
    Store.save();
  },

  addInvoiceExtraInvoice: (invoice, month = appState.selectedMonth) => {
    const monthSettings = ensureMonthSettings(month);
    monthSettings.invoices = normalizeInvoicesConfig(monthSettings.invoices || {});
    monthSettings.invoices.extraInvoices.push(normalizeExtraInvoiceRecord(invoice));
    Store.save();
  },

  updateInvoiceExtraInvoice: (invoiceId, updates = {}, month = appState.selectedMonth) => {
    if (!invoiceId) return;
    const monthSettings = ensureMonthSettings(month);
    monthSettings.invoices = normalizeInvoicesConfig(monthSettings.invoices || {});
    const index = monthSettings.invoices.extraInvoices.findIndex(invoice => invoice.id === invoiceId);
    if (index === -1) return;
    monthSettings.invoices.extraInvoices[index] = normalizeExtraInvoiceRecord(updates, monthSettings.invoices.extraInvoices[index]);
    Store.save();
  },

  deleteInvoiceExtraInvoice: (invoiceId, month = appState.selectedMonth) => {
    if (!invoiceId) return;
    const monthSettings = ensureMonthSettings(month);
    monthSettings.invoices = normalizeInvoicesConfig(monthSettings.invoices || {});
    monthSettings.invoices.extraInvoices = monthSettings.invoices.extraInvoices.filter(invoice => invoice.id !== invoiceId);
    Store.save();
  },

  // Settings
  updateSettings: (newSettings) => {
    appState.settings = normalizeSharedSettings({ ...appState.settings, ...newSettings });
    Store.save();
  },

  updateAppearanceSettings: (newSettings) => {
    localSettings = normalizeLocalSettings({ ...localSettings, ...newSettings });
    Store.saveLocalSettings();
  },

  resetAppearanceSettings: (newSettings = createDefaultLocalSettings()) => {
    localSettings = normalizeLocalSettings(newSettings);
    Store.saveLocalSettings();
  },

  importState: (rawState) => {
    const payload = rawState && typeof rawState === 'object' && !Array.isArray(rawState) && rawState.data && typeof rawState.data === 'object'
      ? rawState.data
      : rawState;

    if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
      throw new Error('Nieprawidłowy format pliku z bazą danych.');
    }

    if (!APP_STATE_KEYS.some(key => key in payload)) {
      throw new Error('Plik nie zawiera danych aplikacji.');
    }

    appState = normalizeState(payload);
    ensureMonthSettings(appState.selectedMonth);
    Store.save();
    return appState;
  }
};
