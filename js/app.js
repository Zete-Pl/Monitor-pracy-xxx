console.log("MONITOR PRACY APP.JS V3 - LOADED");
// Firebase Configuration
const firebaseConfig = {
  apiKey: "AIzaSyAER1cBnWBVFjyMiEcZTKwgt8f-EgvNEVA",
  authDomain: "monitor-pracy-a9440.firebaseapp.com",
  databaseURL: "https://monitor-pracy-a9440-default-rtdb.europe-west1.firebasedatabase.app",
  projectId: "monitor-pracy-a9440",
  storageBucket: "monitor-pracy-a9440.firebasestorage.app",
  messagingSenderId: "751538429834",
  appId: "1:751538429834:web:a20261f25896fca8727273"
};

if (typeof firebase !== 'undefined' && !firebase.apps.length) {
  firebase.initializeApp(firebaseConfig);
}

function appIsFirebasePermissionDeniedError(error) {
  const code = (error?.code || '').toString().toUpperCase();
  const message = (error?.message || '').toString().toLowerCase();
  return code === 'PERMISSION_DENIED' || message.includes('permission_denied');
}

function appBuildSharedDataOnlyUpdates(updates = {}) {
  return Object.fromEntries(
    Object.entries(updates).filter(([path]) => path.startsWith('/shared_data/'))
  );
}

function appUpdateFirebaseRootWithFallback(updates = {}) {
  if (typeof firebase === 'undefined' || !firebase.database) return Promise.resolve();

  if (typeof Store !== 'undefined' && Store.recordFirebaseUpload) {
    Store.recordFirebaseUpload(updates);
  }

  return firebase.database().ref().update(updates).catch(error => {
    if (!appIsFirebasePermissionDeniedError(error)) {
      throw error;
    }

    const legacyUpdates = appBuildSharedDataOnlyUpdates(updates);
    if (Object.keys(legacyUpdates).length === 0) {
      throw error;
    }

    console.warn('firebase: Brak uprawnień do nowych gałęzi meta/historii. Używam zgodności tylko z /shared_data.', error);
    if (typeof Store !== 'undefined' && Store.recordFirebaseUpload) {
      Store.recordFirebaseUpload(legacyUpdates);
    }
    return firebase.database().ref().update(legacyUpdates);
  });
}

function getTrackedFirebaseSnapshotValue(snapshot) {
  const value = snapshot?.val ? snapshot.val() : null;
  if (typeof Store !== 'undefined' && Store.recordFirebaseDownload) {
    Store.recordFirebaseDownload(value);
  }
  return value;
}

function formatFirebaseTransferBytes(bytes = 0) {
  const normalizedBytes = Math.max(0, parseFloat(bytes) || 0);
  if (normalizedBytes < 1024) return `${normalizedBytes.toFixed(0)} B`;
  if (normalizedBytes < (1024 * 1024)) return `${(normalizedBytes / 1024).toFixed(1)} KB`;
  return `${(normalizedBytes / (1024 * 1024)).toFixed(2)} MB`;
}

function updateSettingsFirebaseTransferStats() {
  const container = document.getElementById('settings-firebase-transfer-stats');
  const downloadedEl = document.getElementById('settings-firebase-bytes-down');
  const uploadedEl = document.getElementById('settings-firebase-bytes-up');
  const readsEl = document.getElementById('settings-firebase-read-count');
  const writesEl = document.getElementById('settings-firebase-write-count');
  const resetButton = document.getElementById('btn-reset-firebase-transfer-stats');

  if (!container || !downloadedEl || !uploadedEl || !readsEl || !writesEl) return;

  const isOnline = !window.isOfflineMode;
  container.style.display = isOnline ? 'block' : 'none';
  if (resetButton) resetButton.style.display = isOnline ? '' : 'none';
  if (!isOnline || typeof Store === 'undefined' || !Store.getFirebaseTransferStats) return;

  const stats = Store.getFirebaseTransferStats();
  downloadedEl.textContent = formatFirebaseTransferBytes(stats.downloadBytes || 0);
  uploadedEl.textContent = formatFirebaseTransferBytes(stats.uploadBytes || 0);
  readsEl.textContent = `${stats.readCount || 0}`;
  writesEl.textContent = `${stats.writeCount || 0}`;
}

function buildFirebaseStateUpdates(state = {}, syncMetadata = {}) {
  if (typeof Store !== 'undefined' && Store.buildFirebaseRootUpdatesFromStateBundle) {
    return Store.buildFirebaseRootUpdatesFromStateBundle(state, syncMetadata);
  }

  return {
    '/shared_data/common': state?.common || {},
    '/shared_meta/common': syncMetadata?.common || {}
  };
}

function buildMergedSyncMetadataFromState(state = null, existingMeta = {}) {
  if (typeof Store === 'undefined' || !Store.buildSyncMetadataFromState || !Store.normalizeSyncMetadata) {
    return { meta: existingMeta || {}, changed: false };
  }

  const rebuiltMeta = Store.buildSyncMetadataFromState(state || { version: 'v3', common: {}, months: {} }) || { common: {}, months: {} };
  const mergedMeta = Store.normalizeSyncMetadata(existingMeta || {});
  let changed = false;

  const hasUsableEntry = (entry = {}) => (parseInt(entry?.revision, 10) || 0) > 0 && (entry?.updatedAt || '').toString().trim() !== '';

  if (hasUsableEntry(rebuiltMeta.common) && !hasUsableEntry(mergedMeta.common)) {
    mergedMeta.common = rebuiltMeta.common;
    changed = true;
  }

  Object.entries(rebuiltMeta.months || {}).forEach(([monthKey, monthMeta]) => {
    if (!mergedMeta.months) mergedMeta.months = {};
    if (!mergedMeta.months[monthKey]) {
      mergedMeta.months[monthKey] = {};
      changed = true;
    }

    ['monthlySheets', 'worksSheets', 'expenses'].forEach(scopeKey => {
      if (!hasUsableEntry(monthMeta?.[scopeKey]) || hasUsableEntry(mergedMeta.months?.[monthKey]?.[scopeKey])) return;
      mergedMeta.months[monthKey][scopeKey] = monthMeta[scopeKey];
      changed = true;
    });

    if (hasUsableEntry(monthMeta?.monthSettings) && !hasUsableEntry(mergedMeta.months?.[monthKey]?.monthSettings)) {
      mergedMeta.months[monthKey].monthSettings = {
        ...(mergedMeta.months[monthKey].monthSettings || {}),
        ...monthMeta.monthSettings
      };
      changed = true;
    }

    if (hasUsableEntry(monthMeta?.monthSettings?.commonSnapshot) && !hasUsableEntry(mergedMeta.months?.[monthKey]?.monthSettings?.commonSnapshot)) {
      if (!mergedMeta.months[monthKey].monthSettings) mergedMeta.months[monthKey].monthSettings = {};
      mergedMeta.months[monthKey].monthSettings.commonSnapshot = monthMeta.monthSettings.commonSnapshot;
      changed = true;
    }
  });

  return { meta: mergedMeta, changed };
}

function summarizeSyncState(state = {}) {
  const monthKeys = Object.keys(state?.months || {}).sort((a, b) => b.localeCompare(a, 'pl-PL'));
  const currentMonth = monthKeys[0] || 'brak';
  const currentMonthData = monthKeys[0] ? (state.months?.[monthKeys[0]] || {}) : {};

  return [
    `osoby: ${(state?.common?.persons || []).length}`,
    `klienci: ${(state?.common?.clients || []).length}`,
    `miesiące: ${monthKeys.length}`,
    `najnowszy miesiąc: ${currentMonth}`,
    `arkusze godzin: ${(currentMonthData?.monthlySheets || []).length}`,
    `arkusze prac: ${(currentMonthData?.worksSheets || []).length}`,
    `koszty: ${(currentMonthData?.expenses || []).length}`
  ].join(', ');
}

function isEffectivelyEmptyLocalSyncState(state = {}) {
  const common = state?.common || {};
  const months = state?.months || {};

  if ((common.persons || []).length > 0) return false;
  if ((common.clients || []).length > 0) return false;

  return !Object.values(months).some(monthRecord => {
    const monthSettings = monthRecord?.monthSettings || {};
    return (monthRecord?.monthlySheets || []).length > 0
      || (monthRecord?.worksSheets || []).length > 0
      || (monthRecord?.expenses || []).length > 0
      || Object.keys(monthSettings.persons || {}).length > 0
      || Object.keys(monthSettings.clients || {}).length > 0
      || Object.keys(monthSettings.settlementConfig || {}).length > 0
      || Object.keys(monthSettings.personContractCharges || {}).length > 0
      || Object.keys(monthSettings.invoices?.clients || {}).length > 0
      || (monthSettings.invoices?.extraInvoices || []).length > 0;
  });
}

function formatSyncConflictEntries(entries = [], label = '') {
  if (!Array.isArray(entries) || entries.length === 0) {
    return `${label}: brak.`;
  }

  const lines = entries.slice(0, 8).map(entry => {
    const monthLabel = entry.month ? ` [${entry.month}]` : '';
    const localUpdatedAt = entry.localEntry?.updatedAt ? formatHistoryTimestamp(entry.localEntry.updatedAt) : 'brak';
    const remoteUpdatedAt = entry.remoteEntry?.updatedAt ? formatHistoryTimestamp(entry.remoteEntry.updatedAt) : 'brak';
    return `- ${entry.scopeLabel}${monthLabel}: lokalnie r${entry.localEntry?.revision || 0} (${localUpdatedAt}) vs online r${entry.remoteEntry?.revision || 0} (${remoteUpdatedAt})`;
  });

  if (entries.length > 8) {
    lines.push(`- ... oraz ${entries.length - 8} kolejnych zakresów`);
  }

  return `${label}:\n${lines.join('\n')}`;
}

function getSyncConflictDialogElements() {
  return {
    overlay: document.getElementById('sync-conflict-dialog-overlay'),
    dialog: document.getElementById('sync-conflict-dialog'),
    title: document.getElementById('sync-conflict-dialog-title'),
    stepLabel: document.getElementById('sync-conflict-dialog-step'),
    body: document.getElementById('sync-conflict-dialog-body'),
    details: document.getElementById('sync-conflict-dialog-details'),
    btnRemote: document.getElementById('btn-sync-conflict-remote'),
    btnLocal: document.getElementById('btn-sync-conflict-local')
  };
}

function setSyncConflictDialogVisibility(isVisible = false) {
  const { overlay } = getSyncConflictDialogElements();
  if (!overlay) return;

  overlay.style.display = isVisible ? 'flex' : 'none';
  overlay.setAttribute('aria-hidden', isVisible ? 'false' : 'true');
  document.body.classList.toggle('sync-conflict-dialog-open', isVisible);
}

function buildSyncConflictDialogBodyHtml(lines = []) {
  return (Array.isArray(lines) ? lines : [])
    .map(line => {
      if (!line) return '<div class="sync-conflict-dialog-spacer"></div>';
      return `<p>${escapeHistoryPreviewText(line)}</p>`;
    })
    .join('');
}

function showSyncConflictDecisionDialog(options = {}) {
  const {
    title = 'Konflikt synchronizacji danych',
    stepLabel = '',
    bodyLines = [],
    detailText = '',
    preferredChoice = 'remote',
    remoteLabel = 'Pobierz Bazę Główną (zalecane)',
    localLabel = 'Zachowaj Bazę Lokalną (nie zalecane)'
  } = options;
  const elements = getSyncConflictDialogElements();
  const { overlay, title: titleEl, stepLabel: stepLabelEl, body, details, btnRemote, btnLocal } = elements;

  if (!overlay || !titleEl || !stepLabelEl || !body || !details || !btnRemote || !btnLocal) {
    const fallbackText = [title, '', ...bodyLines, detailText ? `\n${detailText}` : ''].filter(Boolean).join('\n');
    return Promise.resolve(window.confirm(fallbackText) ? 'remote' : 'local');
  }

  return new Promise(resolve => {
    const previousActiveElement = document.activeElement instanceof HTMLElement ? document.activeElement : null;

    const cleanup = () => {
      btnRemote.onclick = null;
      btnLocal.onclick = null;
      overlay.onkeydown = null;
      setSyncConflictDialogVisibility(false);
      previousActiveElement?.focus?.();
    };

    const finish = (choice) => {
      cleanup();
      resolve(choice);
    };

    titleEl.textContent = title;
    stepLabelEl.textContent = stepLabel;
    body.innerHTML = buildSyncConflictDialogBodyHtml(bodyLines);
    details.innerHTML = detailText
      ? `<pre class="sync-conflict-dialog-pre">${escapeHistoryPreviewText(detailText)}</pre>`
      : '';
    details.style.display = detailText ? 'block' : 'none';

    btnRemote.textContent = remoteLabel;
    btnLocal.textContent = localLabel;

    btnRemote.onclick = () => finish('remote');
    btnLocal.onclick = () => finish('local');
    overlay.onkeydown = (event) => {
      if (event.key === 'Escape') {
        event.preventDefault();
      }
    };

    setSyncConflictDialogVisibility(true);
    window.setTimeout(() => {
      (preferredChoice === 'local' ? btnLocal : btnRemote).focus();
    }, 0);
  });
}

function askUserHowToResolveSyncConflict(conflictInfo = {}, localState = {}, remoteState = {}) {
  const localSummary = summarizeSyncState(localState);
  const remoteSummary = summarizeSyncState(remoteState);

  const detailedComparisonText = [
    `Lokalna baza: ${localSummary}`,
    `Baza Główna na serwerze: ${remoteSummary}`,
    '',
    formatSyncConflictEntries(conflictInfo.localNewer || [], 'Zakresy nowsze lokalnie'),
    '',
    formatSyncConflictEntries(conflictInfo.remoteNewer || [], 'Zakresy nowsze na serwerze')
  ].join('\n');

  return showSyncConflictDecisionDialog({
    title: 'Wykryto nowszą bazę lokalną',
    stepLabel: 'Krok 1 z 2 • Ważny komunikat',
    bodyLines: [
      'Baza Danych w urządzeniu jest nowsza od Bazy Danych Głównej na serwerze.',
      'Zalecane jest pobranie Bazy Danych Głównej z serwera.',
      'Opcja druga służy w awaryjnych sytuacjach, kiedy nie było internetu, a trzeba było zapisać zmiany.'
    ],
    preferredChoice: 'remote'
  }).then(initialChoice => {
    if (initialChoice === 'remote') {
      return 'remote';
    }

    return showSyncConflictDecisionDialog({
    title: 'Szczegóły różnic między bazami danych',
    stepLabel: 'Krok 2 z 2 • Szczegóły i ostateczna decyzja',
    bodyLines: [
      initialChoice === 'remote'
        ? 'W pierwszym kroku wybrano wstępnie pobranie Bazy Głównej z serwera. Poniżej są szczegóły różnic — możesz jeszcze zmienić decyzję.'
        : 'W pierwszym kroku wybrano wstępnie zachowanie Bazy Lokalnej. To opcja awaryjna — poniżej są szczegóły różnic i nadal możesz zmienić decyzję.',
      '',
      'Poniżej pokazano czym różnią się obie bazy danych.'
    ],
    detailText: detailedComparisonText,
    preferredChoice: initialChoice
    });
  });
}

function tryBootstrapSharedMetaFromSharedData(sharedData = null, existingMeta = {}) {
  if (typeof firebase === 'undefined' || !firebase.database || typeof Store === 'undefined' || !Store.buildSyncMetadataFromState) {
    return Promise.resolve({ meta: existingMeta || {}, bootstrapped: false });
  }

  const normalizedState = {
    version: 'v3',
    common: sharedData?.common || {},
    months: sharedData?.months || {}
  };
  const mergedMetaResult = buildMergedSyncMetadataFromState(normalizedState, existingMeta || {});

  if (!mergedMetaResult.changed) {
    return Promise.resolve({ meta: mergedMetaResult.meta, bootstrapped: false });
  }

  const remoteSyncMetadata = (typeof Store !== 'undefined' && Store.serializeSyncMetadataForRemote)
    ? Store.serializeSyncMetadataForRemote(mergedMetaResult.meta)
    : mergedMetaResult.meta;

  const updates = {
    '/shared_meta/common': remoteSyncMetadata.common || {}
  };

  Object.keys(normalizedState.months || {}).forEach(monthKey => {
    updates[`/shared_meta/months/${monthKey}`] = remoteSyncMetadata.months?.[monthKey] || {};
  });

  Store.recordFirebaseUpload?.(updates);
  return firebase.database().ref().update(updates).then(() => {
    console.log('firebase: Utworzono lub uzupełniono brakujące shared_meta na podstawie /shared_data.');
    return { meta: mergedMetaResult.meta, bootstrapped: true };
  }).catch(error => {
    if (appIsFirebasePermissionDeniedError(error)) {
      console.warn('firebase: Nie udało się utworzyć shared_meta z powodu braku uprawnień. Pozostaję przy synchronizacji z /shared_data.', error);
      return { meta: mergedMetaResult.meta, bootstrapped: false };
    }

    console.error('firebase: Błąd podczas bootstrapu shared_meta:', error);
    return { meta: mergedMetaResult.meta, bootstrapped: false };
  });
}

const FORCE_RESYNC_CONTROL_PATH = 'shared_control/force_resync';
const FORCE_RESYNC_ADMIN_UNLOCK_TARGET = 10;

let settingsForceResyncUnlockCount = 0;
let settingsForceResyncControlsUnlocked = false;

function getCurrentFirebaseUserEmail() {
  if (typeof firebase === 'undefined' || !firebase.auth) return '';
  return (firebase.auth().currentUser?.email || '').toString().toLowerCase().trim();
}

function generateForceResyncRequestId() {
  return `force-resync-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

function encodeForceResyncEmailKey(email = '') {
  return (email || '')
    .toString()
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, char => `_${char.charCodeAt(0).toString(16)}_`);
}

function normalizeForceResyncCompletedEmails(completedEmails = {}) {
  return Object.entries(completedEmails || {}).reduce((acc, [key, value]) => {
    const email = (value?.email || value?.userEmail || '').toString().toLowerCase().trim();
    if (!email) return acc;

    acc[key] = {
      email,
      completedAt: (value?.completedAt || '').toString().trim(),
      requestId: (value?.requestId || value?.token || '').toString().trim()
    };
    return acc;
  }, {});
}

function normalizeForceResyncControl(raw = {}) {
  return {
    enabled: raw?.enabled === true,
    requestedAt: (raw?.requestedAt || '').toString().trim(),
    requestedBy: (raw?.requestedBy || '').toString().toLowerCase().trim(),
    requestId: (raw?.requestId || raw?.token || '').toString().trim(),
    completedEmails: normalizeForceResyncCompletedEmails(raw?.completedEmails || {})
  };
}

function getForceResyncCompletedEntry(control = {}, email = '') {
  const normalizedEmail = (email || '').toString().toLowerCase().trim();
  if (!normalizedEmail) return null;

  const completedByKey = control?.completedEmails?.[encodeForceResyncEmailKey(normalizedEmail)] || null;
  if (completedByKey?.email === normalizedEmail) return completedByKey;

  return Object.values(control?.completedEmails || {}).find(entry => entry?.email === normalizedEmail) || null;
}

function getForceResyncCompletedEmailsForCurrentRequest(control = {}) {
  return Object.values(control?.completedEmails || {})
    .filter(entry => {
      if (!entry?.email) return false;
      if (!control?.requestId) return true;
      return entry.requestId === control.requestId;
    })
    .map(entry => entry.email)
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b, 'pl-PL'));
}

function shouldForceResyncForEmail(control = {}, email = '') {
  const normalizedEmail = (email || '').toString().toLowerCase().trim();
  if (!control?.enabled || !normalizedEmail) return false;

  const completedEntry = getForceResyncCompletedEntry(control, normalizedEmail);
  if (!completedEntry) return true;
  if (control.requestId && completedEntry.requestId !== control.requestId) return true;
  return false;
}

function setForceResyncAdminButtonsVisibility(isVisible = false) {
  const container = document.getElementById('settings-force-resync-admin');
  if (!container) return;
  container.style.display = isVisible ? 'block' : 'none';
}

function setForceResyncAdminStatus(message = '', type = '') {
  const statusEl = document.getElementById('settings-force-resync-admin-status');
  if (!statusEl) return;

  statusEl.textContent = message;
  statusEl.classList.remove('is-success', 'is-error');
  if (type === 'success') statusEl.classList.add('is-success');
  if (type === 'error') statusEl.classList.add('is-error');
}

function refreshForceResyncAdminStatus() {
  if (!settingsForceResyncControlsUnlocked) return Promise.resolve();

  if (window.isOfflineMode) {
    setForceResyncAdminStatus('Tryb offline: force_resync jest niedostępny.', 'error');
    return Promise.resolve();
  }

  const userEmail = getCurrentFirebaseUserEmail();
  if (!userEmail) {
    setForceResyncAdminStatus('Zaloguj się, aby sprawdzić status force_resync.');
    return Promise.resolve();
  }

  if (typeof firebase === 'undefined' || !firebase.database) {
    setForceResyncAdminStatus('Firebase Database jest niedostępny.', 'error');
    return Promise.resolve();
  }

  return firebase.database().ref(FORCE_RESYNC_CONTROL_PATH).once('value').then(snapshot => {
    const control = normalizeForceResyncControl(getTrackedFirebaseSnapshotValue(snapshot) || {});
    if (!control.enabled) {
      setForceResyncAdminStatus('force_resync jest OFF.', 'success');
      return;
    }

    const completedEntry = getForceResyncCompletedEntry(control, userEmail);
    const hasCompletedCurrentRequest = !!completedEntry && (!control.requestId || completedEntry.requestId === control.requestId);
    const parts = ['force_resync jest ON'];

    if (control.requestedAt) {
      parts.push(`od ${formatHistoryTimestamp(control.requestedAt)}`);
    }

    parts.push(hasCompletedCurrentRequest
      ? 'Ten e-mail wykonał już reset dla bieżącego żądania.'
      : 'Ten e-mail jeszcze nie wykonał resetu dla bieżącego żądania.');

    const completedEmails = getForceResyncCompletedEmailsForCurrentRequest(control);
    parts.push(`Wykonali już reset: ${completedEmails.length > 0 ? completedEmails.join(', ') : 'nikt'}.`);

    setForceResyncAdminStatus(parts.join(' • '), hasCompletedCurrentRequest ? 'success' : '');
  }).catch(error => {
    console.error('firebase: Błąd odczytu statusu force_resync:', error);
    setForceResyncAdminStatus('Nie udało się pobrać statusu force_resync.', 'error');
  });
}

function registerSettingsForceResyncUnlockNavigation(targetId = '') {
  if (settingsForceResyncControlsUnlocked) return;

  if (targetId === 'settings-view') {
    settingsForceResyncUnlockCount += 1;
    if (settingsForceResyncUnlockCount >= FORCE_RESYNC_ADMIN_UNLOCK_TARGET) {
      settingsForceResyncControlsUnlocked = true;
      setForceResyncAdminButtonsVisibility(true);
      setForceResyncAdminStatus('Ukryte akcje administracyjne zostały odblokowane.');
      refreshForceResyncAdminStatus();
    }
    return;
  }

  settingsForceResyncUnlockCount = 0;
  setForceResyncAdminButtonsVisibility(false);
}

function setForceResyncAdminButtonsDisabled(isDisabled = false) {
  [
    document.getElementById('btn-force-resync-off'),
    document.getElementById('btn-force-resync-run')
  ].forEach(button => {
    if (button) button.disabled = isDisabled;
  });
}

function setForceResyncControlEnabled(enabled = false) {
  if (window.isOfflineMode) {
    return Promise.reject(new Error('Force resync nie działa w trybie offline.'));
  }

  if (typeof firebase === 'undefined' || !firebase.database) {
    return Promise.reject(new Error('Firebase Database jest niedostępny.'));
  }

  const userEmail = getCurrentFirebaseUserEmail();
  if (!userEmail) {
    return Promise.reject(new Error('Zaloguj się, aby zarządzać force_resync.'));
  }

  const db = firebase.database();
  return db.ref(FORCE_RESYNC_CONTROL_PATH).once('value').then(snapshot => {
    const currentControl = normalizeForceResyncControl(getTrackedFirebaseSnapshotValue(snapshot) || {});
    const now = new Date().toISOString();

    if (enabled) {
      const payload = {
        enabled: true,
        requestedAt: now,
        requestedBy: userEmail,
        requestId: generateForceResyncRequestId(),
        completedEmails: {}
      };
      Store.recordFirebaseUpload?.(payload);
      return db.ref(FORCE_RESYNC_CONTROL_PATH).set(payload);
    }

    const payload = {
      ...currentControl,
      enabled: false,
      disabledAt: now,
      disabledBy: userEmail
    };
    Store.recordFirebaseUpload?.(payload);
    return db.ref(FORCE_RESYNC_CONTROL_PATH).set(payload);
  });
}

function markForceResyncCompletedForEmail(forceResyncControl = {}, email = '') {
  const normalizedEmail = (email || '').toString().toLowerCase().trim();
  const emailKey = encodeForceResyncEmailKey(normalizedEmail);
  if (!normalizedEmail || !emailKey || typeof firebase === 'undefined' || !firebase.database) {
    return Promise.resolve();
  }

  const payload = {
    email: normalizedEmail,
    completedAt: new Date().toISOString(),
    requestId: (forceResyncControl?.requestId || '').toString().trim()
  };
  Store.recordFirebaseUpload?.(payload);
  return firebase.database().ref(`${FORCE_RESYNC_CONTROL_PATH}/completedEmails/${emailKey}`).set(payload);
}

function buildRemoteStateFromSharedData(sharedData = {}) {
  if (typeof Migration !== 'undefined' && typeof Migration.isV2State === 'function' && Migration.isV2State(sharedData)) {
    const migratedState = Migration.v2ToV3(sharedData) || {};
    return {
      version: 'v3',
      common: migratedState.common || {},
      months: migratedState.months || {}
    };
  }

  return {
    version: 'v3',
    common: sharedData?.common || {},
    months: sharedData?.months || {}
  };
}

function migrateRemoteMonthCollectionsIfNeeded(monthKey = '', rawMonthData = null) {
  if (!monthKey || !rawMonthData || typeof Store === 'undefined' || !Store.remoteMonthDataNeedsMigration || !Store.remoteMonthDataNeedsMigration(rawMonthData)) {
    return Promise.resolve(false);
  }

  if (typeof firebase === 'undefined' || !firebase.database) {
    return Promise.resolve(false);
  }

  const normalizedMonth = Store.normalizeMonth ? Store.normalizeMonth(rawMonthData) : rawMonthData;
  const rebuiltMeta = Store.buildSyncMetadataFromState
    ? Store.buildSyncMetadataFromState({ version: 'v3', common: {}, months: { [monthKey]: normalizedMonth } })
    : { months: { [monthKey]: {} } };
  const fullUpdates = Store.buildFirebaseRootUpdatesFromStateBundle
    ? Store.buildFirebaseRootUpdatesFromStateBundle({ version: 'v3', common: {}, months: { [monthKey]: normalizedMonth } }, rebuiltMeta)
    : {};
  const monthUpdates = Object.fromEntries(
    Object.entries(fullUpdates).filter(([path]) => path.startsWith(`/shared_data/months/${monthKey}/`) || path.startsWith(`/shared_meta/months/${monthKey}`))
  );

  if (Object.keys(monthUpdates).length === 0) return Promise.resolve(false);

  return appUpdateFirebaseRootWithFallback(monthUpdates).then(() => {
    console.log(`firebase: Zmigrowano strukturę miesiąca ${monthKey} do map kolekcji i uproszczonego shared_meta.`);
    return true;
  }).catch(error => {
    console.error(`firebase: Nie udało się zmigrować struktury miesiąca ${monthKey}:`, error);
    return false;
  });
}

function rebuildEntireRemoteSharedDataStructure() {
  if (typeof firebase === 'undefined' || !firebase.database || typeof Store === 'undefined' || !Store.buildFirebaseRootUpdatesFromStateBundle || !Store.buildSyncMetadataFromState) {
    return Promise.resolve({ state: null, meta: {}, rebuilt: false });
  }

  return firebase.database().ref('shared_data').once('value').then(snapshot => {
    const rawSharedData = getTrackedFirebaseSnapshotValue(snapshot) || {};
    if (!rawSharedData || Object.keys(rawSharedData).length === 0) {
      return { state: null, meta: {}, rebuilt: false };
    }

    const normalizedState = buildRemoteStateFromSharedData(rawSharedData);
    const rebuiltMeta = Store.buildSyncMetadataFromState(normalizedState) || { common: {}, months: {} };
    const updates = Store.buildFirebaseRootUpdatesFromStateBundle(normalizedState, rebuiltMeta);

    return appUpdateFirebaseRootWithFallback(updates).then(() => ({
      state: normalizedState,
      meta: rebuiltMeta,
      rebuilt: true
    }));
  }).catch(error => {
    console.error('firebase: Nie udało się przebudować całej zdalnej struktury shared_data/shared_meta:', error);
    return { state: null, meta: {}, rebuilt: false };
  });
}

function refreshSharedHistoryCachesFromFirebase() {
  if (typeof firebase === 'undefined' || !firebase.database || typeof Store === 'undefined' || !Store.setSharedHistoryCaches) {
    return Promise.resolve();
  }

  return Promise.allSettled([
    loadIndexedFirebaseCollectionWithLegacyFallback(SHARED_HISTORY_INDEX_PATH, 'shared_history', {
      label: 'historii zmian',
      migrationBuilder: (entries) => Store.buildSharedHistoryMigrationUpdates?.(entries)
    }),
    loadIndexedFirebaseCollectionWithLegacyFallback(SHARED_DAILY_BACKUPS_INDEX_PATH, 'shared_backups_daily', {
      label: 'snapshotów dziennych',
      migrationBuilder: (entries) => Store.buildBackupMigrationUpdates?.('daily', entries)
    }),
    loadIndexedFirebaseCollectionWithLegacyFallback(SHARED_MONTHLY_BACKUPS_INDEX_PATH, 'shared_backups_monthly', {
      label: 'snapshotów miesięcznych',
      migrationBuilder: (entries) => Store.buildBackupMigrationUpdates?.('monthly', entries)
    })
  ]).then(([historyResult, dailyResult, monthlyResult]) => {
    Store.setSharedHistoryCaches({
      history: historyResult.status === 'fulfilled' ? (historyResult.value || []) : [],
      dailyBackups: dailyResult.status === 'fulfilled' ? (dailyResult.value || []) : [],
      monthlyBackups: monthlyResult.status === 'fulfilled' ? (monthlyResult.value || []) : []
    }, {
      skipPeriodicBackups: true
    });
  });
}

function loadSharedBackupCachesForLogin() {
  if (typeof firebase === 'undefined' || !firebase.database || typeof Store === 'undefined' || !Store.setSharedHistoryCaches) {
    return Promise.resolve();
  }

  return Promise.allSettled([
    loadIndexedFirebaseCollectionWithLegacyFallback(SHARED_HISTORY_INDEX_PATH, 'shared_history', {
      label: 'historii zmian',
      migrationBuilder: (entries) => Store.buildSharedHistoryMigrationUpdates?.(entries),
      maxEntries: SHARED_HISTORY_REMOTE_LIMIT,
      entriesPath: SHARED_HISTORY_ENTRIES_PATH
    }),
    loadIndexedFirebaseCollectionWithLegacyFallback(SHARED_DAILY_BACKUPS_INDEX_PATH, 'shared_backups_daily', {
      label: 'snapshotów dziennych',
      migrationBuilder: (entries) => Store.buildBackupMigrationUpdates?.('daily', entries)
    }),
    loadIndexedFirebaseCollectionWithLegacyFallback(SHARED_MONTHLY_BACKUPS_INDEX_PATH, 'shared_backups_monthly', {
      label: 'snapshotów miesięcznych',
      migrationBuilder: (entries) => Store.buildBackupMigrationUpdates?.('monthly', entries)
    })
  ]).then(([historyResult, dailyResult, monthlyResult]) => {
    Store.setSharedHistoryCaches({
      history: historyResult.status === 'fulfilled' ? (historyResult.value || []) : [],
      dailyBackups: dailyResult.status === 'fulfilled' ? (dailyResult.value || []) : [],
      monthlyBackups: monthlyResult.status === 'fulfilled' ? (monthlyResult.value || []) : []
    }, {
      skipPeriodicBackups: true
    });
    Store.ensurePeriodicBackupsForLogin?.();
  }).catch(error => {
    console.error('firebase: Nie udało się pobrać snapshotów bezpieczeństwa po zalogowaniu:', error);
    Store.ensurePeriodicBackupsForLogin?.();
  });
}

function loadSharedHistoryManagementData() {
  if (typeof firebase === 'undefined' || !firebase.database || typeof Store === 'undefined' || !Store.setSharedHistoryCaches) {
    return Promise.resolve();
  }

  return Promise.allSettled([
    loadIndexedFirebaseCollectionWithLegacyFallback(SHARED_HISTORY_INDEX_PATH, 'shared_history', {
      label: 'historii zmian',
      migrationBuilder: (entries) => Store.buildSharedHistoryMigrationUpdates?.(entries)
    }),
    loadIndexedFirebaseCollectionWithLegacyFallback(SHARED_DAILY_BACKUPS_INDEX_PATH, 'shared_backups_daily', {
      label: 'snapshotów dziennych',
      migrationBuilder: (entries) => Store.buildBackupMigrationUpdates?.('daily', entries)
    }),
    loadIndexedFirebaseCollectionWithLegacyFallback(SHARED_MONTHLY_BACKUPS_INDEX_PATH, 'shared_backups_monthly', {
      label: 'snapshotów miesięcznych',
      migrationBuilder: (entries) => Store.buildBackupMigrationUpdates?.('monthly', entries)
    })
  ]).then(([historyResult, dailyResult, monthlyResult]) => {
    Store.setSharedHistoryCaches({
      history: historyResult.status === 'fulfilled' ? (historyResult.value || []) : [],
      dailyBackups: dailyResult.status === 'fulfilled' ? (dailyResult.value || []) : [],
      monthlyBackups: monthlyResult.status === 'fulfilled' ? (monthlyResult.value || []) : []
    }, {
      skipPeriodicBackups: true
    });
  });
}

function rebuildLocalDatabaseFromSharedData(sharedData = {}, options = {}) {
  if (typeof firebase === 'undefined' || !firebase.database || typeof Store === 'undefined') {
    return Promise.resolve();
  }

  const {
    userEmail = '',
    forceResyncControl = {},
    markForceResyncCompleted = false,
    refreshSharedCaches = false,
    onAfterRebuild = null
  } = options;
  const remoteState = buildRemoteStateFromSharedData(sharedData);

  return firebase.database().ref('shared_meta').once('value').catch(error => {
    if (appIsFirebasePermissionDeniedError(error)) {
      console.warn('firebase: Brak uprawnień do shared_meta podczas pełnego resetu lokalnej bazy. Odbudowuję lokalne meta z /shared_data.', error);
      return { val: () => ({}) };
    }
    throw error;
  }).then(metaSnapshot => {
    return tryBootstrapSharedMetaFromSharedData(remoteState, getTrackedFirebaseSnapshotValue(metaSnapshot) || {}).then(({ meta }) => {
      window.currentSheetId = null;
      window.currentWorksSheetId = null;

      if (Store.clearLocalDatabaseStorage) {
        Store.clearLocalDatabaseStorage();
      }

      if (Store.applyRemoteStateBundle) {
        window.isImportingFromFirebase = true;
        try {
          Store.applyRemoteStateBundle(remoteState, meta || {});
        } finally {
          window.isImportingFromFirebase = false;
        }
      }

      const followUpTasks = [];

      if (markForceResyncCompleted) {
        followUpTasks.push(
          markForceResyncCompletedForEmail(forceResyncControl, userEmail).catch(error => {
            console.error('firebase: Nie udało się zapisać wykonania force_resync dla e-maila:', error);
          })
        );
      }

      if (refreshSharedCaches) {
        followUpTasks.push(
          refreshSharedHistoryCachesFromFirebase().catch(error => {
            console.error('firebase: Nie udało się odświeżyć lokalnych cache historii po pobraniu Bazy Głównej:', error);
          })
        );
      }

      return Promise.all(followUpTasks).then(() => {
        if (markForceResyncCompleted) {
          console.log('firebase: Wykonano force_resync dla użytkownika:', userEmail);
        }

        if (typeof onAfterRebuild === 'function') {
          onAfterRebuild();
        }

        return { state: remoteState, meta: meta || {} };
      });
    });
  });
}

function resetLocalDatabaseAndFetchMainDatabase() {
  if (window.isOfflineMode) {
    return Promise.reject(new Error('Ta operacja jest dostępna tylko online.'));
  }

  if (typeof firebase === 'undefined' || !firebase.database) {
    return Promise.reject(new Error('Firebase Database jest niedostępny.'));
  }

  if (!getCurrentFirebaseUserEmail()) {
    return Promise.reject(new Error('Zaloguj się, aby pobrać Bazę Główną z Firebase.'));
  }

  return firebase.database().ref('shared_data').once('value').then(snapshot => {
    return rebuildLocalDatabaseFromSharedData(getTrackedFirebaseSnapshotValue(snapshot) || {}, {
      refreshSharedCaches: true
    });
  });
}

function fetchRemoteStateBundleForSharedHistory() {
  if (typeof firebase === 'undefined' || !firebase.database) {
    return Promise.resolve({
      state: { version: 'v3', common: {}, months: {} },
      meta: {}
    });
  }

  return Promise.all([
    firebase.database().ref('shared_data').once('value'),
    firebase.database().ref('shared_meta').once('value').catch(error => {
      if (appIsFirebasePermissionDeniedError(error)) {
        console.warn('firebase: Brak uprawnień do shared_meta podczas przygotowania snapshotu importu. Używam pustych meta.', error);
        return { val: () => ({}) };
      }
      throw error;
    })
  ]).then(([stateSnapshot, metaSnapshot]) => ({
    state: buildRemoteStateFromSharedData(getTrackedFirebaseSnapshotValue(stateSnapshot) || {}),
    meta: getTrackedFirebaseSnapshotValue(metaSnapshot) || {}
  }));
}

function performForceResyncBootstrap(sharedData = {}, userEmail = '', forceResyncControl = {}, setupV3Listeners = () => {}) {
  return rebuildLocalDatabaseFromSharedData(sharedData, {
    userEmail,
    forceResyncControl,
    markForceResyncCompleted: true,
    onAfterRebuild: setupV3Listeners
  });
}

document.addEventListener('DOMContentLoaded', () => {
  console.log("KROK 0: Start DOMContentLoaded");

  const flushPendingRemoteSync = () => {
    Promise.resolve(Store.flushPendingRemoteSync?.()).catch(() => {});
  };
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'hidden') {
      flushPendingRemoteSync();
    }
  });
  window.addEventListener('beforeunload', flushPendingRemoteSync);
  
  // 1. Zawsze inicjalizuj Auth jako pierwsze (musi działać nawet przy błędach danych)
  if (typeof firebase !== 'undefined' && firebase.auth) {
    console.log("KROK 1: Inicjalizacja Autoryzacji Firebase");
    initFirebaseAuth();
  } else {
    console.warn("UWAGA: Firebase nie jest zainicjalizowany lub brak modułu Auth!");
  }

  try {
    console.log("KROK 2: Inicjalizacja komponentów UI");
    initSidebarToggle();
    initHistoryControls();
    initHoursInstructionsControls();
    initGlobalMonthSelector();
    initNavigation();
    initMobileBackNavigation();
    initPeopleManager();
    initMonthlySheets();
    initWorksTracker();
    initClientsTracker();
    initExpensesTracker();
    initSettlement();
    initInvoices();
    initReportsView();
    initPayouts();
    initMobileInputVisibilityHandling();
    initSettings();
    
    console.log("KROK 3: Pierwsze renderowanie danych");
    renderAll();
    applySettings();
    console.log("KROK 4: Aplikacja gotowa");
  } catch (err) {
    console.error("BŁĄD KRYTYCZNY INICJALIZACJI:", err);
    // window.onerror (w index.html) wyświetli szczegóły w okienku alert
  }

  // Re-render on state changes
  window.addEventListener('appStateChanged', () => {
    try {
      applySettings();
      renderAll();
      refreshOpenSheetDetailViews();
    } catch (e) {
      console.error("Błąd podczas przerysowania po zmianie stanu:", e);
    }
  });

  window.addEventListener('resize', () => {
    scheduleMobileCardValueAutoFit();
    scheduleCurrencyValueAutoFit();
  });
});

function initSidebarToggle() {
  const btn = document.getElementById('btn-toggle-sidebar');
  const sidebar = document.getElementById('sidebar');
  const mainContent = document.getElementById('main-content');
  const monthCompact = document.getElementById('global-month-select-compact');

  if (!btn || !sidebar || !mainContent) return;

  const scheduleSidebarNavigationLayoutRefresh = () => {
    requestAnimationFrame(() => {
      updateSidebarNavigationLayout();
      requestAnimationFrame(updateSidebarNavigationLayout);
      window.setTimeout(updateSidebarNavigationLayout, 120);
      window.setTimeout(updateSidebarNavigationLayout, 260);
    });
  };

  const syncCollapsedSidebarControls = () => {
    const isCollapsed = sidebar.classList.contains('collapsed');
    const showCompactMonth = isCollapsed && !window.matchMedia('(max-width: 550px)').matches;
    if (monthCompact) {
      monthCompact.style.display = showCompactMonth ? 'inline-flex' : 'none';
    }

    const logoutBtn = document.getElementById('btn-sidebar-firebase-logout');
    if (logoutBtn) {
      const logoutLabel = logoutBtn.querySelector('.sidebar-logout-label');
      if (logoutLabel) logoutLabel.style.display = isCollapsed ? 'none' : '';
      const logoutIcon = logoutBtn.querySelector('.sidebar-logout-icon');
      if (logoutIcon) logoutIcon.style.display = isCollapsed ? 'inline-flex' : '';
      logoutBtn.title = isCollapsed ? 'Wyloguj' : '';
      logoutBtn.setAttribute('aria-label', 'Wyloguj');
    }
  };

  const isCollapsed = localStorage.getItem('sidebar-collapsed') === 'true';
  if (isCollapsed) {
    sidebar.classList.add('collapsed');
    mainContent.classList.add('expanded');
    btn.innerHTML = '<i data-lucide="panel-left-open" id="sidebar-toggle-icon"></i>';
  }

  syncCollapsedSidebarControls();
  scheduleSidebarNavigationLayoutRefresh();

  btn.addEventListener('click', () => {
    sidebar.classList.toggle('collapsed');
    mainContent.classList.toggle('expanded');
    
    const currentlyCollapsed = sidebar.classList.contains('collapsed');
    localStorage.setItem('sidebar-collapsed', currentlyCollapsed);
    
    btn.innerHTML = `<i data-lucide="${currentlyCollapsed ? 'panel-left-open' : 'panel-left-close'}" id="sidebar-toggle-icon"></i>`;
    lucide.createIcons();
    syncCollapsedSidebarControls();
    scheduleSidebarNavigationLayoutRefresh();
  });

  sidebar.addEventListener('transitionend', (e) => {
    if (e.propertyName === 'width' || e.propertyName === 'padding-left' || e.propertyName === 'padding-right') {
      scheduleSidebarNavigationLayoutRefresh();
    }
  });

  window.addEventListener('resize', scheduleSidebarNavigationLayoutRefresh);
}

function updateSidebarNavigationLayout() {
  const sidebarTopSection = document.getElementById('sidebar-top-section');
  const navLinks = document.querySelector('.nav-links');
  if (!sidebarTopSection || !navLinks) return;

  if (window.matchMedia('(max-width: 550px)').matches) {
    navLinks.style.top = '';
    navLinks.style.maxHeight = '';
    return;
  }

  const appScale = parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--app-scale')) || 1;
  const topOffset = Math.ceil((sidebarTopSection.getBoundingClientRect().height / appScale) + 15.5);
  navLinks.style.top = `${topOffset}px`;
  navLinks.style.maxHeight = `calc(100% - ${topOffset}px)`;
}

const sortableListInstances = new WeakMap();

function isTouchReorderDevice() {
  return window.matchMedia('(pointer: coarse)').matches || window.matchMedia('(hover: none)').matches;
}

function getSortableRowIds(container, selector = 'tr[data-id]') {
  if (!container) return [];
  return Array.from(container.querySelectorAll(selector))
    .map(row => row.getAttribute('data-id'))
    .filter(Boolean);
}

function removeMobileCardDividerRows(container) {
  if (!container) return;
  container.querySelectorAll(':scope > tr.expense-mobile-divider-row').forEach(row => row.remove());
}

function rebuildMobileCardDividerRows(container, colspan = null) {
  if (!container) return;

  removeMobileCardDividerRows(container);

  const dataRows = Array.from(container.querySelectorAll(':scope > tr[data-id]'));
  if (dataRows.length <= 1) return;

  const resolvedColspan = colspan
    || dataRows.find(row => row.children.length > 0)?.children.length
    || 1;

  dataRows.forEach((row, index) => {
    if (index === dataRows.length - 1) return;
    const dividerTr = document.createElement('tr');
    dividerTr.className = 'expense-mobile-divider-row';
    dividerTr.innerHTML = `<td colspan="${resolvedColspan}" class="expense-mobile-divider-cell"><div class="expense-mobile-divider" aria-hidden="true"></div></td>`;
    row.insertAdjacentElement('afterend', dividerTr);
  });
}

function setupListSortable(container, options = {}) {
  if (!container || typeof Sortable === 'undefined') return null;

  const existingInstance = sortableListInstances.get(container);
  if (existingInstance) {
    existingInstance.destroy();
  }

  const isTouch = isTouchReorderDevice();
  let hadMobileDividersBeforeDrag = false;
  let mobileDividerRefreshFrame = null;
  let lastMobileDividerOrderSignature = '';
  const sortableSelector = options.draggable || 'tr[data-id]';

  const refreshMobileDividerRows = () => {
    if (!(hadMobileDividersBeforeDrag || options.mobileDividers === true)) return;

    const currentOrderSignature = getSortableRowIds(container, sortableSelector).join('|');
    if (currentOrderSignature === lastMobileDividerOrderSignature) return;

    rebuildMobileCardDividerRows(container, options.mobileDividerColspan || null);
    lastMobileDividerOrderSignature = currentOrderSignature;
  };

  const scheduleMobileDividerRowsRefresh = () => {
    if (!(hadMobileDividersBeforeDrag || options.mobileDividers === true)) return;
    if (mobileDividerRefreshFrame) return;

    mobileDividerRefreshFrame = window.requestAnimationFrame(() => {
      mobileDividerRefreshFrame = null;
      refreshMobileDividerRows();
    });
  };

  const instance = new Sortable(container, {
    handle: options.handle || '.drag-handle',
    draggable: sortableSelector,
    animation: 150,
    ghostClass: 'sortable-ghost',
    chosenClass: 'sortable-chosen',
    dragClass: 'sortable-drag-active',
    delayOnTouchOnly: true,
    delay: isTouch ? 260 : 0,
    touchStartThreshold: isTouch ? 10 : 0,
    fallbackTolerance: isTouch ? 8 : 0,
    onMove: (evt, originalEvent) => {
      const moveResult = typeof options.onMove === 'function'
        ? options.onMove(evt, originalEvent)
        : true;

      if (moveResult !== false) {
        scheduleMobileDividerRowsRefresh();
      }

      return moveResult;
    },
    onChange: (evt) => {
      scheduleMobileDividerRowsRefresh();

      if (typeof options.onChange === 'function') {
        options.onChange(evt);
      }
    },
    onStart: (evt) => {
      hadMobileDividersBeforeDrag = container.querySelector(':scope > tr.expense-mobile-divider-row') !== null;
      lastMobileDividerOrderSignature = getSortableRowIds(container, sortableSelector).join('|');

      if (typeof options.onStart === 'function') {
        options.onStart(evt);
      }
    },
    onEnd: (evt) => {
      if (mobileDividerRefreshFrame) {
        window.cancelAnimationFrame(mobileDividerRefreshFrame);
        mobileDividerRefreshFrame = null;
      }

      if (hadMobileDividersBeforeDrag || options.mobileDividers === true) {
        rebuildMobileCardDividerRows(container, options.mobileDividerColspan || null);
      }

      hadMobileDividersBeforeDrag = false;
      lastMobileDividerOrderSignature = '';

      if (typeof options.onEnd === 'function') {
        window.setTimeout(() => {
          options.onEnd(evt);
        }, 0);
      }
    }
  });

  sortableListInstances.set(container, instance);
  return instance;
}

function allowSortOnlyWithinSameDate(evt) {
  const draggedDate = evt.dragged?.dataset?.sortDate || '';
  const relatedDate = evt.related?.dataset?.sortDate || '';
  if (!draggedDate || !relatedDate) return true;
  return draggedDate === relatedDate;
}

function moveWorksEntrySubRowAfterMainRow(mainRow) {
  const entryId = mainRow?.dataset?.id || '';
  if (!entryId) return;

  const subRow = document.querySelector(`#works-entries-body tr.works-hours-row[data-parent-id="${entryId}"]`);
  if (!subRow) return;

  mainRow.insertAdjacentElement('afterend', subRow);
}

function scrollNavigationItemIntoView(targetId = '') {
  const navLinks = document.querySelector('.nav-links');
  if (!navLinks || !targetId) return;

  const activeNavTarget = targetId === 'history-view' ? 'settings-view' : targetId;
  const targetLink = navLinks.querySelector(`.nav-item[data-target="${activeNavTarget}"]`);
  if (!targetLink) return;

  const hasHorizontalOverflow = navLinks.scrollWidth > (navLinks.clientWidth + 1);
  const hasVerticalOverflow = navLinks.scrollHeight > (navLinks.clientHeight + 1);
  const scrollBehavior = document.visibilityState === 'hidden' ? 'auto' : 'smooth';

  if (hasHorizontalOverflow) {
    const targetCenter = targetLink.offsetLeft + (targetLink.offsetWidth / 2);
    const maxScrollLeft = Math.max(0, navLinks.scrollWidth - navLinks.clientWidth);
    const nextScrollLeft = Math.max(0, Math.min(maxScrollLeft, targetCenter - (navLinks.clientWidth / 2)));
    const currentVisibleStart = navLinks.scrollLeft;
    const currentVisibleEnd = currentVisibleStart + navLinks.clientWidth;
    const targetStart = targetLink.offsetLeft;
    const targetEnd = targetStart + targetLink.offsetWidth;
    const isFullyVisible = targetStart >= currentVisibleStart && targetEnd <= currentVisibleEnd;

    if (!isFullyVisible || Math.abs(nextScrollLeft - navLinks.scrollLeft) > 1) {
      navLinks.scrollTo({ left: nextScrollLeft, behavior: scrollBehavior });
    }
    return;
  }

  if (hasVerticalOverflow) {
    const targetCenter = targetLink.offsetTop + (targetLink.offsetHeight / 2);
    const maxScrollTop = Math.max(0, navLinks.scrollHeight - navLinks.clientHeight);
    const nextScrollTop = Math.max(0, Math.min(maxScrollTop, targetCenter - (navLinks.clientHeight / 2)));
    const currentVisibleTop = navLinks.scrollTop;
    const currentVisibleBottom = currentVisibleTop + navLinks.clientHeight;
    const targetTop = targetLink.offsetTop;
    const targetBottom = targetTop + targetLink.offsetHeight;
    const isFullyVisible = targetTop >= currentVisibleTop && targetBottom <= currentVisibleBottom;

    if (!isFullyVisible || Math.abs(nextScrollTop - navLinks.scrollTop) > 1) {
      navLinks.scrollTo({ top: nextScrollTop, behavior: scrollBehavior });
    }
  }
}

function scheduleNavigationItemIntoView(targetId = '') {
  if (!targetId) return;

  const syncNavigationScroll = () => {
    scrollNavigationItemIntoView(targetId);
  };

  requestAnimationFrame(() => {
    syncNavigationScroll();
    requestAnimationFrame(syncNavigationScroll);
  });

  window.setTimeout(syncNavigationScroll, 120);
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

  registerSettingsForceResyncUnlockNavigation(targetId);

   const activeNavTarget = targetId === 'history-view' ? 'settings-view' : targetId;

  navLinks.forEach(link => {
    link.classList.toggle('active', link.getAttribute('data-target') === activeNavTarget);
  });
  views.forEach(view => {
    view.classList.toggle('active', view.id === targetId);
  });

  scheduleNavigationItemIntoView(targetId);

  if (mainContent) {
    mainContent.scrollTop = 0;
  }

  if (targetId === 'hours-view' && window.currentSheetId) {
    renderSheetDetail(window.currentSheetId);
    focusMonthlySheetDetail(window.currentSheetId);
  } else if (targetId === 'works-view' && window.currentWorksSheetId) {
    renderWorksSheetDetail(window.currentWorksSheetId);
  } else if (targetId === 'invoices-view') {
    renderInvoices();
  } else if (targetId === 'reports-view') {
    renderReports();
  } else if (targetId === 'payouts-view') {
    renderPayouts();
  } else if (targetId === 'history-view') {
    loadSharedHistoryManagementData().catch(error => {
      console.error('firebase: Nie udało się odświeżyć historii zmian:', error);
    });
    renderHistoryManagement();
  }

  normalizeCurrencySuffixSpacing(document.querySelector('.view.active') || mainContent);
  scheduleMobileCardValueAutoFit();
  scheduleCurrencyValueAutoFit();
}

let mobileBackNavigationObserver = null;
let mobileBackNavigationSyncTimer = null;
let mobileBackNavigationLastKey = '';
let mobileBackNavigationLastDepth = 0;
let mobileBackNavigationSuppressSync = false;

function isMobileBackNavigationLayout() {
  return window.matchMedia('(max-width: 550px)').matches;
}

function isElementActuallyVisible(element) {
  if (!element) return false;
  return window.getComputedStyle(element).display !== 'none';
}

function getActiveViewId() {
  return document.querySelector('.view.active')?.id || 'summary-view';
}

function getMobileBackNavigationSnapshot() {
  const activeViewId = getActiveViewId();

  if (activeViewId === 'summary-view') {
    return {
      key: 'summary-view',
      depth: 0,
      backAction: null
    };
  }

  if (activeViewId === 'history-view') {
    return {
      key: 'history-view',
      depth: 2,
      backAction: { type: 'click', elementId: 'btn-back-to-settings-from-history' }
    };
  }

  if (activeViewId === 'hours-view') {
    if (isElementActuallyVisible(document.getElementById('sheet-active-persons-form'))) {
      return {
        key: `hours-active-persons:${window.currentSheetId || 'none'}`,
        depth: 3,
        backAction: { type: 'click', elementId: 'btn-close-sheet-active-persons' }
      };
    }

    if (isElementActuallyVisible(document.getElementById('sheet-meta-form'))) {
      return {
        key: `hours-form:${document.getElementById('sheet-id')?.value || 'new'}`,
        depth: 2,
        backAction: { type: 'click', elementId: 'btn-cancel-sheet' }
      };
    }

    if (isElementActuallyVisible(document.getElementById('sheet-detail-container')) && window.currentSheetId) {
      return {
        key: `hours-detail:${window.currentSheetId}`,
        depth: 2,
        backAction: { type: 'click', elementId: 'btn-back-to-sheets' }
      };
    }

    return {
      key: 'hours-list',
      depth: 1,
      backAction: { type: 'navigate', targetId: 'summary-view' }
    };
  }

  if (activeViewId === 'works-view') {
    const cancelWorksEntryButton = document.getElementById('btn-cancel-works-entry');
    if (window.currentWorksSheetId && isElementActuallyVisible(cancelWorksEntryButton)) {
      return {
        key: `works-entry-edit:${window.currentWorksSheetId}:${document.getElementById('works-entry-edit-id')?.value || 'draft'}`,
        depth: 3,
        backAction: { type: 'click', elementId: 'btn-cancel-works-entry' }
      };
    }

    if (isElementActuallyVisible(document.getElementById('works-active-persons-form'))) {
      return {
        key: `works-active-persons:${window.currentWorksSheetId || 'none'}`,
        depth: 3,
        backAction: { type: 'click', elementId: 'btn-close-active-persons' }
      };
    }

    if (isElementActuallyVisible(document.getElementById('works-catalog-container'))) {
      return {
        key: `works-catalog:${window.currentWorksSheetId || 'list'}`,
        depth: window.currentWorksSheetId ? 3 : 2,
        backAction: { type: 'click', elementId: 'btn-close-works-catalog' }
      };
    }

    if (isElementActuallyVisible(document.getElementById('works-meta-form'))) {
      return {
        key: `works-form:${document.getElementById('works-sheet-id')?.value || 'new'}`,
        depth: window.currentWorksSheetId ? 3 : 2,
        backAction: { type: 'click', elementId: 'btn-cancel-works-sheet' }
      };
    }

    if (isElementActuallyVisible(document.getElementById('works-detail-container')) && window.currentWorksSheetId) {
      return {
        key: `works-detail:${window.currentWorksSheetId}`,
        depth: 2,
        backAction: { type: 'click', elementId: 'btn-back-to-works-sheets' }
      };
    }

    return {
      key: 'works-list',
      depth: 1,
      backAction: { type: 'navigate', targetId: 'summary-view' }
    };
  }

  if (activeViewId === 'people-view' && isElementActuallyVisible(document.getElementById('person-form'))) {
    return {
      key: `people-form:${document.getElementById('person-id')?.value || 'new'}`,
      depth: 2,
      backAction: { type: 'click', elementId: 'btn-cancel-person' }
    };
  }

  if (activeViewId === 'clients-view' && isElementActuallyVisible(document.getElementById('client-form'))) {
    return {
      key: `clients-form:${document.getElementById('client-id')?.value || 'new'}`,
      depth: 2,
      backAction: { type: 'click', elementId: 'btn-cancel-client' }
    };
  }

  if (activeViewId === 'expenses-view' && isElementActuallyVisible(document.getElementById('expense-form'))) {
    return {
      key: `expenses-form:${document.getElementById('expense-id')?.value || 'new'}`,
      depth: 2,
      backAction: { type: 'click', elementId: 'btn-cancel-expense' }
    };
  }

  if (activeViewId === 'invoices-view' && isElementActuallyVisible(document.getElementById('btn-cancel-invoice-extra'))) {
    return {
      key: `invoices-extra-form:${document.getElementById('invoice-extra-id')?.value || 'new'}`,
      depth: 2,
      backAction: { type: 'click', elementId: 'btn-cancel-invoice-extra' }
    };
  }

  return {
    key: activeViewId,
    depth: 1,
    backAction: { type: 'navigate', targetId: 'summary-view' }
  };
}

function executeMobileBackNavigationAction(action = null) {
  if (!action || typeof action !== 'object') return false;

  if (action.type === 'click' && action.elementId) {
    const element = document.getElementById(action.elementId);
    if (!element || element.disabled || !isElementActuallyVisible(element)) return false;
    element.click();
    return true;
  }

  if (action.type === 'navigate' && action.targetId) {
    activateNavigationTarget(action.targetId);
    return true;
  }

  return false;
}

function syncMobileBackNavigationHistory(mode = 'auto') {
  const snapshot = getMobileBackNavigationSnapshot();

  if (!window.history?.replaceState) {
    mobileBackNavigationLastKey = snapshot.key;
    mobileBackNavigationLastDepth = snapshot.depth;
    return;
  }

  if (snapshot.key === mobileBackNavigationLastKey) return;

  const statePayload = {
    ...(window.history.state || {}),
    __appMobileNav: true,
    __appMobileKey: snapshot.key,
    __appMobileDepth: snapshot.depth
  };

  let shouldPush = false;
  if (isMobileBackNavigationLayout()) {
    if (mode === 'push') {
      shouldPush = true;
    } else if (mode === 'replace') {
      shouldPush = false;
    } else if (mobileBackNavigationLastKey !== '') {
      shouldPush = snapshot.depth > mobileBackNavigationLastDepth || (mobileBackNavigationLastDepth === 0 && snapshot.depth === 1);
    }
  }

  if (shouldPush && window.history.pushState) {
    window.history.pushState(statePayload, '', window.location.href);
  } else {
    window.history.replaceState(statePayload, '', window.location.href);
  }

  mobileBackNavigationLastKey = snapshot.key;
  mobileBackNavigationLastDepth = snapshot.depth;
}

function scheduleMobileBackNavigationHistorySync(mode = 'auto') {
  if (mobileBackNavigationSuppressSync) return;
  if (mobileBackNavigationSyncTimer) {
    window.clearTimeout(mobileBackNavigationSyncTimer);
  }

  mobileBackNavigationSyncTimer = window.setTimeout(() => {
    mobileBackNavigationSyncTimer = null;
    syncMobileBackNavigationHistory(mode);
  }, 0);
}

function initMobileBackNavigation() {
  if (mobileBackNavigationObserver) return;

  const appRoot = document.getElementById('app-root') || document.body;
  if (!appRoot) return;

  mobileBackNavigationObserver = new MutationObserver(() => {
    scheduleMobileBackNavigationHistorySync();
  });

  mobileBackNavigationObserver.observe(appRoot, {
    attributes: true,
    subtree: true,
    attributeFilter: ['class', 'style']
  });

  window.addEventListener('resize', () => {
    scheduleMobileBackNavigationHistorySync('replace');
  });

  window.addEventListener('popstate', () => {
    if (!isMobileBackNavigationLayout()) return;

    const snapshot = getMobileBackNavigationSnapshot();
    if (snapshot.depth <= 0) return;

    mobileBackNavigationSuppressSync = true;
    executeMobileBackNavigationAction(snapshot.backAction);

    window.setTimeout(() => {
      mobileBackNavigationSuppressSync = false;
      scheduleMobileBackNavigationHistorySync('replace');
    }, 0);
  });

  scheduleMobileBackNavigationHistorySync('replace');
}

function updateMonthPickerRestrictions(input, monthKey) {
  if (!input || !monthKey) return;
  const [year, month] = monthKey.split('-').map(Number);
  const firstDay = `${year}-${String(month).padStart(2, '0')}-01`;
  const lastDay = new Date(year, month, 0).getDate();
  const lastDayIso = `${year}-${String(month).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
  
  input.min = firstDay;
  input.max = lastDayIso;
}

function disableContainerInputs(container, isDisabled = true) {
  if (!container) return;
  const selectors = 'input, select, textarea, button:not(.btn-sidebar-firebase-logout):not(.nav-link):not(#btn-sidebar-firebase-logout)';
  const elements = container.matches?.(selectors)
    ? [container, ...container.querySelectorAll(selectors)]
    : [...container.querySelectorAll(selectors)];

  elements.forEach(el => {
    // Nie blokujemy przełączników nawigacji
    if (el.closest('.sidebar')) return;
    if (el.closest('[data-allow-archived="true"]')) return;
    setElementInteractionState(el, isDisabled);
  });
}

function setElementInteractionState(el, isDisabled = true) {
  if (!el) return;
  el.disabled = isDisabled;
  if (isDisabled) {
    el.style.opacity = '0.6';
    el.style.pointerEvents = 'none';
    if (el.classList.contains('btn-danger')) el.style.display = 'none';
  } else {
    el.style.opacity = '';
    el.style.pointerEvents = '';
    if (el.classList.contains('btn-danger')) el.style.display = '';
  }
}

function applyArchivedReadOnlyMode(isArchived = false) {
  const mainContent = document.querySelector('.main-content');
  if (!mainContent) return;

  disableContainerInputs(mainContent, isArchived);

  const reportsView = document.getElementById('reports-view');
  if (reportsView) {
    disableContainerInputs(reportsView, false);
  }

  if (!isArchived) return;

  [
    document.getElementById('settings-view'),
    document.getElementById('history-view'),
    document.getElementById('btn-back-to-sheets'),
    document.getElementById('btn-back-to-works-sheets'),
    document.getElementById('btn-close-works-catalog'),
    document.getElementById('btn-toggle-settlement-details'),
    document.getElementById('btn-copy-invoices-email'),
    document.getElementById('btn-toggle-invoice-extra-panel')
  ].forEach(el => disableContainerInputs(el, false));

  document.querySelectorAll('.btn-open-sheet, .btn-open-works-sheet').forEach(el => setElementInteractionState(el, false));
  document.querySelectorAll('.summary-shortcut-card[data-shortcut-type="monthly-sheet"], .summary-shortcut-card[data-shortcut-type="works-sheet"]').forEach(el => setElementInteractionState(el, false));
}

let mobileCardAutoFitFrame = null;
let currencyValueAutoFitFrame = null;
let hoursSheetAutoFitFrame = null;

function isMobileCardLayoutActive() {
  return window.matchMedia('(max-width: 550px)').matches;
}

function appendMobileCardDividerRow(tbody, colspan = 1, isLast = false) {
  if (!tbody || isLast) return;
  const dividerTr = document.createElement('tr');
  dividerTr.className = 'expense-mobile-divider-row';
  dividerTr.innerHTML = `<td colspan="${colspan}" class="expense-mobile-divider-cell"><div class="expense-mobile-divider" aria-hidden="true"></div></td>`;
  tbody.appendChild(dividerTr);
}

function getMobileCardAutoFitSelectorList() {
  return [
    '#people-table-body td:nth-child(3)',
    '#people-table-body td:nth-child(4)',
    '#clients-table-body td:nth-child(2)',
    '#hours-sheets-body td:nth-child(2)',
    '#hours-sheets-body td:nth-child(3)',
    '#hours-sheets-body td:nth-child(4)',
    '#works-sheets-body td:nth-child(2)',
    '#works-sheets-body td:nth-child(3)',
    '#works-sheets-body td:nth-child(4)',
    '#expenses-table-body td:nth-child(1)',
    '#expenses-table-body td:nth-child(4)',
    '#expenses-table-body td:nth-child(5)',
    '#expenses-table-body td:nth-child(6)',
    '#works-entries-body td:nth-child(1)',
    '#works-entries-body td:nth-child(3)',
    '#works-entries-body td:nth-child(4)',
    '#works-entries-body td:nth-child(5)',
    '#dash-quick-preview td:nth-child(3)',
    '#dash-quick-preview .summary-preview-compensation-row strong'
  ];
}

function resetMobileNavLabelAutoFitElement(element) {
  if (!element) return;
  element.style.fontSize = '';
  element.style.lineHeight = '';
  element.style.whiteSpace = '';
  element.style.maxWidth = '';
  element.style.overflow = '';
  element.style.textOverflow = '';
}

function applyMobileNavLabelAutoFit() {
  const navLabels = Array.from(document.querySelectorAll('.nav-links .nav-item .sidebar-text'));

  if (!isMobileCardLayoutActive()) {
    navLabels.forEach(resetMobileNavLabelAutoFitElement);
    return;
  }

  navLabels.forEach(label => {
    resetMobileNavLabelAutoFitElement(label);

    const navItem = label.closest('.nav-item');
    if (!navItem) return;

    const computedStyle = window.getComputedStyle(label);
    const navItemStyle = window.getComputedStyle(navItem);
    const baseFontSize = parseFloat(label.dataset.navAutofitBaseFontSize || computedStyle.fontSize);
    if (!Number.isFinite(baseFontSize) || baseFontSize <= 0) return;

    label.dataset.navAutofitBaseFontSize = `${baseFontSize}`;

    const baseLineHeightRaw = parseFloat(computedStyle.lineHeight);
    const baseLineHeight = Number.isFinite(baseLineHeightRaw) ? baseLineHeightRaw : (baseFontSize * 1.2);
    const horizontalPadding = (parseFloat(navItemStyle.paddingLeft) || 0) + (parseFloat(navItemStyle.paddingRight) || 0);
    const availableWidth = Math.max(18, navItem.clientWidth - horizontalPadding - 2);
    const minFontSize = Math.max(5.5, baseFontSize * 0.42);

    let currentFontSize = baseFontSize;
    label.style.whiteSpace = 'nowrap';
    label.style.overflow = 'visible';
    label.style.textOverflow = 'unset';
    label.style.maxWidth = `${availableWidth}px`;
    label.style.fontSize = `${currentFontSize}px`;
    label.style.lineHeight = `${baseLineHeight}px`;

    while (currentFontSize > minFontSize && label.scrollWidth > (availableWidth + 1)) {
      currentFontSize = Math.max(minFontSize, currentFontSize - 0.25);
      const scale = currentFontSize / baseFontSize;
      label.style.fontSize = `${currentFontSize}px`;
      label.style.lineHeight = `${Math.max(currentFontSize, baseLineHeight * scale)}px`;
    }

    if (label.scrollWidth > (availableWidth + 1) && currentFontSize > 0) {
      const proportionalFontSize = Math.max(minFontSize, Math.floor((currentFontSize * availableWidth / label.scrollWidth) * 100) / 100);
      const scale = proportionalFontSize / baseFontSize;
      label.style.fontSize = `${proportionalFontSize}px`;
      label.style.lineHeight = `${Math.max(proportionalFontSize, baseLineHeight * scale)}px`;
    }
  });
}

function getMobileCardAutoFitElement(target) {
  if (!target) return null;

  if (target.tagName === 'TD') {
    const existingWrapper = target.querySelector(':scope > .mobile-card-value');
    if (existingWrapper) return existingWrapper;

    const textNodes = Array.from(target.childNodes).filter(node => node.nodeType === Node.TEXT_NODE && node.textContent.trim() !== '');
    if (textNodes.length === 0) return null;

    const wrapper = document.createElement('span');
    wrapper.className = 'mobile-card-value';
    target.insertBefore(wrapper, target.firstChild || null);
    textNodes.forEach(node => wrapper.appendChild(node));
    return wrapper;
  }

  target.classList.add('mobile-card-value');
  return target;
}

function resetMobileCardAutoFitElement(element) {
  if (!element) return;
  element.style.display = '';
  element.style.fontSize = '';
  element.style.lineHeight = '';
  element.style.maxWidth = '';
  element.style.whiteSpace = '';
  element.style.overflowWrap = '';
  element.style.wordBreak = '';
}

function getAutoFitAvailableWidth(element) {
  if (!element) return 0;

  const getAvailableWidthInParentRow = (child, parent) => {
    if (!child || !parent) return 0;

    const parentWidth = parent.clientWidth || 0;
    if (!(parentWidth > 0)) return 0;

    const parentStyle = window.getComputedStyle(parent);
    const isFlexRow = (parentStyle.display === 'flex' || parentStyle.display === 'inline-flex')
      && !String(parentStyle.flexDirection || 'row').startsWith('column');

    if (!isFlexRow) return 0;

    const gapRaw = parseFloat(parentStyle.columnGap || parentStyle.gap);
    const gap = Number.isFinite(gapRaw) ? gapRaw : 0;
    const siblingsWidth = Array.from(parent.children)
      .filter(sibling => sibling !== child)
      .reduce((sum, sibling) => sum + sibling.getBoundingClientRect().width, 0);

    return Math.max(24, parentWidth - siblingsWidth - (gap * (parent.children.length - 1)) - 2);
  };

  let currentChild = element;
  let currentParent = element.parentElement;

  while (currentParent) {
    const availableWidth = getAvailableWidthInParentRow(currentChild, currentParent);
    if (availableWidth > 0) {
      return availableWidth;
    }

    currentChild = currentParent;
    currentParent = currentParent.parentElement;
  }

  return Math.max(element.clientWidth || 0, element.parentElement?.clientWidth || 0);
}

function autoFitMobileCardValue(element, minScale = 0.6) {
  if (!element) return;

  resetMobileCardAutoFitElement(element);

  const computedStyle = window.getComputedStyle(element);
  const baseFontSize = parseFloat(element.dataset.autofitBaseFontSize || computedStyle.fontSize);
  if (!Number.isFinite(baseFontSize) || baseFontSize <= 0) return;

  element.dataset.autofitBaseFontSize = `${baseFontSize}`;

  const baseLineHeightRaw = parseFloat(computedStyle.lineHeight);
  const baseLineHeight = Number.isFinite(baseLineHeightRaw) ? baseLineHeightRaw : (baseFontSize * 1.2);
  const minFontSize = baseFontSize * minScale;
  const availableWidth = getAutoFitAvailableWidth(element);

  if (!(availableWidth > 0)) return;

  let currentFontSize = baseFontSize;
  element.style.display = computedStyle.display === 'inline' ? 'inline-block' : computedStyle.display;
  element.style.maxWidth = `${availableWidth}px`;
  element.style.whiteSpace = 'nowrap';
  element.style.fontSize = `${currentFontSize}px`;
  element.style.lineHeight = `${baseLineHeight}px`;

  while (currentFontSize > minFontSize && element.scrollWidth > (availableWidth + 1)) {
    currentFontSize = Math.max(minFontSize, currentFontSize - 0.5);
    const scale = currentFontSize / baseFontSize;
    element.style.fontSize = `${currentFontSize}px`;
    element.style.lineHeight = `${Math.max(currentFontSize, baseLineHeight * scale)}px`;
  }

  if (element.scrollWidth > (availableWidth + 1)) {
    element.style.whiteSpace = 'normal';
    element.style.overflowWrap = 'anywhere';
    element.style.wordBreak = 'break-word';
  }
}

function applyMobileCardValueAutoFit() {
  const elements = getMobileCardAutoFitSelectorList()
    .flatMap(selector => Array.from(document.querySelectorAll(selector)))
    .map(getMobileCardAutoFitElement)
    .filter(Boolean);

  if (!isMobileCardLayoutActive()) {
    elements.forEach(resetMobileCardAutoFitElement);
    applyMobileNavLabelAutoFit();
    return;
  }

  elements.forEach(element => autoFitMobileCardValue(element, 0.6));
  applyMobileNavLabelAutoFit();
}

function scheduleMobileCardValueAutoFit() {
  if (mobileCardAutoFitFrame) {
    window.cancelAnimationFrame(mobileCardAutoFitFrame);
  }

  mobileCardAutoFitFrame = window.requestAnimationFrame(() => {
    mobileCardAutoFitFrame = null;
    applyMobileCardValueAutoFit();
  });
}

function hasCurrencyValueText(value) {
  return /\bzł\b/i.test((value || '').replace(/\s+/g, ' ').trim());
}

function getCurrencyValueAutoFitElement(target) {
  if (!target) return null;

  if (target.tagName === 'TD') {
    const existingWrapper = target.querySelector(':scope > .autofit-value--block, :scope > .mobile-card-value');
    if (existingWrapper) {
      existingWrapper.classList.add('autofit-value', 'autofit-value--block');
      return existingWrapper;
    }

    const textNodes = Array.from(target.childNodes).filter(node => node.nodeType === Node.TEXT_NODE && hasCurrencyValueText(node.textContent));
    if (textNodes.length === 0) return null;

    const wrapper = document.createElement('span');
    wrapper.className = 'autofit-value autofit-value--block';
    target.insertBefore(wrapper, target.firstChild || null);
    textNodes.forEach(node => wrapper.appendChild(node));
    return wrapper;
  }

  if (target.children.length > 0) return null;

  target.classList.add('autofit-value');
  return target;
}

function applyCurrencyValueAutoFit() {
  const mainContent = document.getElementById('main-content');
  if (!mainContent) return;

  const elements = Array.from(new Set(
    Array.from(mainContent.querySelectorAll('td, strong, span, div, p'))
      .filter(element => hasCurrencyValueText(element.textContent))
      .filter(element => !(element.closest('.view') && !element.closest('.view').classList.contains('active')))
      .filter(element => element.tagName === 'TD' || element.children.length === 0)
      .filter(element => element.getClientRects().length > 0)
      .map(getCurrencyValueAutoFitElement)
      .filter(Boolean)
  ));

  const minScale = isMobileCardLayoutActive() ? 0.55 : 0.72;
  elements.forEach(element => autoFitMobileCardValue(element, minScale));
}

function scheduleCurrencyValueAutoFit() {
  if (currencyValueAutoFitFrame) {
    window.cancelAnimationFrame(currencyValueAutoFitFrame);
  }

  currencyValueAutoFitFrame = window.requestAnimationFrame(() => {
    currencyValueAutoFitFrame = null;
    applyCurrencyValueAutoFit();
  });
}

function applyHoursSheetAutoFit() {
  const container = document.getElementById('sheet-detail-container');
  if (!container || container.style.display === 'none' || container.getClientRects().length === 0) return;

  const isPortraitMobile = window.matchMedia('(max-width: 550px)').matches;
  const headerMinScale = isPortraitMobile ? 0.48 : 0.58;
  const footerMinScale = isPortraitMobile ? 0.5 : 0.62;

  const headerLines = Array.from(container.querySelectorAll('.sheet-person-header-line'))
    .filter(element => element.getClientRects().length > 0);
  const footerValues = Array.from(container.querySelectorAll('.sheet-footer-earnings-fit'))
    .filter(element => element.getClientRects().length > 0);

  headerLines.forEach(element => {
    resetMobileCardAutoFitElement(element);
    autoFitMobileCardValue(element, headerMinScale);
  });

  footerValues.forEach(element => {
    resetMobileCardAutoFitElement(element);
    autoFitMobileCardValue(element, footerMinScale);
  });
}

function scheduleHoursSheetAutoFit() {
  if (hoursSheetAutoFitFrame) {
    window.cancelAnimationFrame(hoursSheetAutoFitFrame);
  }

  hoursSheetAutoFitFrame = window.requestAnimationFrame(() => {
    hoursSheetAutoFitFrame = null;
    applyHoursSheetAutoFit();
  });
}

function setHoursInstructionsExpanded(isExpanded = false) {
  const content = document.getElementById('hours-instructions-content');
  const toggle = document.getElementById('btn-toggle-hours-instructions');
  const icon = document.getElementById('hours-instructions-icon');
  if (!content || !toggle || !icon) return;

  content.style.display = isExpanded ? 'block' : 'none';
  toggle.dataset.expanded = isExpanded ? 'true' : 'false';
  toggle.setAttribute('aria-expanded', isExpanded ? 'true' : 'false');
  icon.style.transform = isExpanded ? 'rotate(180deg)' : 'rotate(0deg)';
}

function initHoursInstructionsControls() {
  const toggle = document.getElementById('btn-toggle-hours-instructions');
  if (!toggle || toggle.dataset.bound === 'true') return;

  toggle.dataset.bound = 'true';
  toggle.addEventListener('click', () => {
    setHoursInstructionsExpanded(toggle.dataset.expanded !== 'true');
  });

  setHoursInstructionsExpanded(false);
}

function formatPolishCurrencyNumber(value, minimumFractionDigits = 2, maximumFractionDigits = 2) {
  const amount = parseFloat(value);
  const safeAmount = Number.isFinite(amount) ? amount : 0;
  return safeAmount.toLocaleString('pl-PL', {
    minimumFractionDigits,
    maximumFractionDigits
  });
}

function formatPolishCurrencyWithSuffix(value, suffix = 'zł', minimumFractionDigits = 2, maximumFractionDigits = 2) {
  return `${formatPolishCurrencyNumber(value, minimumFractionDigits, maximumFractionDigits)}\u00A0${suffix}`;
}

function replaceCurrencySuffix(formattedValue, suffix = 'zł/h') {
  return (formattedValue || '').replace(/\u00A0zł$/, `\u00A0${suffix}`);
}

function normalizeCurrencySuffixSpacing(container = document.getElementById('main-content') || document.body) {
  if (!container) return;

  const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      const parent = node.parentElement;
      if (!parent) return NodeFilter.FILTER_REJECT;
      if (parent.closest('script, style, textarea, input, select, option')) return NodeFilter.FILTER_REJECT;
      if (!node.textContent || !node.textContent.includes('zł')) return NodeFilter.FILTER_REJECT;
      return NodeFilter.FILTER_ACCEPT;
    }
  });

  const textNodes = [];
  let currentNode = walker.nextNode();
  while (currentNode) {
    textNodes.push(currentNode);
    currentNode = walker.nextNode();
  }

  textNodes.forEach(node => {
    node.textContent = node.textContent.replace(/([+\-−]?)(\d[\d\s]*(?:[.,]\d+)?)(?:\s|\u00A0)*(zł(?:\/h|\/dzień)?)/g, (match, signRaw, rawNumber, suffix) => {
      const normalizedNumber = (rawNumber || '').replace(/\s+/g, '').replace(',', '.');
      const parsed = parseFloat(normalizedNumber);
      if (!Number.isFinite(parsed)) {
        return `${signRaw || ''}${rawNumber || ''}\u00A0${suffix}`;
      }

      const sign = signRaw === '+'
        ? '+'
        : (signRaw === '-' || signRaw === '−' ? '-' : '');
      const formattedNumber = formatPolishCurrencyNumber(parsed, 2, 2);
      return `${sign}${formattedNumber}\u00A0${suffix}`;
    });
  });
}

function renderAll() {
  const state = Store.getState();
  const isArchived = state.isArchived === true;
  const mainContent = document.getElementById('main-content');

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
  renderReports();
  renderPayouts();
  renderArchiveMonthManagement();
  renderHistoryManagement();
  updateHistoryButtonsState();
  normalizeCurrencySuffixSpacing(mainContent);
  lucide.createIcons();

  // Globalna blokada dla zarchiwizowanego miesiąca
  if (mainContent) {
    mainContent.classList.toggle('is-month-read-only', isArchived);
  }

  applyArchivedReadOnlyMode(isArchived);

  requestAnimationFrame(updateSidebarNavigationLayout);
  scheduleMobileCardValueAutoFit();
  scheduleCurrencyValueAutoFit();
  scheduleHoursSheetAutoFit();
}

function refreshOpenSheetDetailViews() {
  const monthlyDetailContainer = document.getElementById('sheet-detail-container');
  const worksDetailContainer = document.getElementById('works-detail-container');

  if (window.currentSheetId && monthlyDetailContainer?.style.display === 'block') {
    renderSheetDetail(window.currentSheetId);
  }

  if (window.currentWorksSheetId && worksDetailContainer?.style.display === 'block') {
    renderWorksSheetDetail(window.currentWorksSheetId);
  }
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

function normalizeMonthKeyValue(value) {
  return /^\d{4}-\d{2}$/.test((value || '').toString().trim()) ? value.toString().trim() : '';
}

function getArchiveUiSelectedMonthKeys() {
  return [...new Set([
    normalizeMonthKeyValue(getSelectedMonthKey()),
    normalizeMonthKeyValue(document.getElementById('global-month-select')?.value)
  ].filter(Boolean))];
}

function renderGlobalMonthSelector() {
  const monthInput = document.getElementById('global-month-select');
  const monthCompact = document.getElementById('global-month-select-compact');
  const archivedIndicator = document.getElementById('sidebar-month-archived-indicator');
  const state = Store.getState();
  const sidebar = document.getElementById('sidebar');
  const isCollapsed = sidebar?.classList.contains('collapsed') === true;
  const isMobile = window.matchMedia('(max-width: 550px)').matches;
  const showCompactMonth = isCollapsed && !window.matchMedia('(max-width: 550px)').matches;
  if (monthInput) {
    monthInput.value = getSelectedMonthKey();
    monthInput.dataset.lastValidMonth = monthInput.value;
  }

  if (monthCompact) {
    const selectedMonth = getSelectedMonthKey();
    const [year, month] = (selectedMonth || '').split('-');
    monthCompact.textContent = selectedMonth && month && year ? `${month}.${year.slice(-2)}` : selectedMonth;
    monthCompact.style.display = showCompactMonth ? 'inline-flex' : 'none';
    monthCompact.style.color = (state.isArchived === true || state.hasCommonSnapshot === true) ? 'var(--warning)' : '';
  }

  const summarySubtitle = document.getElementById('summary-subtitle');
  if (summarySubtitle) {
    summarySubtitle.textContent = `Podsumowanie ${formatMonthLabel(getSelectedMonthKey())}`;
  }

  const settlementSubtitle = document.getElementById('settlement-subtitle');
  if (settlementSubtitle) {
    settlementSubtitle.textContent = `Kalkulacja zarobku, kosztów i wypłat ${formatMonthLabel(getSelectedMonthKey())}`;
  }

  if (archivedIndicator) {
    if (!showCompactMonth && (state.isArchived === true || state.hasCommonSnapshot === true)) {
      archivedIndicator.textContent = state.isArchived === true ? 'Zarchiwizowany' : 'Edytowany z Archiwum';
      archivedIndicator.style.display = isMobile ? 'block' : 'block';
      archivedIndicator.style.color = 'var(--warning)';
    } else {
      archivedIndicator.style.display = 'none';
    }
  }

  requestAnimationFrame(updateSidebarNavigationLayout);
}

function setArchiveManagementExpanded(isExpanded) {
  const panel = document.getElementById('settings-archive-panel-content');
  const icon = document.getElementById('archive-management-panel-icon');
  const toggle = document.getElementById('btn-toggle-archive-management');
  if (!panel || !icon || !toggle) return;

  panel.style.display = isExpanded ? 'block' : 'none';
  toggle.dataset.expanded = isExpanded ? 'true' : 'false';
  icon.style.transform = isExpanded ? 'rotate(180deg)' : 'rotate(0deg)';
}

function getMonthArchiveStatusLabel(status) {
  if (status.isArchived) {
    return 'Zarchiwizowany • miesiąc tylko do odczytu';
  }
  if (status.hasSnapshot) {
    return 'Edytowalny • snapshot zachowany z czasu rozliczenia';
  }
  return status.hasData ? 'Edytowalny' : 'Brak danych miesiąca';
}

function renderArchiveMonthManagement() {
  const listEl = document.getElementById('settings-archive-months-list');
  if (!listEl) return;

  const selectedMonthKeys = getArchiveUiSelectedMonthKeys();

  const knownMonths = [...new Set([...selectedMonthKeys, ...Store.getKnownMonths()])]
    .filter(month => /^\d{4}-\d{2}$/.test(month))
    .sort((a, b) => b.localeCompare(a, 'pl-PL'));

  listEl.innerHTML = knownMonths.length > 0
    ? knownMonths.map(monthKey => {
        const status = Store.getMonthArchiveStatus(monthKey);
        const isSelectedMonth = selectedMonthKeys.includes(normalizeMonthKeyValue(monthKey));
        return `
          <div class="glass-panel" style="padding: 0.9rem 1rem; border-radius: 12px; display: flex; justify-content: space-between; align-items: center; gap: 1rem; flex-wrap: wrap;">
            <div>
              <div style="font-weight: 700; color: var(--text-primary);">${formatMonthLabel(monthKey)}${isSelectedMonth ? ' <span style="font-size: 0.78rem; color: var(--accent-primary); font-weight: 600;">(aktualnie wybrany)</span>' : ''}</div>
              <div style="font-size: 0.84rem; color: var(--text-secondary); margin-top: 0.2rem;">${getMonthArchiveStatusLabel(status)}</div>
            </div>
            <div style="display: flex; gap: 0.5rem; flex-wrap: wrap;">
              <button type="button" class="btn ${status.isArchived ? 'btn-secondary' : 'btn-danger'} btn-manage-archive-month" data-month="${monthKey}" data-action="toggle-archive">${status.isArchived ? 'Przywróć edycję' : 'Zarchiwizuj'}</button>
              ${!status.isArchived && status.hasSnapshot ? `<button type="button" class="btn btn-secondary btn-manage-archive-month" data-month="${monthKey}" data-action="delete-snapshot">Usuń snapshot</button>` : ''}
            </div>
          </div>
        `;
      }).join('')
    : '<p style="color: var(--text-muted);">Brak miesięcy do zarządzania.</p>';
}

function setHistoryManagementExpanded(isExpanded) {
  const panel = document.getElementById('settings-history-panel-content');
  const icon = document.getElementById('history-management-panel-icon');
  const toggle = document.getElementById('btn-toggle-history-management');
  if (!panel || !icon || !toggle) return;

  panel.style.display = isExpanded ? 'block' : 'none';
  toggle.dataset.expanded = isExpanded ? 'true' : 'false';
  icon.style.transform = isExpanded ? 'rotate(180deg)' : 'rotate(0deg)';
}

function formatHistoryTimestamp(value) {
  if (!value) return 'Brak daty';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleString('pl-PL', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit'
  });
}

function getHistoryMonthBadge(monthKey = '') {
  return monthKey ? ` • ${formatMonthLabel(monthKey)}` : '';
}

const SHARED_HISTORY_INDEX_PATH = 'shared_history/index';
const SHARED_HISTORY_ENTRIES_PATH = 'shared_history/entries';
const SHARED_DAILY_BACKUPS_INDEX_PATH = 'shared_backups_daily/index';
const SHARED_DAILY_BACKUPS_ENTRIES_PATH = 'shared_backups_daily/entries';
const SHARED_MONTHLY_BACKUPS_INDEX_PATH = 'shared_backups_monthly/index';
const SHARED_MONTHLY_BACKUPS_ENTRIES_PATH = 'shared_backups_monthly/entries';
const SHARED_HISTORY_REMOTE_LIMIT = 20;
const SHARED_DAILY_BACKUPS_REMOTE_LIMIT = 10;
const SHARED_MONTHLY_BACKUPS_REMOTE_LIMIT = 12;

const sharedHistoryEntryDetailsUiState = {};

function normalizeFirebaseEntryCollection(value = null) {
  if (Array.isArray(value)) return value;
  if (value && typeof value === 'object') return Object.values(value);
  return [];
}

function sortFirebaseIndexedEntriesByTimestamp(entries = []) {
  return [...normalizeFirebaseEntryCollection(entries)]
    .filter(entry => entry && typeof entry === 'object')
    .sort((a, b) => {
      const timestampDiff = (b?.timestamp || '').toString().localeCompare((a?.timestamp || '').toString(), 'pl-PL');
      if (timestampDiff !== 0) return timestampDiff;
      return (b?.id || '').toString().localeCompare((a?.id || '').toString(), 'pl-PL');
    });
}

function pruneIndexedFirebaseCollectionByLimit(indexValue = null, options = {}) {
  const {
    indexPath = '',
    entriesPath = '',
    limit = 0,
    label = 'kolekcji'
  } = options;
  const sortedEntries = sortFirebaseIndexedEntriesByTimestamp(indexValue);

  if (!(limit > 0) || sortedEntries.length <= limit || !indexPath || !entriesPath) {
    return Promise.resolve(sortedEntries);
  }

  const entriesToRemove = sortedEntries.slice(limit).filter(entry => entry?.id);
  if (entriesToRemove.length === 0) {
    return Promise.resolve(sortedEntries.slice(0, limit));
  }

  const updates = entriesToRemove.reduce((acc, entry) => {
    acc[`/${indexPath}/${entry.id}`] = null;
    acc[`/${entriesPath}/${entry.id}`] = null;
    return acc;
  }, {});

  return appUpdateFirebaseRootWithFallback(updates).then(() => {
    console.log(`firebase: Przycięto ${label} do ${limit} wpisów na podstawie index.`);
    return sortedEntries.slice(0, limit);
  }).catch(error => {
    console.error(`firebase: Nie udało się przyciąć ${label} na podstawie index:`, error);
    return sortedEntries;
  });
}

function isLegacyFirebaseEntryCollection(value = null) {
  if (Array.isArray(value)) return true;
  if (!value || typeof value !== 'object') return false;
  if (Object.prototype.hasOwnProperty.call(value, 'index') || Object.prototype.hasOwnProperty.call(value, 'entries')) return false;
  return Object.values(value).some(item => item && typeof item === 'object' && item.id);
}

function bootstrapIndexedFirebaseCollectionFromLegacy(legacyEntries = [], migrationBuilder = null, label = '') {
  if (!Array.isArray(legacyEntries) || legacyEntries.length === 0 || typeof migrationBuilder !== 'function') {
    return Promise.resolve();
  }

  const updates = migrationBuilder(legacyEntries) || {};
  if (Object.keys(updates).length === 0) return Promise.resolve();

  return appUpdateFirebaseRootWithFallback(updates).then(() => {
    console.log(`firebase: Zmigrowano legacy ${label} do gałęzi index/entries.`);
  }).catch(error => {
    console.error(`firebase: Nie udało się zmigrować legacy ${label} do gałęzi index/entries:`, error);
  });
}

function loadIndexedFirebaseCollectionWithLegacyFallback(indexPath = '', legacyPath = '', options = {}) {
  const {
    label = 'kolekcji',
    migrationBuilder = null,
    maxEntries = 0,
    entriesPath = ''
  } = options;

  if (typeof firebase === 'undefined' || !firebase.database) {
    return Promise.resolve([]);
  }

  return firebase.database().ref(indexPath).once('value').then(snapshot => {
    const indexValue = getTrackedFirebaseSnapshotValue(snapshot);
    const indexEntries = normalizeFirebaseEntryCollection(indexValue);
    if (indexEntries.length > 0) {
      return pruneIndexedFirebaseCollectionByLimit(indexValue, {
        indexPath,
        entriesPath,
        limit: maxEntries,
        label
      });
    }

    return firebase.database().ref(legacyPath).once('value').then(legacySnapshot => {
      const legacyValue = getTrackedFirebaseSnapshotValue(legacySnapshot);
      const legacyEntries = normalizeFirebaseEntryCollection(legacyValue);

      if (!isLegacyFirebaseEntryCollection(legacyValue) || legacyEntries.length === 0) {
        return indexValue || [];
      }

      bootstrapIndexedFirebaseCollectionFromLegacy(legacyEntries, migrationBuilder, label);
      return legacyValue;
    });
  }).catch(error => {
    console.error(`firebase: Nie udało się pobrać ${label}:`, error);
    return [];
  });
}

function fetchFirebaseEntryDetailWithLegacyFallback(entriesPath = '', legacyPath = '', entryId = '') {
  if (!entryId || typeof firebase === 'undefined' || !firebase.database) {
    return Promise.resolve(null);
  }

  return firebase.database().ref(`${entriesPath}/${entryId}`).once('value').then(snapshot => {
    const detailValue = getTrackedFirebaseSnapshotValue(snapshot);
    if (detailValue && typeof detailValue === 'object') {
      return detailValue;
    }

    return firebase.database().ref(legacyPath).once('value').then(legacySnapshot => {
      const legacyEntries = normalizeFirebaseEntryCollection(getTrackedFirebaseSnapshotValue(legacySnapshot));
      return legacyEntries.find(entry => entry?.id === entryId) || null;
    });
  }).catch(error => {
    console.error(`firebase: Nie udało się pobrać szczegółów wpisu ${entryId}:`, error);
    return null;
  });
}

function getSharedHistoryEntryDetailsState(entryId = '') {
  if (!sharedHistoryEntryDetailsUiState[entryId]) {
    sharedHistoryEntryDetailsUiState[entryId] = { expanded: false, loading: false, error: '' };
  }
  return sharedHistoryEntryDetailsUiState[entryId];
}

function pruneSharedHistoryEntryDetailsUiState(entries = []) {
  const validEntryIds = new Set((entries || []).map(entry => entry?.id).filter(Boolean));
  Object.keys(sharedHistoryEntryDetailsUiState).forEach(entryId => {
    if (!validEntryIds.has(entryId)) {
      delete sharedHistoryEntryDetailsUiState[entryId];
    }
  });
}

function ensureSharedHistoryEntryDetailLoaded(entryId = '') {
  const existingEntry = Store.getSharedHistoryEntry?.(entryId);
  if (existingEntry?.detailsLoaded === true) {
    return Promise.resolve(existingEntry);
  }

  return fetchFirebaseEntryDetailWithLegacyFallback(SHARED_HISTORY_ENTRIES_PATH, 'shared_history', entryId).then(entry => {
    if (!entry) return null;
    Store.cacheSharedHistoryEntryDetail?.(entry);
    return Store.getSharedHistoryEntry?.(entryId) || entry;
  });
}

function ensureBackupEntryDetailLoaded(type = 'daily', entryId = '') {
  const existingEntry = Store.getBackupEntry?.(type, entryId);
  if (existingEntry?.detailsLoaded === true) {
    return Promise.resolve(existingEntry);
  }

  const entriesPath = type === 'monthly' ? SHARED_MONTHLY_BACKUPS_ENTRIES_PATH : SHARED_DAILY_BACKUPS_ENTRIES_PATH;
  const legacyPath = type === 'monthly' ? 'shared_backups_monthly' : 'shared_backups_daily';

  return fetchFirebaseEntryDetailWithLegacyFallback(entriesPath, legacyPath, entryId).then(entry => {
    if (!entry) return null;
    Store.cacheBackupEntryDetail?.(type, entry);
    return Store.getBackupEntry?.(type, entryId) || entry;
  });
}

const HISTORY_SECRET_SNAPSHOT_UNLOCKS = {
  'daily-backups': { cycleCount: 0, pendingExpand: false, unlocked: false },
  'monthly-backups': { cycleCount: 0, pendingExpand: false, unlocked: false }
};

function getHistorySecretSnapshotUnlockState(sectionKey = '') {
  return HISTORY_SECRET_SNAPSHOT_UNLOCKS[sectionKey] || null;
}

function updateHistorySecretSnapshotActionVisibility() {
  const dailyActions = document.getElementById('history-daily-backups-actions');
  const monthlyActions = document.getElementById('history-monthly-backups-actions');

  if (dailyActions) {
    dailyActions.style.display = getHistorySecretSnapshotUnlockState('daily-backups')?.unlocked ? 'block' : 'none';
  }

  if (monthlyActions) {
    monthlyActions.style.display = getHistorySecretSnapshotUnlockState('monthly-backups')?.unlocked ? 'block' : 'none';
  }
}

function registerHistorySecretSnapshotSectionToggle(sectionKey = '', wasExpanded = false, isExpanded = false) {
  const state = getHistorySecretSnapshotUnlockState(sectionKey);
  if (!state || state.unlocked) {
    updateHistorySecretSnapshotActionVisibility();
    return;
  }

  if (wasExpanded && !isExpanded) {
    state.pendingExpand = true;
  } else if (!wasExpanded && isExpanded && state.pendingExpand) {
    state.pendingExpand = false;
    state.cycleCount += 1;
    if (state.cycleCount >= 5) {
      state.unlocked = true;
    }
  }

  updateHistorySecretSnapshotActionVisibility();
}

function escapeHistoryPreviewText(value = '') {
  return (value ?? '').toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getHistoryCollectionDiffRecords(beforeSnapshot, afterSnapshot) {
  const beforeItems = Array.isArray(beforeSnapshot) ? beforeSnapshot.filter(item => item && typeof item === 'object') : [];
  const afterItems = Array.isArray(afterSnapshot) ? afterSnapshot.filter(item => item && typeof item === 'object') : [];
  const beforeById = new Map(beforeItems.filter(item => item.id).map(item => [item.id, item]));
  const afterById = new Map(afterItems.filter(item => item.id).map(item => [item.id, item]));

  return {
    beforeItems,
    afterItems,
    addedRecords: afterItems
      .filter(item => item.id && !beforeById.has(item.id))
      .map(afterItem => ({ beforeItem: null, afterItem })),
    removedRecords: beforeItems
      .filter(item => item.id && !afterById.has(item.id))
      .map(beforeItem => ({ beforeItem, afterItem: null })),
    updatedRecords: afterItems
      .filter(item => item.id && beforeById.has(item.id) && !deepEqual(beforeById.get(item.id), item))
      .map(afterItem => ({ beforeItem: beforeById.get(afterItem.id) || null, afterItem }))
  };
}

function getHistoryChangedCollectionItems(beforeSnapshot, afterSnapshot) {
  const diff = getHistoryCollectionDiffRecords(beforeSnapshot, afterSnapshot);

  if (diff.addedRecords.length > 0) {
    return { beforeItem: null, afterItem: diff.addedRecords[0].afterItem, changeType: 'added' };
  }

  if (diff.removedRecords.length > 0) {
    return { beforeItem: diff.removedRecords[0].beforeItem, afterItem: null, changeType: 'removed' };
  }

  if (diff.updatedRecords.length > 0) {
    return {
      beforeItem: diff.updatedRecords[0].beforeItem,
      afterItem: diff.updatedRecords[0].afterItem,
      changeType: 'updated'
    };
  }

  const fallbackBefore = diff.beforeItems[0] || null;
  const fallbackAfter = diff.afterItems[0] || null;

  return {
    beforeItem: fallbackBefore,
    afterItem: fallbackAfter,
    changeType: fallbackBefore && fallbackAfter
      ? (deepEqual(fallbackBefore, fallbackAfter) ? 'unchanged' : 'updated')
      : (fallbackAfter ? 'added' : (fallbackBefore ? 'removed' : 'unchanged'))
  };
}

function getHistoryChangedCollectionItem(beforeSnapshot, afterSnapshot) {
  const { afterItem, beforeItem } = getHistoryChangedCollectionItems(beforeSnapshot, afterSnapshot);
  return afterItem || beforeItem || null;
}

function getHistoryPreviewState(monthKey = getSelectedMonthKey()) {
  const normalizedMonthKey = monthKey || getSelectedMonthKey();
  return Store.getStateForMonth ? Store.getStateForMonth(normalizedMonthKey) : Store.getState();
}

function getHistoryMonthSettings(monthKey = getSelectedMonthKey()) {
  return Store.getMonthSettings ? Store.getMonthSettings(monthKey || getSelectedMonthKey()) : {};
}

function getHistoryPersonTypeLabel(type = '') {
  if (type === 'PARTNER') return 'Wspólnik';
  if (type === 'SEPARATE_COMPANY') return 'Osobna Firma';
  if (type === 'WORKING_PARTNER') return 'Wspólnik Pracujący';
  return 'Pracownik';
}

function getHistoryClientActiveLabel(client, monthKey) {
  const monthSettings = getHistoryMonthSettings(monthKey);
  if (monthSettings.clients?.[client.id] !== undefined) {
    return monthSettings.clients[client.id] !== false ? 'Aktywny' : 'Nieaktywny';
  }
  return client?.isActive !== false ? 'Aktywny' : 'Nieaktywny';
}

function getHistoryPersonActiveLabel(person, monthKey) {
  const monthSettings = getHistoryMonthSettings(monthKey);
  if (monthSettings.persons?.[person.id] !== undefined) {
    return monthSettings.persons[person.id] !== false ? 'Aktywny' : 'Nieaktywny';
  }
  return person?.isActive !== false ? 'Aktywny' : 'Nieaktywny';
}

function buildPersonHistoryPreviewFromRecord(person, monthKey = getSelectedMonthKey()) {
  if (!person) return '';

  const state = getHistoryPreviewState(monthKey);
  const employerName = (person.type === 'EMPLOYEE' || person.type === 'WORKING_PARTNER')
    ? (getEmployerNameById(state, person.employerId) || 'Nie przypisano')
    : '-';
  const costLabel = (person.type === 'PARTNER' || person.type === 'WORKING_PARTNER' || person.type === 'SEPARATE_COMPANY')
    ? (doesPersonParticipateInCosts(person) ? 'Tak' : 'Nie')
    : '-';
  const rateLabel = person.type === 'EMPLOYEE'
    ? (person.hourlyRate ? `${parseFloat(person.hourlyRate).toFixed(2)} zł` : '-')
    : (person.hourlyRate ? `${parseFloat(person.hourlyRate).toFixed(2)} zł` : 'Stawka Klienta');

  return `[${getPersonDisplayName(person) || '-'}\t${getHistoryPersonTypeLabel(person.type)}\t${rateLabel}\t${employerName}\t${costLabel}\t${getHistoryPersonActiveLabel(person, monthKey)}]`;
}

function buildPersonHistoryPreview(change = {}) {
  const { afterItem, beforeItem } = getHistoryChangedCollectionItems(change.beforeSnapshot, change.afterSnapshot);
  return buildPersonHistoryPreviewFromRecord(afterItem || beforeItem, change.month || getSelectedMonthKey());
}

function buildClientHistoryPreviewFromRecord(client, monthKey = getSelectedMonthKey()) {
  if (!client) return '';

  const rateLabel = Number.isFinite(parseFloat(client.hourlyRate)) && parseFloat(client.hourlyRate) > 0
    ? `${parseFloat(client.hourlyRate).toFixed(2)} zł/h`
    : 'Brak stawki';

  return `[${client.name || '-'}\t${rateLabel}\t${getHistoryClientActiveLabel(client, monthKey)}]`;
}

function buildClientHistoryPreview(change = {}) {
  const { afterItem, beforeItem } = getHistoryChangedCollectionItems(change.beforeSnapshot, change.afterSnapshot);
  return buildClientHistoryPreviewFromRecord(afterItem || beforeItem, change.month || getSelectedMonthKey());
}

function buildExpenseHistoryPreviewFromRecord(expense, monthKey = getSelectedMonthKey()) {
  if (!expense) return '';

  const state = Store.getStateForMonth ? Store.getStateForMonth(monthKey) : Store.getState();
  const typeMeta = expense.type === 'BONUS'
    ? { label: 'Premia' }
    : expense.type === 'DIETA'
      ? { label: 'Dieta' }
      : expense.type === 'REFUND'
        ? { label: 'Zwrot' }
      : expense.type === 'ADVANCE'
        ? { label: 'Zaliczka' }
        : { label: 'Koszt' };
  const title = expense.type === 'ADVANCE'
    ? 'Zaliczka'
    : (expense.type === 'BONUS'
        ? 'Premia'
        : (expense.type === 'DIETA'
            ? 'Dieta'
            : (expense.type === 'REFUND' ? `Zwrot za ${expense.name || 'Zwrot'}` : (expense.name || 'Koszt'))));
  const payerName = (() => {
    if (!expense.paidById) return 'Nieznany';
    if (expense.paidById === BONUS_EXPENSE_PAYER_ID) return 'Od Wszystkich Wspólników';
    if (expense.type === 'DIETA') return getSettlementDietaMeta(expense, state, monthKey);
    if (expense.paidById.startsWith('client_')) {
      const clientId = expense.paidById.slice(7);
      return `${state.clients.find(client => client.id === clientId)?.name || 'Nieznany'} (Klient)`;
    }
    return getPersonDisplayName(state.persons.find(person => person.id === expense.paidById)) || 'Nieznany';
  })();
  const recipientName = (expense.type === 'ADVANCE' || expense.type === 'BONUS' || expense.type === 'DIETA' || expense.type === 'REFUND')
    ? getExpenseRecipientName(expense, state)
    : '-';
  const resolvedAmount = Calculations.getExpenseEffectiveAmount(expense, state, monthKey);

  return `[${expense.date || '-'}\t${title}\t${typeMeta.label}\t${resolvedAmount.toFixed(2)} zł\t${payerName}\t${recipientName}]`;
}

function buildExpenseHistoryPreview(change = {}) {
  const { afterItem, beforeItem } = getHistoryChangedCollectionItems(change.beforeSnapshot, change.afterSnapshot);
  return buildExpenseHistoryPreviewFromRecord(afterItem || beforeItem, change.month || getSelectedMonthKey());
}

function buildMonthlySheetHistoryPreviewFromRecord(sheet, monthKey = getSelectedMonthKey()) {
  if (!sheet) return '';

  const state = getHistoryPreviewState(monthKey);
  const visiblePersonIds = new Set(getVisiblePersonsForSheet(state, sheet).map(person => person.id));
  const totalHours = Calculations.getSheetTotalHours(sheet, visiblePersonIds);
  const revenue = totalHours * Calculations.getSheetClientRate(sheet, state);

  return `[${sheet.client || '-'}\t${sheet.site || '-'}\t${totalHours.toFixed(1)}h\t${revenue.toFixed(2)} zł]`;
}

function buildMonthlySheetHistoryPreview(change = {}) {
  const { afterItem, beforeItem } = getHistoryChangedCollectionItems(change.beforeSnapshot, change.afterSnapshot);
  return buildMonthlySheetHistoryPreviewFromRecord(afterItem || beforeItem, change.month || afterItem?.month || beforeItem?.month || getSelectedMonthKey());
}

function buildWorksSheetHistoryPreviewFromRecord(sheet) {
  if (!sheet) return '';

  const entries = Array.isArray(sheet.entries) ? sheet.entries : [];
  const totalItems = entries.length;
  const totalValue = entries.reduce((sum, entry) => sum + ((parseFloat(entry.quantity) || 0) * (parseFloat(entry.price) || 0)), 0);

  return `[${sheet.client || '-'}\t${sheet.site || '-'}\t${totalItems} pozycji\t${totalValue.toFixed(2)} zł]`;
}

function buildWorksSheetHistoryPreview(change = {}) {
  const { afterItem, beforeItem } = getHistoryChangedCollectionItems(change.beforeSnapshot, change.afterSnapshot);
  return buildWorksSheetHistoryPreviewFromRecord(afterItem || beforeItem);
}

function buildInvoicesHistoryPreviewFromRecord(invoiceRecord, monthKey = getSelectedMonthKey(), snapshot = null) {
  const state = getHistoryPreviewState(monthKey);

  if (invoiceRecord) {
    const clientName = invoiceRecord.clientId
      ? (state.clients.find(client => client.id === invoiceRecord.clientId)?.name || invoiceRecord.clientName || 'Nieznany klient')
      : (invoiceRecord.clientName || 'Nieznany klient');
    const issuerName = getPersonDisplayName(state.persons.find(person => person.id === invoiceRecord.issuerId)) || 'Nieznany wystawca';
    return `[${clientName}\t${issuerName}\t${(parseFloat(invoiceRecord.amount) || 0).toFixed(2)} zł]`;
  }

  const invoices = snapshot && Object.keys(snapshot).length > 0 ? snapshot : {};
  const clientsCount = Object.keys(invoices.clients || {}).length;
  const extraInvoicesCount = Array.isArray(invoices.extraInvoices) ? invoices.extraInvoices.length : 0;
  return `[Data wystawienia: ${invoices.issueDate || '-'}\tKonfiguracje klientów: ${clientsCount}\tDodatkowe faktury: ${extraInvoicesCount}]`;
}

function buildInvoicesHistoryPreview(change = {}) {
  const beforeInvoices = change.beforeSnapshot || {};
  const afterInvoices = change.afterSnapshot || {};
  const { afterItem, beforeItem } = getHistoryChangedCollectionItems(beforeInvoices.extraInvoices, afterInvoices.extraInvoices);
  const monthKey = change.month || getSelectedMonthKey();
  return buildInvoicesHistoryPreviewFromRecord(afterItem || beforeItem, monthKey, Object.keys(afterInvoices || {}).length > 0 ? afterInvoices : beforeInvoices);
}

function buildHistoryPreviewChangeLines(beforePreview = '', afterPreview = '', changeType = 'updated') {
  const normalizedBefore = (beforePreview || '').trim();
  const normalizedAfter = (afterPreview || '').trim();

  if (changeType === 'added' && normalizedAfter) {
    return [`dodano nowe: ${normalizedAfter}`];
  }

  if (changeType === 'removed' && normalizedBefore) {
    return [`usunięto: ${normalizedBefore}`];
  }

  const lines = [];
  if (normalizedBefore) lines.push(`przed: ${normalizedBefore}`);
  if (normalizedAfter) lines.push(`po: ${normalizedAfter}`);
  return lines;
}

function buildHistoryCollectionPreviewLines(diff = {}, previewBuilder = () => '') {
  const lines = [];

  (diff.removedRecords || []).forEach(record => {
    const preview = previewBuilder(record.beforeItem || null, 'before');
    if (preview) lines.push(`usunięto: ${preview}`);
  });

  (diff.addedRecords || []).forEach(record => {
    const preview = previewBuilder(record.afterItem || null, 'after');
    if (preview) lines.push(`dodano nowe: ${preview}`);
  });

  (diff.updatedRecords || []).forEach(record => {
    const beforePreview = previewBuilder(record.beforeItem || null, 'before');
    const afterPreview = previewBuilder(record.afterItem || null, 'after');
    if (beforePreview) lines.push(`przed: ${beforePreview}`);
    if (afterPreview) lines.push(`po: ${afterPreview}`);
  });

  if (lines.length === 0 && ((diff.beforeItems || []).length > 0 || (diff.afterItems || []).length > 0)) {
    const fallbackBefore = previewBuilder(diff.beforeItems?.[0] || null, 'before');
    const fallbackAfter = previewBuilder(diff.afterItems?.[0] || null, 'after');
    return buildHistoryPreviewChangeLines(fallbackBefore, fallbackAfter, fallbackBefore && fallbackAfter ? 'updated' : (fallbackAfter ? 'added' : 'removed'));
  }

  return lines;
}

function getHistoryScopeCollectionDiff(change = {}) {
  if (change.scope === 'month.monthSettings.invoices') {
    return getHistoryCollectionDiffRecords(change.beforeSnapshot?.extraInvoices, change.afterSnapshot?.extraInvoices);
  }

  return getHistoryCollectionDiffRecords(change.beforeSnapshot, change.afterSnapshot);
}

function historyScopeHasRemovedItems(change = {}) {
  if (change?.hasRemovedItems === true) return true;
  if (change?.hasRemovedItems === false && !change?.detailsLoaded && !change?.beforeSnapshot && !change?.afterSnapshot) return false;
  return getHistoryScopeCollectionDiff(change).removedRecords.length > 0;
}

function historyScopeHasAddedItems(change = {}) {
  if (change?.hasAddedItems === true) return true;
  if (change?.hasAddedItems === false && !change?.detailsLoaded && !change?.beforeSnapshot && !change?.afterSnapshot) return false;
  return getHistoryScopeCollectionDiff(change).addedRecords.length > 0;
}

function getHistoryEntryAdditionalActionButtons(entry = {}) {
  const changes = Array.isArray(entry.changes) ? entry.changes : [];
  const hasRemovedItems = changes.some(historyScopeHasRemovedItems);
  const hasAddedItems = changes.some(historyScopeHasAddedItems);
  const buttons = [];

  if (hasRemovedItems) {
    buttons.push({ type: 'shared-history-merge-before', label: 'Przywróć i połącz sprzed zmiany' });
  }

  if (hasAddedItems) {
    buttons.push({ type: 'shared-history-remove-added', label: 'Usuń wpisy dodane przez zmianę' });
  }

  return buttons;
}

function buildHistoryChangePreviewLines(change = {}) {
  if (Array.isArray(change.previewLines)) {
    return change.previewLines.filter(Boolean);
  }

  const monthKey = change.month || getSelectedMonthKey();

  if (change.scope === 'common.persons' || change.scope === 'month.monthSettings.commonSnapshot.persons') {
    return buildHistoryCollectionPreviewLines(
      getHistoryCollectionDiffRecords(change.beforeSnapshot, change.afterSnapshot),
      item => buildPersonHistoryPreviewFromRecord(item, monthKey)
    );
  }

  if (change.scope === 'common.clients' || change.scope === 'month.monthSettings.commonSnapshot.clients') {
    return buildHistoryCollectionPreviewLines(
      getHistoryCollectionDiffRecords(change.beforeSnapshot, change.afterSnapshot),
      item => buildClientHistoryPreviewFromRecord(item, monthKey)
    );
  }

  if (change.scope === 'month.monthlySheets') {
    return buildHistoryCollectionPreviewLines(
      getHistoryCollectionDiffRecords(change.beforeSnapshot, change.afterSnapshot),
      item => buildMonthlySheetHistoryPreviewFromRecord(item, monthKey)
    );
  }

  if (change.scope === 'month.worksSheets') {
    return buildHistoryCollectionPreviewLines(
      getHistoryCollectionDiffRecords(change.beforeSnapshot, change.afterSnapshot),
      item => buildWorksSheetHistoryPreviewFromRecord(item)
    );
  }

  if (change.scope === 'month.expenses') {
    return buildHistoryCollectionPreviewLines(
      getHistoryCollectionDiffRecords(change.beforeSnapshot, change.afterSnapshot),
      item => buildExpenseHistoryPreviewFromRecord(item, monthKey)
    );
  }

  if (change.scope === 'month.monthSettings.invoices') {
    const beforeInvoices = change.beforeSnapshot || {};
    const afterInvoices = change.afterSnapshot || {};
    const invoiceDiff = getHistoryCollectionDiffRecords(beforeInvoices.extraInvoices, afterInvoices.extraInvoices);

    if (invoiceDiff.addedRecords.length > 0 || invoiceDiff.removedRecords.length > 0 || invoiceDiff.updatedRecords.length > 0) {
      return buildHistoryCollectionPreviewLines(
        invoiceDiff,
        (item) => buildInvoicesHistoryPreviewFromRecord(item, monthKey)
      );
    }

    return buildHistoryPreviewChangeLines(
      buildInvoicesHistoryPreviewFromRecord(null, monthKey, beforeInvoices),
      buildInvoicesHistoryPreviewFromRecord(null, monthKey, afterInvoices),
      (Object.keys(beforeInvoices || {}).length === 0 && Object.keys(afterInvoices || {}).length > 0) ? 'added' : 'updated'
    );
  }

  const previewText = buildHistoryChangePreview(change);
  return previewText ? [previewText] : [];
}

function buildHistoryChangePreview(change = {}) {
  if (change.scope === 'common.persons' || change.scope === 'month.monthSettings.commonSnapshot.persons') {
    return buildPersonHistoryPreview(change);
  }
  if (change.scope === 'common.clients' || change.scope === 'month.monthSettings.commonSnapshot.clients') {
    return buildClientHistoryPreview(change);
  }
  if (change.scope === 'month.monthlySheets') {
    return buildMonthlySheetHistoryPreview(change);
  }
  if (change.scope === 'month.worksSheets') {
    return buildWorksSheetHistoryPreview(change);
  }
  if (change.scope === 'month.expenses') {
    return buildExpenseHistoryPreview(change);
  }
  if (change.scope === 'month.monthSettings.invoices') {
    return buildInvoicesHistoryPreview(change);
  }
  return '';
}

function setHistorySectionExpanded(sectionKey, isExpanded = true) {
  const content = document.getElementById(`history-${sectionKey}-content`);
  const icon = document.getElementById(`history-${sectionKey}-icon`);
  const toggle = document.querySelector(`.history-section-toggle[data-history-group="${sectionKey}"]`);
  if (!content || !icon || !toggle) return;

  content.style.display = isExpanded ? 'block' : 'none';
  toggle.dataset.expanded = isExpanded ? 'true' : 'false';
  icon.style.transform = isExpanded ? 'rotate(180deg)' : 'rotate(0deg)';
  updateHistorySecretSnapshotActionVisibility();
}

function buildHistoryEntryScopeListHtml(scopeItems = [], options = {}) {
  const {
    includePreviewLines = true
  } = options;

  if (!Array.isArray(scopeItems) || scopeItems.length === 0) return '';

  return `<ul class="history-entry-scope-list">${scopeItems.map(change => {
    const previewLines = includePreviewLines ? buildHistoryChangePreviewLines(change) : [];
    return `<li><div>${change.scopeLabel || change.scope}${change.month ? ` (${formatMonthLabel(change.month)})` : ''}</div>${previewLines.length > 0 ? `<div style="margin-top: 0.35rem; color: var(--text-secondary); font-size: 0.82rem; font-family: Consolas, 'Courier New', monospace; white-space: pre-wrap; word-break: break-word;">${previewLines.map(line => `<div>${escapeHistoryPreviewText(line)}</div>`).join('')}</div>` : ''}</li>`;
  }).join('')}</ul>`;
}

function renderHistoryEntryList(containerId, entries = [], options = {}) {
  const container = document.getElementById(containerId);
  if (!container) return;

  if (!Array.isArray(entries) || entries.length === 0) {
    container.innerHTML = `<p class="history-empty-state">${options.emptyText || 'Brak wpisów.'}</p>`;
    return;
  }

  container.innerHTML = entries.map((entry, index) => {
    const scopeItems = Array.isArray(entry.changes) ? entry.changes : [];
    const isSharedHistory = options.entryKind === 'shared-history';
    const detailsState = isSharedHistory ? getSharedHistoryEntryDetailsState(entry.id) : null;
    const detailsExpanded = detailsState?.expanded === true;
    const detailsLoaded = entry.detailsLoaded === true;
    const configuredActionButtons = typeof options.actionButtons === 'function'
      ? options.actionButtons(entry, index, entries)
      : options.actionButtons;
    const actionButtons = Array.isArray(configuredActionButtons) && configuredActionButtons.length > 0
      ? configuredActionButtons
      : (options.restoreAction ? [{ type: options.restoreAction, label: options.restoreLabel || 'Przywróć' }] : []);
    const detailsToggleButton = isSharedHistory
      ? `<button type="button" class="btn btn-secondary btn-toggle-history-entry-details" data-entry-id="${entry.id}">${detailsExpanded ? 'Ukryj Szczegóły' : 'Szczegóły'}</button>`
      : '';
    const actionButton = (detailsToggleButton || actionButtons.length > 0)
      ? `<div class="history-entry-actions">${detailsToggleButton}${actionButtons.map(action => `<button type="button" class="btn btn-secondary btn-restore-history-entry" data-entry-id="${entry.id}" data-restore-type="${action.type}">${action.label}</button>`).join('')}</div>`
      : '';
    const scopeListHtml = isSharedHistory
      ? buildHistoryEntryScopeListHtml(scopeItems, { includePreviewLines: false })
      : buildHistoryEntryScopeListHtml(scopeItems, { includePreviewLines: true });
    const detailsHtml = isSharedHistory && detailsExpanded
      ? `<div class="history-entry-details-panel">${detailsState?.loading ? '<p class="history-entry-details-status">Ładowanie szczegółów...</p>' : (detailsState?.error ? `<p class="history-entry-details-status history-entry-details-status--error">${escapeHistoryPreviewText(detailsState.error)}</p>` : (detailsLoaded ? buildHistoryEntryScopeListHtml(scopeItems, { includePreviewLines: true }) : '<p class="history-entry-details-status">Brak szczegółów do wyświetlenia.</p>'))}</div>`
      : '';

    return `
      <div class="history-entry-card">
        <div class="history-entry-header">
          <div>
            <div style="font-weight: 700; color: var(--text-primary);">${entry.label || 'Zmiana'}${getHistoryMonthBadge(entry.month)}</div>
            <div class="history-entry-meta">${formatHistoryTimestamp(entry.timestamp)} • ${entry.author || 'Nieznany użytkownik'}</div>
          </div>
        </div>
        ${scopeListHtml}
        ${detailsHtml}
        ${actionButton}
      </div>
    `;
  }).join('');
}

function renderHistoryManagement() {
  const summary = document.getElementById('history-summary');
  const localHistoryTitle = document.getElementById('history-local-history-title');
  const sharedHistoryTitle = document.getElementById('history-shared-history-title');
  const dailyBackupsTitle = document.getElementById('history-daily-backups-title');
  const monthlyBackupsTitle = document.getElementById('history-monthly-backups-title');
  const historyState = Store.getHistoryState ? Store.getHistoryState() : null;
  if (!historyState) return;

  pruneSharedHistoryEntryDetailsUiState(historyState.sharedEntries || []);

  if (summary) {
    summary.innerHTML = `
      <div class="history-summary-card">
        <span style="color: var(--text-secondary); font-size: 0.82rem;">Undo lokalnie</span>
        <strong>${historyState.undoCount || 0} / 50</strong>
      </div>
      <div class="history-summary-card">
        <span style="color: var(--text-secondary); font-size: 0.82rem;">Redo lokalnie</span>
        <strong>${historyState.redoCount || 0} / 50</strong>
      </div>
    `;
  }

  if (localHistoryTitle) {
    localHistoryTitle.textContent = `Historia lokalna sesji (${historyState.undoCount || 0} / 50)`;
  }
  if (sharedHistoryTitle) {
    sharedHistoryTitle.textContent = `Wspólna historia zmian (${(historyState.sharedEntries || []).length} / 20)`;
  }
  if (dailyBackupsTitle) {
    dailyBackupsTitle.textContent = `Snapshoty dzienne (${(historyState.dailyBackups || []).length} / 10)`;
  }
  if (monthlyBackupsTitle) {
    monthlyBackupsTitle.textContent = `Snapshoty miesięczne (${(historyState.monthlyBackups || []).length} / 12)`;
  }

  renderHistoryEntryList('history-local-history-list', historyState.localUndoEntries || [], {
    emptyText: 'Brak lokalnych zmian w tej sesji.',
    actionButtons: (entry, index) => index === 0
      ? [{ type: 'local-history-undo', label: 'Cofnij' }]
      : []
  });
  renderHistoryEntryList('history-shared-history-list', historyState.sharedEntries || [], {
    entryKind: 'shared-history',
    emptyText: 'Brak wspólnej historii zmian.',
    actionButtons: (entry) => ([
      { type: 'shared-history-before', label: 'Przywróć stan sprzed zmiany' },
      { type: 'shared-history-after', label: 'Przywróć stan po zmianie' },
      ...getHistoryEntryAdditionalActionButtons(entry)
    ])
  });
  renderHistoryEntryList('history-daily-backups-list', historyState.dailyBackups || [], {
    emptyText: 'Brak snapshotów dziennych.',
    restoreAction: 'daily-backup',
    restoreLabel: 'Przywróć snapshot dzienny'
  });
  renderHistoryEntryList('history-monthly-backups-list', historyState.monthlyBackups || [], {
    emptyText: 'Brak snapshotów miesięcznych.',
    restoreAction: 'monthly-backup',
    restoreLabel: 'Przywróć snapshot miesięczny'
  });
}

function updateHistoryButtonsState() {
  const undoBtn = document.getElementById('btn-history-undo');
  const redoBtn = document.getElementById('btn-history-redo');
  const historyState = Store.getHistoryState ? Store.getHistoryState() : null;
  if (!historyState) return;

  if (undoBtn) {
    undoBtn.disabled = historyState.canUndo !== true;
    undoBtn.title = historyState.canUndo ? 'Cofnij ostatnią zmianę' : 'Brak zmian do cofnięcia';
  }
  if (redoBtn) {
    redoBtn.disabled = historyState.canRedo !== true;
    redoBtn.title = historyState.canRedo ? 'Ponów ostatnią zmianę' : 'Brak zmian do ponowienia';
  }
}

function initHistoryControls() {
  const undoBtn = document.getElementById('btn-history-undo');
  const redoBtn = document.getElementById('btn-history-redo');
  const openHistoryBtn = document.getElementById('btn-open-history-view');
  const backToSettingsBtn = document.getElementById('btn-back-to-settings-from-history');

  if (undoBtn) {
    undoBtn.addEventListener('click', () => {
      const result = Store.undo ? Store.undo() : { success: false, message: 'Undo jest niedostępne.' };
      if (!result?.success && result?.message) alert(result.message);
    });
  }

  if (redoBtn) {
    redoBtn.addEventListener('click', () => {
      const result = Store.redo ? Store.redo() : { success: false, message: 'Redo jest niedostępne.' };
      if (!result?.success && result?.message) alert(result.message);
    });
  }

  if (openHistoryBtn) {
    openHistoryBtn.addEventListener('click', () => {
      activateNavigationTarget('history-view');
    });
  }

  if (backToSettingsBtn) {
    backToSettingsBtn.addEventListener('click', () => {
      activateNavigationTarget('settings-view');
    });
  }

  document.querySelectorAll('.history-section-toggle').forEach(toggle => {
    toggle.addEventListener('click', () => {
      const sectionKey = toggle.dataset.historyGroup;
      if (!sectionKey) return;
      const wasExpanded = toggle.dataset.expanded === 'true';
      const isExpanded = !wasExpanded;
      setHistorySectionExpanded(sectionKey, isExpanded);
      registerHistorySecretSnapshotSectionToggle(sectionKey, wasExpanded, isExpanded);
    });
  });

  window.addEventListener('historyStateChanged', () => {
    updateHistoryButtonsState();
    renderHistoryManagement();
  });

  setHistorySectionExpanded('local-history', true);
  setHistorySectionExpanded('shared-history', false);
  setHistorySectionExpanded('daily-backups', false);
  setHistorySectionExpanded('monthly-backups', false);

  updateHistorySecretSnapshotActionVisibility();
  updateHistoryButtonsState();
}

function manageMonthArchive(action, monthKey = getSelectedMonthKey()) {
  const month = monthKey || getSelectedMonthKey();
  const status = Store.getMonthArchiveStatus(month);

  if (action === 'archive') {
    if (status.isArchived) return;
    if (!confirm(`Czy chcesz zarchiwizować miesiąc ${formatMonthLabel(month)}? Zostanie zablokowany do odczytu i zapisany zostanie snapshot danych wspólnych.`)) return;
    Store.setMonthArchived(true, month);
    return;
  }

  if (action === 'unarchive') {
    if (!status.isArchived) return;
    if (!confirm(`Czy chcesz przywrócić możliwość edycji dla miesiąca ${formatMonthLabel(month)}? Snapshot danych zostanie zachowany.`)) return;
    Store.setMonthArchived(false, month);
    return;
  }

  if (action === 'delete-snapshot') {
    if (!status.hasSnapshot) return;
    if (!confirm(`Czy chcesz usunąć snapshot danych wspólnych dla miesiąca ${formatMonthLabel(month)}? Tej operacji nie da się cofnąć.`)) return;
    Store.deleteMonthCommonSnapshot(month);
  }
}

function initGlobalMonthSelector() {
  const monthInput = document.getElementById('global-month-select');
  const monthCompact = document.getElementById('global-month-select-compact');
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
  if (monthCompact) {
    monthCompact.addEventListener('click', openMonthPicker);
  }
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
    window.currentWorksSheetId = null;
    Store.setSelectedMonth(nextMonth);
  });
}

function updateMonthPickerRestrictions(input) {
  if (!input) return;
  const selectedMonth = getSelectedMonthKey();
  const [year, month] = selectedMonth.split('-');
  const daysInMonth = new Date(year, month, 0).getDate();
  input.min = `${selectedMonth}-01`;
  input.max = `${selectedMonth}-${daysInMonth}`;
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

function ensureMonthlySheetDayData(sheet, dayKey) {
  if (!sheet.days) sheet.days = {};
  if (!sheet.days[dayKey]) sheet.days[dayKey] = {};
  if (!sheet.days[dayKey].hours) sheet.days[dayKey].hours = {};
  if (!sheet.days[dayKey].manual) sheet.days[dayKey].manual = {};
  return sheet.days[dayKey];
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
    return `<div class="sheet-person-header-fit" title="${title}"><span class="sheet-person-header-line sheet-person-header-line--single">${firstName}</span></div>`;
  }

  return `<div class="sheet-person-header-fit" title="${title}"><span class="sheet-person-header-line sheet-person-header-line--first">${firstName}</span><span class="sheet-person-header-line sheet-person-header-line--last">${lastName}</span></div>`;
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
  const isArchived = state.isArchived === true;

  const items = [];

  if (!isArchived) {
    items.push({
      type: 'expense-form',
      targetId: '',
      icon: 'plus',
      title: 'Nowy Koszt / Zaliczka',
      meta: 'Koszty i Zaliczki',
      description: 'Otwórz panel „Nowy Dowód/Koszt”.'
    });
  }

  if (monthlySheets.length > 0) {
    items.push(...monthlySheets.map(sheet => ({
          type: 'monthly-sheet',
          targetId: sheet.id,
          icon: 'clock',
          title: 'Tabela z Godzinami',
          meta: sheet.site ? `${sheet.client} • ${sheet.site}` : (sheet.client || 'Bez klienta'),
          description: 'Przejdź do konkretnego arkusza godzin.'
        })));
  } else if (!isArchived) {
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
  } else if (!isArchived) {
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
      const tableContainer = document.querySelector('#sheet-detail-container .table-container');
      if (tableContainer) {
        const targetTop = Math.max(0, todayRow.offsetTop - ((tableContainer.clientHeight - todayRow.offsetHeight) / 2));
        tableContainer.scrollTo({ top: targetTop, behavior: 'smooth' });
        return;
      }

      todayRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  };

  requestAnimationFrame(scrollToToday);
  setTimeout(scrollToToday, 250);
}

function centerMonthlySheetTableInViewport() {
  const mainContent = document.getElementById('main-content');
  const tableContainer = document.querySelector('#sheet-detail-container .table-container');
  if (!mainContent || !tableContainer) return;

  const appScale = parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--app-scale')) || 1;
  const navLinks = document.querySelector('.nav-links');
  const isPortraitMobile = window.matchMedia('(max-width: 550px)').matches;
  const mainRect = mainContent.getBoundingClientRect();
  const tableRect = tableContainer.getBoundingClientRect();
  const tableTopInScroll = mainContent.scrollTop + ((tableRect.top - mainRect.top) / appScale);
  const tableHeight = tableRect.height / appScale;
  const bottomTabsHeight = isPortraitMobile && navLinks
    ? (navLinks.getBoundingClientRect().height / appScale)
    : 0;
  const bottomInset = 16 + bottomTabsHeight;
  const preferredTopOffset = Math.max(0, mainContent.clientHeight - tableHeight - bottomInset);
  const targetTop = Math.max(0, tableTopInScroll - preferredTopOffset);

  mainContent.scrollTo({ top: targetTop, behavior: 'smooth' });
}

function focusMonthlySheetDetail(sheetId = window.currentSheetId) {
  const sheet = Store.getMonthlySheet(sheetId);
  if (!sheet) return;

  const today = new Date();
  const todayMonthKey = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;
  const shouldScrollToToday = sheet.month === todayMonthKey;

  const syncFocus = () => {
    centerMonthlySheetTableInViewport();
    if (shouldScrollToToday) {
      scrollCurrentMonthlySheetToTodayIfNeeded(sheetId);
    }
  };

  requestAnimationFrame(syncFocus);
  setTimeout(syncFocus, 80);
  setTimeout(syncFocus, 220);
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

function buildSummaryPreviewRows(state, settlement, activePersons = Calculations.getActivePersons(state)) {
  const summaryRowsById = new Map();

  settlement.partners.forEach(p => {
    summaryRowsById.set(p.person.id, {
      personId: p.person.id,
      personType: p.person.type,
      name: getPersonDisplayName(p.person),
      badgeClass: 'badge-partner',
      label: 'Wspólnik',
      hours: p.hours,
      compensationMode: 'partner-like',
      netAmount: parseFloat(p.netAfterAccounting) || 0,
      grossAmount: parseFloat(p.toPayout) || 0,
      grossWithSalariesAmount: Number.isFinite(parseFloat(p.grossWithEmployeeSalaries))
        ? parseFloat(p.grossWithEmployeeSalaries)
        : (parseFloat(p.toPayout) || 0),
      showGrossWithSalaries: Math.abs((parseFloat(p.grossWithEmployeeSalaries) || 0) - (parseFloat(p.toPayout) || 0)) >= 0.005
    });
  });

  (settlement.separateCompanies || []).forEach(company => {
    summaryRowsById.set(company.person.id, {
      personId: company.person.id,
      personType: company.person.type,
      name: getPersonDisplayName(company.person),
      badgeClass: 'badge-partner',
      label: 'Osobna Firma',
      hours: company.hours,
      compensationMode: 'partner-like',
      netAmount: parseFloat(company.netAfterAccounting) || 0,
      showNetAmount: false,
      grossAmount: parseFloat(company.toPayout) || 0,
      grossWithSalariesAmount: Number.isFinite(parseFloat(company.grossWithEmployeeSalaries))
        ? parseFloat(company.grossWithEmployeeSalaries)
        : (parseFloat(company.toPayout) || 0),
      showGrossWithSalaries: Math.abs((parseFloat(company.grossWithEmployeeSalaries) || 0) - (parseFloat(company.toPayout) || 0)) >= 0.005
    });
  });

  settlement.workingPartners.forEach(wp => {
    summaryRowsById.set(wp.person.id, {
      personId: wp.person.id,
      personType: wp.person.type,
      name: getPersonDisplayName(wp.person),
      badgeClass: 'badge-working-partner',
      label: 'Wspólnik Pracujący',
      hours: wp.hours,
      compensationMode: 'partner-like',
      netAmount: parseFloat(wp.netAfterAccounting) || 0,
      grossAmount: parseFloat(wp.toPayout) || 0,
      grossWithSalariesAmount: Number.isFinite(parseFloat(wp.grossWithEmployeeSalaries))
        ? parseFloat(wp.grossWithEmployeeSalaries)
        : (parseFloat(wp.toPayout) || 0),
      showGrossWithSalaries: Math.abs((parseFloat(wp.grossWithEmployeeSalaries) || 0) - (parseFloat(wp.toPayout) || 0)) >= 0.005
    });
  });

  settlement.employees.forEach(e => {
    summaryRowsById.set(e.person.id, {
      personId: e.person.id,
      personType: e.person.type,
      name: getPersonDisplayName(e.person),
      badgeClass: 'badge-employee',
      label: 'Pracownik',
      hours: e.hours,
      compensationMode: 'single',
      amount: e.toPayout,
      bonusAmount: parseFloat(e.bonusAmount) || 0,
      showBonusAmount: (parseFloat(e.bonusAmount) || 0) > 0,
      dietaAmount: parseFloat(e.dietaAmount) || 0,
      showDietaAmount: (parseFloat(e.dietaAmount) || 0) > 0,
      advancesTaken: parseFloat(e.advancesTaken) || 0,
      showAdvancesTaken: (parseFloat(e.advancesTaken) || 0) > 0
    });
  });

  return activePersons
    .map(person => summaryRowsById.get(person.id))
    .filter(Boolean);
}

function renderSummaryPreviewTable(tbody, previewRows = [], options = {}) {
  if (!tbody) return;

  const {
    emptyMessage = 'Brak dodanych osób. Dodaj osoby w zakładce "Osoby".',
    dividerColumnCount = 4
  } = options;

  tbody.innerHTML = '';
  if (!previewRows.length) {
    tbody.innerHTML = `<tr><td colspan="4" style="text-align:center;color:var(--text-muted)">${emptyMessage}</td></tr>`;
    return;
  }

  previewRows.forEach((row, index) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${row.name}</td>
      <td><span class="badge ${row.badgeClass}">${row.label}</span></td>
      <td>${row.hours.toFixed(1)}h</td>
      <td>${getSummaryPreviewCompensationHtml(row)}</td>
    `;
    tbody.appendChild(tr);
    appendMobileCardDividerRow(tbody, dividerColumnCount, index === previewRows.length - 1);
  });
}

function getSummaryPreviewCompensationHtml(row) {
  if (!row) return '';

  if (row.compensationMode === 'partner-like') {
    return `
      <div class="summary-preview-compensation">
        ${row.showNetAmount !== false ? `
        <div class="summary-preview-compensation-row">
          <span>Netto</span>
          <strong>${row.netAmount.toFixed(2)} zł</strong>
        </div>
        ` : ''}
        <div class="summary-preview-compensation-row">
          <span>Brutto</span>
          <strong>${row.grossAmount.toFixed(2)} zł</strong>
        </div>
        ${row.showGrossWithSalaries ? `
        <div class="summary-preview-compensation-row summary-preview-compensation-row--warning">
          <span>Z pensjami</span>
          <strong>${row.grossWithSalariesAmount.toFixed(2)} zł</strong>
        </div>
        ` : ''}
      </div>
    `;
  }

  return `
    <div class="summary-preview-compensation summary-preview-compensation--single">
      <div class="summary-preview-compensation-row">
        <span>Do wypłaty</span>
        <strong>${row.amount.toFixed(2)} zł</strong>
      </div>
      ${row.showBonusAmount ? `
      <div class="summary-preview-compensation-row summary-preview-compensation-row--warning">
        <span>Premie</span>
        <strong>+${row.bonusAmount.toFixed(2)} zł</strong>
      </div>
      ` : ''}
      ${row.showDietaAmount ? `
      <div class="summary-preview-compensation-row summary-preview-compensation-row--warning">
        <span>Dieta</span>
        <strong>+${row.dietaAmount.toFixed(2)} zł</strong>
      </div>
      ` : ''}
      ${row.showAdvancesTaken ? `
      <div class="summary-preview-compensation-row summary-preview-compensation-row--danger">
        <span>Zaliczki</span>
        <strong>-${row.advancesTaken.toFixed(2)} zł</strong>
      </div>
      ` : ''}
    </div>
  `;
}

function renderSummary() {
  const state = Store.getState();
  const settlement = Calculations.generateSettlement(state);
  const activePersons = Calculations.getActivePersons(state);
  const previewRows = buildSummaryPreviewRows(state, settlement, activePersons);

  renderSummaryShortcuts(state);

  document.getElementById('dash-people-count').textContent = activePersons.length;
  document.getElementById('dash-total-hours').textContent = `${settlement.totalTeamHours.toFixed(1)}h`;
  document.getElementById('dash-total-revenue').textContent = `${settlement.commonRevenue.toFixed(2)} zł`;
  document.getElementById('dash-profit').textContent = `${settlement.profitToSplit.toFixed(2)} zł`;

  const tbody = document.getElementById('dash-quick-preview');
  renderSummaryPreviewTable(tbody, previewRows);
}

const REPORTS_PREVIEW_DOCUMENT_WIDTH_MM = 186;
const REPORTS_PREVIEW_DOCUMENT_HEIGHT_MM = 273;
const REPORTS_PREVIEW_ZOOM_MIN = 30;
const REPORTS_PREVIEW_ZOOM_MAX = 200;
const REPORTS_PREVIEW_ZOOM_STEP = 5;
const REPORTS_HOURS_ROWS_PER_PAGE = 25; // Base rows per page for the new 100% (physical 0.7)
const REPORTS_PRINT_PAGE_MARGIN_MM = 12;

function getReportsPageMetrics(orientation = 'portrait') {
  const isLandscape = orientation === 'landscape';
  const sheetWidthMm = isLandscape ? 297 : 210;
  const sheetHeightMm = isLandscape ? 210 : 297;
  return {
    sheetWidthMm,
    sheetHeightMm,
    contentWidthMm: sheetWidthMm - (REPORTS_PRINT_PAGE_MARGIN_MM * 2),
    contentHeightMm: sheetHeightMm - (REPORTS_PRINT_PAGE_MARGIN_MM * 2),
    marginMm: REPORTS_PRINT_PAGE_MARGIN_MM
  };
}

function clampReportsPreviewZoom(value = 100) {
  const parsed = parseFloat(value);
  if (!Number.isFinite(parsed)) return 100;
  const clamped = Math.max(REPORTS_PREVIEW_ZOOM_MIN, Math.min(REPORTS_PREVIEW_ZOOM_MAX, parsed));
  return Math.round(clamped / REPORTS_PREVIEW_ZOOM_STEP) * REPORTS_PREVIEW_ZOOM_STEP;
}

function getReportsViewState(activePersons = []) {
  if (!window.reportViewState) {
    window.reportViewState = {
      personId: 'all',
      employeeVisibility: 'internal',
      previewZoomMode: 'fit-width',
      previewZoomPercent: 100,
      previewZoom: 100,
      previewTheme: 'app',
      printOrientation: 'portrait',
      printContentScale: 'auto'
    };
  }

  const activeIds = new Set(activePersons.map(person => person.id));
  if (activePersons.length > 0 && window.reportViewState.personId !== 'all' && !activeIds.has(window.reportViewState.personId)) {
    window.reportViewState.personId = 'all';
  }

  if (!['internal', 'employee'].includes(window.reportViewState.employeeVisibility)) {
    window.reportViewState.employeeVisibility = 'internal';
  }

  if (!['fit-width', 'fit-page', 'manual'].includes(window.reportViewState.previewZoomMode)) {
    window.reportViewState.previewZoomMode = 'manual';
  }

  window.reportViewState.previewTheme = 'print';

  if (!['portrait', 'landscape'].includes(window.reportViewState.printOrientation)) {
    window.reportViewState.printOrientation = 'portrait';
  }

  if (window.reportViewState.printContentScale !== 'auto' && (typeof window.reportViewState.printContentScale !== 'number' || window.reportViewState.printContentScale < 0.2)) {
    window.reportViewState.printContentScale = 1.0;
  }

  window.reportViewState.previewZoomPercent = clampReportsPreviewZoom(window.reportViewState.previewZoomPercent);

  return window.reportViewState;
}

function syncReportsPersonFilterOptions(select, activePersons = [], selectedPersonId = 'all') {
  if (!select) return 'all';

  const optionsHtml = ['<option value="all">Wszystkie aktywne osoby</option>']
    .concat(activePersons.map(person => `<option value="${escapeReportHtml(person.id)}">${escapeReportHtml(getPersonDisplayName(person))} • ${escapeReportHtml(getReportPersonTypeLabel(person.type))}</option>`))
    .join('');

  if (select.innerHTML.trim() !== optionsHtml.trim()) {
    select.innerHTML = optionsHtml;
  }

  const activePersonIds = new Set(activePersons.map(person => person.id));
  const resolvedPersonId = selectedPersonId !== 'all' && activePersonIds.has(selectedPersonId)
    ? selectedPersonId
    : 'all';

  select.value = resolvedPersonId;
  return resolvedPersonId;
}

function getReportsSelectedPersonsData(options = {}) {
  const state = Store.getState();
  const month = getSelectedMonthKey();
  const settlement = Calculations.generateSettlement(state);
  const invoiceReconciliation = settlement.invoiceReconciliation || Calculations.calculateInvoiceReconciliation(state, month);
  const activePersons = Calculations.getActivePersons(state, month);
  const viewState = getReportsViewState(activePersons);
  const personId = options.personId || viewState.personId || 'all';
  const employeeVisibility = options.employeeVisibility || viewState.employeeVisibility || 'internal';
  const selectedPersons = personId === 'all'
    ? activePersons
    : activePersons.filter(person => person.id === personId);

  return {
    state,
    month,
    settlement,
    invoiceReconciliation,
    activePersons,
    viewState,
    personId,
    employeeVisibility,
    selectedPersons
  };
}

function buildReportsPreviewTitle(selectionData = {}) {
  const { personId = 'all', month = getSelectedMonthKey(), selectedPersons = [] } = selectionData;
  return personId === 'all'
    ? `${month.replace('-', '_')} - Zestawienie Wszystkich - ${formatMonthLabel(month)}`
    : `${month.replace('-', '_')} - Zestawienie ${getPersonDisplayName(selectedPersons[0])} - ${formatMonthLabel(month)}`;
}

function buildReportsPreviewBodyHtml(selectionData = {}) {
  return buildReportsUnifiedDocumentHtml(selectionData, {
    contentScale: selectionData?.viewState?.printContentScale ?? 'auto',
    printOrientation: selectionData?.viewState?.printOrientation || 'portrait',
    employeeVisibility: selectionData?.employeeVisibility || 'internal'
  });
}

function buildReportsUnifiedDocumentHtml(selectionData = {}, options = {}) {
  const {
    state,
    month,
    settlement,
    invoiceReconciliation,
    selectedPersons = []
  } = selectionData;

  if (!Array.isArray(selectedPersons) || selectedPersons.length === 0) {
    return '<p style="text-align:center; padding:2rem; color:var(--text-secondary)">Brak aktywnych osób do przygotowania zestawienia.</p>';
  }

  const renderOptions = {
    printMode: true,
    contentScale: options.contentScale ?? selectionData?.viewState?.printContentScale ?? 1.0,
    employeeVisibility: options.employeeVisibility ?? selectionData?.employeeVisibility ?? selectionData?.viewState?.employeeVisibility ?? 'internal',
    printOrientation: options.printOrientation ?? selectionData?.viewState?.printOrientation ?? 'portrait'
  };

  return selectedPersons
    .map(person => buildPersonReportCardHtml(
      buildPersonReportData(state, settlement, person, month, invoiceReconciliation),
      renderOptions
    ))
    .join('');
}




function escapeReportHtml(value = '') {
  return (value ?? '')
    .toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function formatReportCurrency(value = 0) {
  return formatPolishCurrencyWithSuffix(parseFloat(value) || 0);
}

function formatReportAmountCell(value = 0) {
  const amount = parseFloat(value) || 0;
  if (Math.abs(amount) < 0.005) return '';
  return escapeReportHtml(formatReportCurrency(amount));
}

function formatReportHours(value = 0) {
  return `${(parseFloat(value) || 0).toFixed(1)}h`;
}

function formatReportHoursCell(value = null) {
  if (value === null || value === '' || typeof value === 'undefined') return '';
  const parsed = parseFloat(value);
  if (!Number.isFinite(parsed) || Math.abs(parsed) < 0.005) return '';
  return `${parsed.toFixed(1)}h`;
}

function getReportPersonTypeLabel(type = '') {
  switch (type) {
    case 'PARTNER': return 'Wspólnik';
    case 'WORKING_PARTNER': return 'Wspólnik Pracujący';
    case 'SEPARATE_COMPANY': return 'Osobna Firma';
    case 'EMPLOYEE': return 'Pracownik';
    default: return type || 'Osoba';
  }
}

function getSettlementEntryForPerson(settlement, personId = '') {
  const collections = [
    { role: 'partner', entries: settlement?.partners || [] },
    { role: 'workingPartner', entries: settlement?.workingPartners || [] },
    { role: 'separateCompany', entries: settlement?.separateCompanies || [] },
    { role: 'employee', entries: settlement?.employees || [] }
  ];

  for (const collection of collections) {
    const entry = collection.entries.find(item => item?.person?.id === personId);
    if (entry) return { role: collection.role, entry };
  }

  return { role: '', entry: null };
}

function getReportExpenseColumns(state, personId, month) {
  const expenses = (state?.expenses || []).filter(expense => Calculations.isDateInMonth(expense?.date, month));
  const definitions = [
    {
      key: 'costsPaid',
      label: 'Koszty zapłacone',
      matches: expense => expense?.type === 'COST' && expense?.paidById === personId,
      amount: expense => parseFloat(expense?.amount) || 0
    },
    {
      key: 'advancesPaid',
      label: 'Zaliczki zapłacone',
      matches: expense => expense?.type === 'ADVANCE' && expense?.paidById === personId,
      amount: expense => parseFloat(expense?.amount) || 0
    },
    {
      key: 'advancesTaken',
      label: 'Zaliczki pobrane',
      matches: expense => expense?.type === 'ADVANCE' && expense?.advanceForId === personId,
      amount: expense => parseFloat(expense?.amount) || 0
    },
    {
      key: 'bonuses',
      label: 'Premie',
      matches: expense => expense?.type === 'BONUS' && expense?.advanceForId === personId,
      amount: expense => Calculations.getExpenseEffectiveAmount(expense, state, month)
    },
    {
      key: 'dietas',
      label: 'Diety',
      matches: expense => expense?.type === 'DIETA' && expense?.advanceForId === personId,
      amount: expense => Calculations.getExpenseEffectiveAmount(expense, state, month)
    },
    {
      key: 'refundsPaid',
      label: 'Zwroty wypłacone',
      matches: expense => expense?.type === 'REFUND' && expense?.paidById === personId,
      amount: expense => Calculations.getExpenseEffectiveAmount(expense, state, month)
    },
    {
      key: 'refundsReceived',
      label: 'Zwroty otrzymane',
      matches: expense => expense?.type === 'REFUND',
      amount: expense => Calculations.getRefundReceivedAmountForPerson
        ? Calculations.getRefundReceivedAmountForPerson(expense, personId, state, month)
        : 0
    }
  ];

  return definitions.map(definition => {
    const valuesByDate = {};
    let total = 0;

    expenses.forEach(expense => {
      if (!definition.matches(expense)) return;
      const amount = parseFloat(definition.amount(expense)) || 0;
      if (Math.abs(amount) < 0.005) return;
      valuesByDate[expense.date] = (valuesByDate[expense.date] || 0) + amount;
      total += amount;
    });

    if (!Object.keys(valuesByDate).length && Math.abs(total) < 0.005) return null;
    return { ...definition, valuesByDate, total };
  }).filter(Boolean);
}

function buildPersonHoursAggregateData(state, person, month) {
  const monthlySheets = (state?.monthlySheets || []).filter(sheet => sheet?.month === month);
  const [year, monthNumber] = month.split('-').map(Number);
  const daysInMonth = new Date(year, monthNumber, 0).getDate();
  const columns = [];

  monthlySheets.forEach(sheet => {
    const visiblePersonIds = new Set((getVisiblePersonsForSheet(state, sheet) || []).map(candidate => candidate.id));
    const hasExplicitHours = Object.values(sheet?.days || {}).some(day => typeof day?.hours?.[person.id] !== 'undefined');
    if (!visiblePersonIds.has(person.id) && !hasExplicitHours) return;

    const valuesByDay = {};
    let totalHours = 0;

    for (let day = 1; day <= daysInMonth; day++) {
      const dayData = sheet?.days?.[day] || {};
      const isInactive = typeof isMonthlySheetPersonInactiveOnDay === 'function'
        ? isMonthlySheetPersonInactiveOnDay(sheet, person.id, day)
        : false;
      const rawHours = dayData?.hours?.[person.id];
      const parsedHours = isInactive ? 0 : (Number.isFinite(parseFloat(rawHours)) ? parseFloat(rawHours) : null);
      if (Number.isFinite(parsedHours)) totalHours += parsedHours;
      valuesByDay[day] = {
        timeLabel: dayData?.globalStart && dayData?.globalEnd ? `${dayData.globalStart}-${dayData.globalEnd}` : '',
        hours: parsedHours
      };
    }

    if (totalHours > 0.005) {
      columns.push({
        sheetId: sheet.id,
        client: sheet.client || 'Klient',
        site: sheet.site || 'Budowa',
        label: `${sheet.client || 'Klient'} • ${sheet.site || 'Budowa'}`,
        valuesByDay,
        totalHours
      });
    }
  });

  const expenseColumns = getReportExpenseColumns(state, person.id, month);
  const holidays = typeof getPolishHolidays === 'function' ? getPolishHolidays(year) : {};
  const rows = [];
  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(year, monthNumber - 1, day);
    const dateKey = `${month}-${String(day).padStart(2, '0')}`;
    const dateStr = `${String(monthNumber).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    
    rows.push({
      day,
      dateKey,
      isHoliday: date.getDay() === 0 || !!holidays[dateStr],
      isSaturday: date.getDay() === 6,
      dateLabel: `${String(day).padStart(2, '0')}.${String(monthNumber).padStart(2, '0')}`,
      weekdayLabel: date.toLocaleDateString('pl-PL', { weekday: 'short' }),
      sheetValues: columns.map(column => column.valuesByDay[day] || { timeLabel: '', hours: null }),
      expenseValues: expenseColumns.map(column => column.valuesByDate[dateKey] || 0)
    });
  }

  return {
    columns,
    expenseColumns,
    rows,
    hasData: columns.length > 0 || expenseColumns.length > 0
  };
}

function buildPersonWorksAggregateData(state, person, month) {
  const worksSheets = (state?.worksSheets || []).filter(sheet => sheet?.month === month);
  const sections = worksSheets.map(sheet => {
    const items = (sheet?.entries || []).map(entry => {
      const personHours = parseFloat(entry?.hours?.[person.id]);
      if (!Number.isFinite(personHours) || personHours <= 0) return null;

      const metrics = Calculations.getWorksEntryMetrics(entry, sheet, state);
      return {
        entryId: entry.id,
        date: entry.date || '',
        name: entry.name || 'Pozycja',
        quantityText: metrics.isRoboczogodziny
          ? `${personHours.toFixed(2)} h`
          : `${parseFloat(entry.quantity) || 0} ${entry.unit || ''}`.trim(),
        priceText: metrics.isRoboczogodziny
          ? replaceCurrencySuffix(formatReportCurrency(Calculations.getSheetClientRate(sheet, state)), 'zł/h')
          : formatReportCurrency(entry.price || 0),
        totalValueText: formatReportCurrency(metrics.revenue || 0),
        personHours,
        personValueText: formatReportCurrency(metrics.isRoboczogodziny
          ? personHours * Calculations.getSheetClientRate(sheet, state)
          : metrics.revenue || 0)
      };
    }).filter(Boolean);

    if (!items.length) return null;
    return {
      sheetId: sheet.id,
      client: sheet.client || 'Klient',
      site: sheet.site || 'Budowa',
      label: `${sheet.client || 'Klient'} • ${sheet.site || 'Budowa'}`,
      totalHours: items.reduce((sum, item) => sum + (item.personHours || 0), 0),
      items
    };
  }).filter(Boolean);

  return {
    sections,
    hasData: sections.length > 0
  };
}

function buildEmployerSettlementData(settlement, person) {
  const employeeItems = (settlement?.employees || [])
    .filter(entry => entry?.person?.employerId === person.id)
    .map(entry => ({
      typeLabel: 'Pracownik',
      person: entry.person,
      hours: entry.hours || 0,
      salary: entry.salary || 0,
      benefits: (entry.bonusAmount || 0) + (entry.dietaAmount || 0),
      taxAmount: entry.contractTaxAmount || 0,
      zusAmount: entry.contractZusAmount || 0,
      toPayout: entry.toPayout || 0,
      note: entry.contractChargesPaidByEmployer ? 'Podatek/ZUS opłaca pracodawca' : ''
    }));

  const workingPartnerItems = (settlement?.workingPartners || [])
    .filter(entry => entry?.person?.employerId === person.id)
    .map(entry => ({
      typeLabel: 'Wspólnik pracujący',
      person: entry.person,
      hours: entry.hours || 0,
      salary: entry.salary || 0,
      benefits: 0,
      taxAmount: (entry.ownTaxAmount || 0) + (entry.contractTaxAmount || 0),
      zusAmount: entry.contractZusAmount || 0,
      toPayout: entry.netAfterAccounting || 0,
      note: entry.contractChargesPaidByEmployer ? 'Część UZ opłaca pracodawca' : ''
    }));

  const items = [...employeeItems, ...workingPartnerItems];
  return {
    items,
    totals: {
      salary: items.reduce((sum, item) => sum + (item.salary || 0), 0),
      benefits: items.reduce((sum, item) => sum + (item.benefits || 0), 0),
      taxes: items.reduce((sum, item) => sum + (item.taxAmount || 0), 0),
      zus: items.reduce((sum, item) => sum + (item.zusAmount || 0), 0),
      payouts: items.reduce((sum, item) => sum + (item.toPayout || 0), 0)
    },
    hasData: items.length > 0
  };
}

function buildPersonReportData(state, settlement, person, month, invoiceReconciliation) {
  const { role, entry } = getSettlementEntryForPerson(settlement, person.id);
  const summaryRow = buildSummaryPreviewRows(state, settlement, [person])[0] || null;
  const invoiceEntry = (invoiceReconciliation?.issuers || []).find(item => item?.issuerId === person.id) || null;
  const hoursData = buildPersonHoursAggregateData(state, person, month);
  const worksData = buildPersonWorksAggregateData(state, person, month);
  const payrollData = buildEmployerSettlementData(settlement, person);

  return {
    person,
    role,
    entry,
    summaryRow,
    invoiceEntry,
    hoursData,
    worksData,
    payrollData,
    month
  };
}

function buildReportMetricListHtml(items = []) {
  const visibleItems = items.filter(item => item && item.hide !== true);
  if (!visibleItems.length) {
    return '<p class="settlement-detail-empty">Brak danych do pokazania.</p>';
  }

  return `
    <div class="settlement-detail-stack">
      ${visibleItems.map(item => {
        if (item.isGroupBox) {
          return `
            <div class="settlement-cost-group-box">
              ${item.items.filter(sub => sub && sub.hide !== true).map(sub => `
                <div class="settlement-detail-row">
                  <span>${escapeReportHtml(sub.label)}</span>
                  <strong class="${sub.tone || ''}">${sub.valueHtml}</strong>
                </div>
              `).join('')}
            </div>
          `;
        }
        return `
          <div class="settlement-detail-row ${item.isDivider ? 'settlement-detail-row--divider' : ''}">
            <span>${escapeReportHtml(item.label)}</span>
            <strong class="${item.tone || ''}">${item.valueHtml}</strong>
          </div>
        `;
      }).join('')}
    </div>
  `;
}

function buildReportSettlementSummaryHtml(reportData, options = {}) {
  const { person, entry } = reportData;
  if (!entry) {
    return '<p class="settlement-detail-empty">Brak danych rozliczenia dla tej osoby w wybranym miesiącu.</p>';
  }

  const employeeVisibility = options.employeeVisibility || 'internal';
  const rows = [];
  const pt = person.type;

  const hoursLabel = (pt === 'EMPLOYEE' || pt === 'WORKING_PARTNER') && entry.effectiveRate > 0 
    ? `Godziny: <span style="font-weight: normal; margin-left: 0.3rem;">(${formatReportCurrency(entry.effectiveRate)}/h)</span>`
    : 'Godziny';

  // PARTNER
  if (pt === 'PARTNER') {
    rows.push(
      { label: 'Godziny', valueHtml: escapeReportHtml(formatReportHours(entry.hours || 0)) },
      { label: 'Zarobek z własnych godzin', valueHtml: escapeReportHtml(formatReportCurrency(entry.salary || 0)) },
      { label: 'Podział zysku z pracowników', valueHtml: escapeReportHtml(formatReportCurrency(entry.revenueShare || 0)) },
      { label: 'Przychód własny (Brutto)', valueHtml: escapeReportHtml(formatReportCurrency(entry.ownGrossAmount || 0)), tone: 'settlement-accent-bold' },
      { label: 'Zwrot kosztów', valueHtml: `+${escapeReportHtml(formatReportCurrency(entry.paidCosts || 0))}`, tone: 'settlement-accent-positive', hide: (entry.paidCosts || 0) < 0.005 },
      { label: 'Zwrot zaliczek', valueHtml: `+${escapeReportHtml(formatReportCurrency(entry.paidAdvances || 0))}`, tone: 'settlement-accent-positive', hide: (entry.paidAdvances || 0) < 0.005 },
      { label: 'Zwroty otrzymane', valueHtml: `+${escapeReportHtml(formatReportCurrency(entry.refundsReceived || 0))}`, tone: 'settlement-accent-positive', hide: (entry.refundsReceived || 0) < 0.005 },
      { label: 'Zwroty wypłacone', valueHtml: `-${escapeReportHtml(formatReportCurrency(entry.refundsPaid || 0))}`, tone: 'settlement-accent-negative', hide: (entry.refundsPaid || 0) < 0.005 },
      { label: 'Udział w kosztach', valueHtml: `- ${escapeReportHtml(formatReportCurrency(entry.costShareApplied || 0))}`, tone: 'settlement-accent-negative', hide: (entry.costShareApplied || 0) < 0.005 },
      { label: 'Zaliczki', valueHtml: `-${escapeReportHtml(formatReportCurrency(entry.advancesTaken || 0))}`, tone: 'settlement-accent-negative', hide: (entry.advancesTaken || 0) < 0.005 },
      { label: 'Przychód (Brutto)', valueHtml: escapeReportHtml(formatReportCurrency(entry.toPayout || 0)), tone: 'settlement-accent-positive settlement-accent-bold' },
      { label: 'Pensje pracowników', valueHtml: escapeReportHtml(formatReportCurrency(entry.employeeSalaries || 0)), tone: 'settlement-accent-warning', hide: (entry.employeeSalaries || 0) < 0.005 },
      { label: 'Zwrot Podatku i ZUS za pracowników', valueHtml: escapeReportHtml(formatReportCurrency(entry.employeeAccountingRefund || 0)), tone: 'settlement-accent-warning', hide: (entry.employeeAccountingRefund || 0) < 0.005 },
      { label: 'Do odebrania od pracowników', valueHtml: `-${escapeReportHtml(formatReportCurrency(entry.employeeReceivables || 0))}`, tone: 'settlement-accent-negative', hide: (entry.employeeReceivables || 0) < 0.005 },
      { label: 'Przychód (Brutto) z Pensjami', valueHtml: escapeReportHtml(formatReportCurrency(entry.grossWithEmployeeSalaries || 0)), tone: 'settlement-accent-warning settlement-accent-bold', hide: (entry.employeeSalaries || 0) < 0.005 && (entry.employeeReceivables || 0) < 0.005 },
      {
        isGroupBox: true,
        items: [
          { label: 'Podatek własny', valueHtml: escapeReportHtml(formatReportCurrency(entry.ownTaxAmount || 0)), tone: 'settlement-accent-negative' },
          { label: 'Podatek wspólny', valueHtml: escapeReportHtml(formatReportCurrency(entry.sharedCompanyTaxAmount || 0)), tone: 'settlement-accent-negative' },
          { label: 'ZUS', valueHtml: escapeReportHtml(formatReportCurrency(entry.zusAmount || 0)), tone: 'settlement-accent-negative' }
        ]
      },
      { label: 'Przychód (Netto)', valueHtml: escapeReportHtml(formatReportCurrency(entry.netAfterAccounting || 0)), tone: 'settlement-accent-positive settlement-accent-bold' }
    );
  }

  // WORKING PARTNER
  if (pt === 'WORKING_PARTNER') {
    rows.push(
      { label: 'Godziny', valueHtml: `${escapeReportHtml(formatReportHours(entry.hours || 0))} ${entry.effectiveRate > 0 ? `(${formatReportCurrency(entry.effectiveRate)}/h)` : ''}` },
      { label: 'Zarobek', valueHtml: escapeReportHtml(formatReportCurrency(entry.salary || 0)) },
      { label: 'Zarobek (Wykonane Prace)', valueHtml: escapeReportHtml(formatReportCurrency(entry.worksShare || 0)), tone: 'settlement-accent-primary', hide: (entry.worksShare || 0) < 0.005 },
      { label: 'Zwrot kosztów', valueHtml: `+${escapeReportHtml(formatReportCurrency(entry.paidCosts || 0))}`, tone: 'settlement-accent-positive', hide: (entry.paidCosts || 0) < 0.005 },
      { label: 'Zwrot zaliczek', valueHtml: `+${escapeReportHtml(formatReportCurrency(entry.paidAdvances || 0))}`, tone: 'settlement-accent-positive', hide: (entry.paidAdvances || 0) < 0.005 },
      { label: 'Zwroty otrzymane', valueHtml: `+${escapeReportHtml(formatReportCurrency(entry.refundsReceived || 0))}`, tone: 'settlement-accent-positive', hide: (entry.refundsReceived || 0) < 0.005 },
      { label: 'Zwroty wypłacone', valueHtml: `-${escapeReportHtml(formatReportCurrency(entry.refundsPaid || 0))}`, tone: 'settlement-accent-negative', hide: (entry.refundsPaid || 0) < 0.005 },
      { label: 'Udział w kosztach', valueHtml: `- ${escapeReportHtml(formatReportCurrency(entry.costShareApplied || 0))}`, tone: 'settlement-accent-negative', hide: (entry.costShareApplied || 0) < 0.005 },
      { label: 'Zaliczki pobrane', valueHtml: `-${escapeReportHtml(formatReportCurrency(entry.advancesTaken || 0))}`, tone: 'settlement-accent-negative', hide: (entry.advancesTaken || 0) < 0.005 },
      { label: 'Przychód (Brutto)', valueHtml: escapeReportHtml(formatReportCurrency(entry.toPayout || 0)), tone: 'settlement-accent-positive settlement-accent-bold' },
      { label: 'Pensje pracowników', valueHtml: escapeReportHtml(formatReportCurrency(entry.employeeSalaries || 0)), tone: 'settlement-accent-warning', hide: (entry.employeeSalaries || 0) < 0.005 },
      { label: 'Zwrot Podatku i ZUS za pracowników', valueHtml: escapeReportHtml(formatReportCurrency(entry.employeeAccountingRefund || 0)), tone: 'settlement-accent-warning', hide: (entry.employeeAccountingRefund || 0) < 0.005 },
      { label: 'Do odebrania od pracowników', valueHtml: `-${escapeReportHtml(formatReportCurrency(entry.employeeReceivables || 0))}`, tone: 'settlement-accent-negative', hide: (entry.employeeReceivables || 0) < 0.005 },
      { label: 'Podatek własny', valueHtml: `-${escapeReportHtml(formatReportCurrency(entry.ownTaxAmount || 0))}`, tone: 'settlement-accent-negative', hide: (entry.ownTaxAmount || 0) < 0.005 },
      { label: `Podatek Umowa Zlecenie${entry.contractChargesPaidByEmployer ? ' (opłaca pracodawca)' : ''}`, valueHtml: `${entry.contractChargesPaidByEmployer ? '' : '- '}${escapeReportHtml(formatReportCurrency(entry.contractTaxAmount || 0))}`, tone: entry.contractChargesPaidByEmployer ? 'settlement-accent-warning' : 'settlement-accent-negative', hide: (entry.contractTaxAmount || 0) < 0.005 },
      { label: `ZUS Umowa Zlecenie${entry.contractChargesPaidByEmployer ? ' (opłaca pracodawca)' : ''}`, valueHtml: `${entry.contractChargesPaidByEmployer ? '' : '- '}${escapeReportHtml(formatReportCurrency(entry.contractZusAmount || 0))}`, tone: entry.contractChargesPaidByEmployer ? 'settlement-accent-warning' : 'settlement-accent-negative', hide: (entry.contractZusAmount || 0) < 0.005 },
      { label: 'Przychód (Netto)', valueHtml: escapeReportHtml(formatReportCurrency(entry.netAfterAccounting || 0)), tone: 'settlement-accent-positive settlement-accent-bold' }
    );
  }

  // EMPLOYEE
  if (pt === 'EMPLOYEE') {
    rows.push(
      { label: 'Godziny', valueHtml: `${escapeReportHtml(formatReportHours(entry.hours || 0))} ${entry.effectiveRate > 0 ? `(${formatReportCurrency(entry.effectiveRate)}/h)` : ''}` },
      { label: 'Zarobek', valueHtml: escapeReportHtml(formatReportCurrency(entry.salary || 0)) },
      { label: 'Premia', valueHtml: `+${escapeReportHtml(formatReportCurrency(entry.bonusAmount || 0))}`, tone: 'settlement-accent-warning', hide: (entry.bonusAmount || 0) < 0.005 },
      { label: 'Diety', valueHtml: `+${escapeReportHtml(formatReportCurrency(entry.dietaAmount || 0))}`, tone: 'settlement-accent-warning', hide: (entry.dietaAmount || 0) < 0.005 },
      { label: 'Zwrot kosztów', valueHtml: `+${escapeReportHtml(formatReportCurrency(entry.paidCosts || 0))}`, tone: 'settlement-accent-positive', hide: (entry.paidCosts || 0) < 0.005 },
      { label: 'Zwrot zaliczek', valueHtml: `+${escapeReportHtml(formatReportCurrency(entry.paidAdvances || 0))}`, tone: 'settlement-accent-positive', hide: (entry.paidAdvances || 0) < 0.005 },
      { label: 'Zwroty otrzymane', valueHtml: `+${escapeReportHtml(formatReportCurrency(entry.refundsReceived || 0))}`, tone: 'settlement-accent-positive', hide: (entry.refundsReceived || 0) < 0.005 },
      { label: 'Zwroty wypłacone', valueHtml: `-${escapeReportHtml(formatReportCurrency(entry.refundsPaid || 0))}`, tone: 'settlement-accent-negative', hide: (entry.refundsPaid || 0) < 0.005 },
      { label: 'Zaliczki pobrane', valueHtml: `-${escapeReportHtml(formatReportCurrency(entry.advancesTaken || 0))}`, tone: 'settlement-accent-negative', hide: (entry.advancesTaken || 0) < 0.005 },
      { label: `Podatek Umowa Zlecenie${entry.contractChargesPaidByEmployer ? ' (opłaca pracodawca)' : ''}`, valueHtml: `${entry.contractChargesPaidByEmployer ? '' : '- '}${escapeReportHtml(formatReportCurrency(entry.contractTaxAmount || 0))}`, tone: entry.contractChargesPaidByEmployer ? 'settlement-accent-warning' : 'settlement-accent-negative', hide: (entry.contractTaxAmount || 0) < 0.005 },
      { label: `ZUS Umowa Zlecenie${entry.contractChargesPaidByEmployer ? ' (opłaca pracodawca)' : ''}`, valueHtml: `${entry.contractChargesPaidByEmployer ? '' : '- '}${escapeReportHtml(formatReportCurrency(entry.contractZusAmount || 0))}`, tone: entry.contractChargesPaidByEmployer ? 'settlement-accent-warning' : 'settlement-accent-negative', hide: (entry.contractZusAmount || 0) < 0.005 },
      { label: 'Wypracowany zysk z godzin (bez UZ)', valueHtml: escapeReportHtml(getEmployeeGeneratedProfitDisplay(entry).valueText), tone: 'settlement-accent-primary', hide: employeeVisibility === 'employee' },
      { label: 'Do Wypłaty', valueHtml: escapeReportHtml(formatReportCurrency(entry.toPayout || 0)), tone: 'settlement-accent-positive settlement-accent-bold' }
    );
  }

  // SEPARATE COMPANY
  if (pt === 'SEPARATE_COMPANY') {
    rows.push(
      { label: 'Godziny', valueHtml: `${escapeReportHtml(formatReportHours(entry.hours || 0))} ${entry.effectiveRate > 0 ? `(${formatReportCurrency(entry.effectiveRate)}/h)` : ''}` },
      { label: 'Zarobek z własnych godzin', valueHtml: escapeReportHtml(formatReportCurrency(entry.salary || 0)) },
      { label: 'Podział zysku z pracowników', valueHtml: escapeReportHtml(formatReportCurrency(entry.revenueShare || 0)) },
      { label: 'Zarobek (Wykonane Prace)', valueHtml: escapeReportHtml(formatReportCurrency(entry.worksShare || 0)), tone: 'settlement-accent-primary', hide: (entry.worksShare || 0) < 0.005 },
      { label: 'Zwrot kosztów', valueHtml: `+${escapeReportHtml(formatReportCurrency(entry.paidCosts || 0))}`, tone: 'settlement-accent-positive', hide: (entry.paidCosts || 0) < 0.005 },
      { label: 'Zwrot zaliczek', valueHtml: `+${escapeReportHtml(formatReportCurrency(entry.paidAdvances || 0))}`, tone: 'settlement-accent-positive', hide: (entry.paidAdvances || 0) < 0.005 },
      { label: 'Zwroty otrzymane', valueHtml: `+${escapeReportHtml(formatReportCurrency(entry.refundsReceived || 0))}`, tone: 'settlement-accent-positive', hide: (entry.refundsReceived || 0) < 0.005 },
      { label: 'Zwroty wypłacone', valueHtml: `-${escapeReportHtml(formatReportCurrency(entry.refundsPaid || 0))}`, tone: 'settlement-accent-negative', hide: (entry.refundsPaid || 0) < 0.005 },
      { label: 'Udział w kosztach', valueHtml: `- ${escapeReportHtml(formatReportCurrency(entry.costShareApplied || 0))}`, tone: 'settlement-accent-negative', hide: (entry.costShareApplied || 0) < 0.005 },
      { label: 'Zaliczki', valueHtml: `-${escapeReportHtml(formatReportCurrency(entry.advancesTaken || 0))}`, tone: 'settlement-accent-negative', hide: (entry.advancesTaken || 0) < 0.005 },
      { label: 'Przychód (Brutto)', valueHtml: escapeReportHtml(formatReportCurrency(entry.toPayout || 0)), tone: 'settlement-accent-positive settlement-accent-bold' },
      { label: 'Pensje pracowników', valueHtml: escapeReportHtml(formatReportCurrency(entry.employeeSalaries || 0)), tone: 'settlement-accent-warning', hide: (entry.employeeSalaries || 0) < 0.005 },
      { label: 'Podatek i ZUS za pracowników', valueHtml: escapeReportHtml(formatReportCurrency(entry.employeeAccountingRefund || 0)), tone: 'settlement-accent-warning', hide: (entry.employeeAccountingRefund || 0) < 0.005 },
      { label: 'Przychód (Brutto) z Pensjami', valueHtml: escapeReportHtml(formatReportCurrency(entry.grossWithEmployeeSalaries || 0)), tone: 'settlement-accent-warning settlement-accent-bold', hide: (entry.employeeSalaries || 0) < 0.005 }
    );
  }

  return buildReportMetricListHtml(rows);
}

function buildReportPaymentListHtml(items = [], emptyMessage = 'Brak pozycji.') {
  if (!items.length) {
    return `<p class="settlement-detail-empty">${escapeReportHtml(emptyMessage)}</p>`;
  }

  return `
    <ul class="settlement-inline-list">
      ${items.map(item => `<li>${escapeReportHtml(item.label)}: <strong>${escapeReportHtml(formatReportCurrency(item.amount || 0))}</strong></li>`).join('')}
    </ul>
  `;
}

function buildReportInvoicePanelHtml(reportData) {
  const { invoiceEntry } = reportData;
  if (!invoiceEntry) {
    return '';
  }

  const rows = [
    { label: 'Wpływ z faktur', valueHtml: escapeReportHtml(formatReportCurrency(invoiceEntry.receivedAmount || 0)) },
    { label: 'Docelowy przychód z rozliczenia', valueHtml: escapeReportHtml(formatReportCurrency(invoiceEntry.revenueTargetAmount || 0)) },
    { label: 'Podatek od wystawionych faktur', valueHtml: escapeReportHtml(formatReportCurrency(invoiceEntry.actualInvoiceTaxAmount || 0)), tone: 'settlement-accent-negative' },
    { label: 'Docelowy ciężar podatku', valueHtml: escapeReportHtml(formatReportCurrency(invoiceEntry.targetTaxAmount || 0)), tone: 'settlement-accent-negative' },
    { label: 'Rzeczywisty ciężar podatku', valueHtml: escapeReportHtml(formatReportCurrency(invoiceEntry.actualTaxBurdenAmount || 0)), tone: 'settlement-accent-negative' },
    { label: 'Saldo po wyrównaniach', valueHtml: escapeReportHtml(formatReportCurrency(invoiceEntry.currentBalance || 0)), tone: (invoiceEntry.currentBalance || 0) >= 0 ? 'settlement-accent-positive' : 'settlement-accent-negative' }
  ];

  return `
    <div class="settlement-detail-card">
      <h3>Wystawione faktury i podatki</h3>
      ${buildReportMetricListHtml(rows)}
    </div>
  `;
}

function buildReportEqualizationPanelHtml(reportData) {
  const { invoiceEntry } = reportData;
  if (!invoiceEntry) {
    return '';
  }

  const outgoing = [
    ...(invoiceEntry.transferPayments || []).map(item => ({ label: `${item.type || 'Przelew'} → ${item.recipientName}`, amount: item.amount || 0 })),
    ...(invoiceEntry.salaryPayments || [])
      .filter(item => (parseFloat(item.amount) || 0) >= 0.005)
      .map(item => ({ label: `${item.type || 'Wypłata'} → ${item.recipientName}`, amount: item.amount || 0 })),
    ...(invoiceEntry.officePayments || []).map(item => ({ label: item.type || 'Płatność do urzędu', amount: item.amount || 0 }))
  ];

  const incoming = [
    ...(invoiceEntry.incomingTransferPayments || []).map(item => ({ label: `${item.type || 'Wpływ'} ← ${item.payerName}`, amount: item.amount || 0 })),
    ...(invoiceEntry.salaryPayments || [])
      .filter(item => (parseFloat(item.amount) || 0) <= -0.005)
      .map(item => ({ label: `Do odebrania od ${item.recipientName}`, amount: Math.abs(item.amount || 0) }))
  ];

  return `
    <div class="settlement-detail-card">
      <h3>Wyrównanie i skrót zwrotów</h3>
      <div class="reports-two-column-grid">
        <div>
          <h4>Do przekazania</h4>
          ${buildReportPaymentListHtml(outgoing, 'Brak płatności wychodzących.')}
        </div>
        <div>
          <h4>Do otrzymania</h4>
          ${buildReportPaymentListHtml(incoming, 'Brak wpływów do otrzymania.')}
        </div>
      </div>
    </div>
  `;
}

function buildReportHoursTableHtml(hoursData, options = {}) {
  const visibleRows = Array.isArray(options.rows) ? options.rows : (hoursData?.rows || []);
  const includeFooter = options.includeFooter !== false;

  if (!hoursData?.hasData || visibleRows.length === 0) {
    return '<p class="settlement-detail-empty">Brak godzin, kosztów lub zaliczek dla tej osoby w wybranym miesiącu.</p>';
  }

  // Handle optional chunking of columns horizontally
  let visibleCols = hoursData.columns || [];
  let visibleExpCols = hoursData.expenseColumns || [];
  let sheetOffset = 0;
  let sheetCount = visibleCols.length;

  if (options.colChunk) {
    visibleCols = options.colChunk.columns;
    visibleExpCols = options.colChunk.expenseColumns;
    sheetOffset = options.colChunk.sheetOffset;
    sheetCount = visibleCols.length;
  }

  return `
    <div class="reports-table-container">
      <table class="reports-detail-table reports-hours-table">
        <thead>
          <tr>
            <th rowspan="2">Dzień</th>
            ${visibleCols.map((column, idx) => `<th colspan="2" class="${idx > 0 ? 'reports-td-sheet-boundary' : ''}">${escapeReportHtml(column.client)}<br><span class="reports-th-subtitle">${escapeReportHtml(column.site)}</span></th>`).join('')}
            ${visibleExpCols.map((column, idx) => `<th rowspan="2" class="${idx === 0 && visibleCols.length > 0 ? 'reports-td-sheet-boundary' : ''}">${escapeReportHtml(column.label)}</th>`).join('')}
          </tr>
          ${visibleCols.length > 0 ? `
          <tr>
            ${visibleCols.map((_, idx) => `<th class="${idx > 0 ? 'reports-td-sheet-boundary' : ''}">Czas</th><th>h</th>`).join('')}
          </tr>
          ` : ''}
        </thead>
        <tbody>
          ${visibleRows.map(row => {
            let rowClass = [];
            if (row.isHoliday) rowClass.push('is-holiday');
            if (row.isSaturday) rowClass.push('is-saturday');
            
            // Extract only the sheet values corresponding to the current chunk
            const rowSheetValues = (row.sheetValues || []).slice(sheetOffset, sheetOffset + sheetCount);
            
            return `
            <tr class="${rowClass.join(' ')}">
              <td class="reports-td-day">${row.day}<br><span class="reports-td-weekday">${escapeReportHtml(row.weekdayLabel)}</span></td>
              ${rowSheetValues.map((value, idx) => `
                <td class="reports-td-time ${idx > 0 ? 'reports-td-sheet-boundary' : ''}">${value.timeLabel ? escapeReportHtml(value.timeLabel) : ''}</td>
                <td class="reports-td-hours-centered">${formatReportHoursCell(value.hours)}</td>
              `).join('')}
              ${visibleExpCols.map((value, idx) => {
                 // For Expense values, they are mapped sequentially. 
                 // If we match exactly the same structure, wait... visibleExpCols is not 'value', it's columns.
                 // Need to index row.expenseValues
                 // Assuming expenseValues corresponds to hoursData.expenseColumns...
                 // If we chunk expense columns, we need to know their original index to fetch from row.expenseValues.
                 // But wait, expense columns are always appended as a whole chunk right now in our plan, so we can just use the indices if we pass them, or pass them all.
                 // To make it safe, let's just assume visibleExpCols are all expenses if isLastChunk.
                 const expValIdx = hoursData.expenseColumns.findIndex(ec => ec.label === value.label);
                 const expVal = row.expenseValues[expValIdx] || 0;
                 return `<td class="reports-td-amount-centered ${idx === 0 && visibleCols.length > 0 ? 'reports-td-sheet-boundary' : ''}">${formatReportAmountCell(expVal)}</td>`;
              }).join('')}
            </tr>
            `;
          }).join('')}
        </tbody>
        ${includeFooter ? `
        <tfoot>
          <tr>
            <th class="reports-td-total-label">Razem</th>
            ${visibleCols.map((column, idx) => `<td class="${idx > 0 ? 'reports-td-sheet-boundary' : ''}"></td><td class="reports-td-hours-centered">${escapeReportHtml(formatReportHours(column.totalHours || 0))}</td>`).join('')}
            ${visibleExpCols.map((column, idx) => `<td class="reports-td-amount-centered ${idx === 0 && visibleCols.length > 0 ? 'reports-td-sheet-boundary' : ''}">${formatReportAmountCell(column.total || 0)}</td>`).join('')}
          </tr>
        </tfoot>
        ` : ''}
      </table>
    </div>
  `;
}

function chunkReportHoursRows(rows = [], chunkSize = REPORTS_HOURS_ROWS_PER_PAGE) {
  const normalizedRows = Array.isArray(rows) ? rows : [];
  const normalizedChunkSize = Math.max(1, parseInt(chunkSize, 10) || REPORTS_HOURS_ROWS_PER_PAGE);

  if (normalizedRows.length === 0) {
    return [[]];
  }

  const chunks = [];
  for (let index = 0; index < normalizedRows.length; index += normalizedChunkSize) {
    chunks.push(normalizedRows.slice(index, index + normalizedChunkSize));
  }

  return chunks;
}

function buildReportWorksTableHtml(worksData) {
  if (!worksData?.hasData) {
    return '<p class="settlement-detail-empty">Brak pozycji z wykonanych prac dla tej osoby w wybranym miesiącu.</p>';
  }

  return worksData.sections.map(section => `
    <div class="settlement-detail-card" style="margin-top: 1rem;">
      <div class="settlement-detail-item-header">
        <div>
          <h4>${escapeReportHtml(section.client)}</h4>
          <div class="settlement-detail-meta">${escapeReportHtml(section.site)}</div>
        </div>
        <strong>${escapeReportHtml(formatReportHours(section.totalHours || 0))}</strong>
      </div>
      <div class="reports-table-container reports-table-container--fit">
        <table class="reports-detail-table reports-fit-table">
          <thead>
            <tr>
              <th>Data</th>
              <th>Pozycja</th>
              <th>Udział osoby</th>
              <th>Cena</th>
              <th>Wartość pozycji</th>
              <th>Wartość udziału</th>
            </tr>
          </thead>
          <tbody>
            ${section.items.map(item => `
              <tr>
                <td>${escapeReportHtml(item.date || '—')}</td>
                <td>${escapeReportHtml(item.name)}</td>
                <td>${escapeReportHtml(item.quantityText)}</td>
                <td>${escapeReportHtml(item.priceText)}</td>
                <td>${escapeReportHtml(item.totalValueText)}</td>
                <td>${escapeReportHtml(item.personValueText)}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </div>
  `).join('');
}

function buildReportPayrollTableHtml(payrollData, emptyMessage) {
  if (!payrollData?.hasData) {
    return `<p class="settlement-detail-empty">${escapeReportHtml(emptyMessage)}</p>`;
  }

  return `
    <div class="reports-table-container reports-table-container--fit">
      <table class="reports-detail-table reports-fit-table">
        <thead>
          <tr>
            <th>Osoba</th>
            <th>Rodzaj współpracy</th>
            <th>Godziny</th>
            <th>Podstawa</th>
            <th>Premie / diety</th>
            <th>Podatki i ZUS</th>
            <th>Do wypłaty</th>
          </tr>
        </thead>
        <tbody>
          ${payrollData.items.map(item => `
            <tr>
              <td>${escapeReportHtml(getPersonDisplayName(item.person))}${item.note ? `<div class="settlement-detail-meta">${escapeReportHtml(item.note)}</div>` : ''}</td>
              <td>${escapeReportHtml(item.typeLabel)}</td>
              <td>${escapeReportHtml(formatReportHours(item.hours || 0))}</td>
              <td>${escapeReportHtml(formatReportCurrency(item.salary || 0))}</td>
              <td>${escapeReportHtml(formatReportCurrency(item.benefits || 0))}</td>
              <td>${escapeReportHtml(formatReportCurrency((item.taxAmount || 0) + (item.zusAmount || 0)))}</td>
              <td>${escapeReportHtml(formatReportCurrency(item.toPayout || 0))}</td>
            </tr>
          `).join('')}
        </tbody>
        <tfoot>
          <tr>
            <th colspan="3">Razem</th>
            <th>${escapeReportHtml(formatReportCurrency(payrollData.totals.salary || 0))}</th>
            <th>${escapeReportHtml(formatReportCurrency(payrollData.totals.benefits || 0))}</th>
            <th>${escapeReportHtml(formatReportCurrency((payrollData.totals.taxes || 0) + (payrollData.totals.zus || 0)))}</th>
            <th>${escapeReportHtml(formatReportCurrency(payrollData.totals.payouts || 0))}</th>
          </tr>
        </tfoot>
      </table>
    </div>
  `;
}

function buildReportPagesHtml(reportData, options = {}) {
  const userScale = options.contentScale || 1.0;
  
  // SMART A4 CHUNKING (v11) - Orientation aware
  const orientation = options.printOrientation || 'portrait';
  const pageMetrics = getReportsPageMetrics(orientation);
  const PAGE_HEIGHT_AVAIL = pageMetrics.contentHeightMm; // mm
  const BASE_ROW_HEIGHT = 13; // mm natural (pre-zoom) - 2-line day cell: 2×(14px×1.1 line-height) + 2×9.6px padding + 2px border ≈ 50px ≈ 13mm
  const HOURS_NATURAL_HEADER_MM = 40; // mm natural height for page title + table thead rows + margins (pre-zoom)
  const HOURS_DAY_COLUMN_WIDTH_MM = 16;
  const HOURS_SHEET_PAIR_WIDTH_MM = 34;
  const HOURS_EXPENSE_COLUMN_WIDTH_MM = 24;

  let effectiveUserScale = userScale;
  
  // Create Column Chunks (Horizontal chunking)
  const maxCols = orientation === 'landscape' ? 7 : 4;
  const colChunks = [];
  const totalSheets = reportData.hoursData?.columns?.length || 0;
  
  if (totalSheets > 0) {
    for (let i = 0; i < totalSheets; i += maxCols) {
      const isLast = (i + maxCols) >= totalSheets;
      colChunks.push({
        columns: reportData.hoursData.columns.slice(i, i + maxCols),
        expenseColumns: isLast ? (reportData.hoursData.expenseColumns || []) : [],
        sheetOffset: i,
        isLastChunk: isLast
      });
    }
  } else if (reportData.hoursData?.expenseColumns?.length > 0) {
    colChunks.push({
      columns: [],
      expenseColumns: reportData.hoursData.expenseColumns,
      sheetOffset: 0,
      isLastChunk: true
    });
  } else {
    colChunks.push({ columns: [], expenseColumns: [], sheetOffset: 0, isLastChunk: true });
  }

  // Auto-scale calculation
  if (userScale === 'auto') {
    const allRowsCount = reportData.hoursData?.rows?.length || 0;
    // +1 accounts for the tfoot "Razem" row on the last page
    // ÷0.97 adds a 3% safety margin so the same formula used for HOURS_PAGE_ROWS guarantees fit
    const maxScaleByHeight = allRowsCount > 0
      ? PAGE_HEIGHT_AVAIL / (0.7 * ((allRowsCount + 1) * BASE_ROW_HEIGHT + HOURS_NATURAL_HEADER_MM) / 0.97)
      : 1;
    const maxScaleByWidth = colChunks.length > 0
      ? Math.min(...colChunks.map(colChunk => {
          const naturalWidthMm = HOURS_DAY_COLUMN_WIDTH_MM
            + (colChunk.columns.length * HOURS_SHEET_PAIR_WIDTH_MM)
            + (colChunk.expenseColumns.length * HOURS_EXPENSE_COLUMN_WIDTH_MM);
          return naturalWidthMm > 0 ? (pageMetrics.contentWidthMm / naturalWidthMm) : 1;
        }))
      : 1;

    effectiveUserScale = Math.max(0.5, Math.min(1, maxScaleByHeight, maxScaleByWidth));
  }

  const physicalScale = effectiveUserScale * 0.7;
  const contentScaleStyle = `zoom: ${physicalScale}; width: 100%;`;
  const tablePageScaleStyle = `zoom: ${physicalScale}; width: ${(100 / physicalScale).toFixed(2)}%;`;
  
  const wrapperOpen = `<div class="reports-preview-document-content" style="${contentScaleStyle}">`;
  const tableWrapperOpen = `<div class="reports-preview-document-content" style="${tablePageScaleStyle}">`;
  const wrapperClose = `</div>`;

  const isWorkingPartner = reportData.person.type === 'WORKING_PARTNER';
  const isEmployee = reportData.person.type === 'EMPLOYEE';
  const isSeparateCompany = reportData.person.type === 'SEPARATE_COMPANY';

  // Work in natural (pre-zoom) mm: content zoomed at physicalScale fits PAGE_HEIGHT_AVAIL / physicalScale natural mm.
  // -1 reserves one slot for the tfoot "Razem" row on the last chunk; *0.97 is a 3% safety margin.
  const naturalPageHeight = PAGE_HEIGHT_AVAIL / physicalScale;
  const HOURS_PAGE_ROWS = Math.max(5, Math.floor((naturalPageHeight - HOURS_NATURAL_HEADER_MM) / BASE_ROW_HEIGHT * 0.97) - 1);

  const allRows = reportData.hoursData?.rows || [];
  const hoursRowChunks = [];
  if (allRows.length > 0) {
    for (let i = 0; i < allRows.length; i += HOURS_PAGE_ROWS) {
      hoursRowChunks.push(allRows.slice(i, i + HOURS_PAGE_ROWS));
    }
  }

  const summaryPreview = reportData.summaryRow ? `
    <div class="reports-summary-inline-table">
      <table class="reports-detail-table reports-fit-table">
        <thead>
          <tr>
            <th>Imię</th>
            <th>Rodzaj współpracy</th>
            <th>Godziny</th>
            <th>Szacowane wynagrodzenie</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>${escapeReportHtml(reportData.summaryRow.name)}</td>
            <td><span class="badge ${reportData.summaryRow.badgeClass}">${escapeReportHtml(reportData.summaryRow.label)}</span></td>
            <td>${escapeReportHtml(formatReportHours(reportData.summaryRow.hours || 0))}</td>
            <td>${getSummaryPreviewCompensationHtml(reportData.summaryRow)}</td>
          </tr>
        </tbody>
      </table>
    </div>
  ` : '<p class="settlement-detail-empty">Brak skrótu zarobków dla tej osoby.</p>';

  const payrollTitle = reportData.person.type === 'EMPLOYEE'
    ? 'Twoje rozliczenie i potrącenia'
    : 'Pensje pracowników i podatki do wypłaty';

  const payrollHtml = reportData.person.type === 'EMPLOYEE'
    ? buildReportSettlementSummaryHtml(reportData, options)
    : buildReportPayrollTableHtml(reportData.payrollData, 'Ta osoba nie ma przypisanych pracowników lub wspólników pracujących.');

  // Page 1 Content
  let page1Content = `
    <div class="reports-page-header">
      <div>
        <div class="summary-shortcut-kicker">Strona 1</div>
        <h2>${escapeReportHtml(getPersonDisplayName(reportData.person))}</h2>
        <p class="reports-person-subtitle">${escapeReportHtml(getReportPersonTypeLabel(reportData.person.type))} • ${escapeReportHtml(formatMonthLabel(reportData.month))}</p>
      </div>
    </div>
    ${summaryPreview}
    <div class="reports-two-column-grid reports-summary-grid">
      <div class="settlement-detail-card">
        <h3>Rozliczenie główne</h3>
        ${buildReportSettlementSummaryHtml(reportData, options)}
      </div>
      ${buildReportInvoicePanelHtml(reportData)}
    </div>
    <div class="reports-single-column-grid">
      ${buildReportEqualizationPanelHtml(reportData)}
    </div>
  `;

  // If Working Partner, append Payroll (Page 3 content) to Page 1
  if (isWorkingPartner) {
    page1Content += `
      <div class="settlement-detail-card" style="margin-top: 1.5rem;">
        <h3>${escapeReportHtml(payrollTitle)}</h3>
        ${payrollHtml}
      </div>
    `;
  }

  const pages = [];
  let pageNumber = 1;

  pages.push(`
    <section class="reports-person-page">
      ${wrapperOpen}
      ${page1Content}
      ${wrapperClose}
    </section>
  `.replace('Strona 1', `Strona ${pageNumber}`));
  pageNumber += 1;
  hoursRowChunks.forEach((rows, vIndex) => {
    colChunks.forEach((colChunk, hIndex) => {
      let pageTitle = 'Godziny i wykonane prace';
      if (vIndex > 0 || hIndex > 0) pageTitle += ' (cd.)';

      pages.push(`
        <section class="reports-person-page">
          ${tableWrapperOpen}
          <div class="reports-page-header">
            <div>
              <div class="summary-shortcut-kicker">Strona ${pageNumber}</div>
              <h2>${pageTitle}</h2>
            </div>
          </div>
          ${buildReportHoursTableHtml(reportData.hoursData, { 
            rows: rows, 
            includeFooter: vIndex === hoursRowChunks.length - 1,
            colChunk: colChunk
          })}
          ${wrapperClose}
        </section>
      `);
      pageNumber += 1;
    });
  });

  if (reportData.worksData && reportData.worksData.length > 0) {
    pages.push(`
      <section class="reports-person-page">
        ${wrapperOpen}
        <div class="reports-page-header">
          <div>
            <div class="summary-shortcut-kicker">Strona ${pageNumber}</div>
            <h2>Wykonane prace z udziałem osoby</h2>
          </div>
        </div>
        <div class="settlement-detail-card">
          ${buildReportWorksTableHtml(reportData.worksData)}
        </div>
        ${wrapperClose}
      </section>
    `);
    pageNumber += 1;
  }

  if (!isWorkingPartner && !isEmployee && reportData.payrollData?.hasData) {
    pages.push(`
      <section class="reports-person-page">
        ${wrapperOpen}
        <div class="reports-page-header">
          <div>
            <div class="summary-shortcut-kicker">Strona ${pageNumber}</div>
            <h2>${escapeReportHtml(payrollTitle)}</h2>
          </div>
        </div>
        <div class="settlement-detail-card">
          ${payrollHtml}
        </div>
        ${wrapperClose}
      </section>
    `);
    pageNumber += 1;
  }

  if (isSeparateCompany && reportData.payrollData?.hasData) {
    pages.push(`
      <section class="reports-person-page">
        ${wrapperOpen}
        <div class="reports-page-header">
          <div>
            <div class="summary-shortcut-kicker">Strona ${pageNumber}</div>
            <h2>Pracownicy osobnej firmy</h2>
          </div>
        </div>
        ${buildReportPayrollTableHtml(reportData.payrollData, 'Brak rozliczeń pracowników tej firmy.')}
        ${wrapperClose}
      </section>
    `);
  }

  return pages.join('');
}

function buildPersonReportCardHtml(reportData, options = {}) {
  return buildReportPagesHtml(reportData, {
    ...options
  });
}

function buildReportsOverviewHtml(settlement, month, personsCount) {
  return `
    <div class="grid-cards" style="margin-bottom: 1rem;">
      <div class="glass-panel stat-card">
        <span class="stat-label">Aktywne zestawienia</span>
        <span class="stat-value">${personsCount}</span>
      </div>
      <div class="glass-panel stat-card">
        <span class="stat-label">Godziny zespołu</span>
        <span class="stat-value">${escapeReportHtml(formatReportHours(settlement.totalTeamHours || 0))}</span>
      </div>
      <div class="glass-panel stat-card">
        <span class="stat-label">Przychód miesiąca</span>
        <span class="stat-value" style="color: var(--primary);">${escapeReportHtml(formatReportCurrency((settlement.commonRevenue || 0) - (settlement.clientAdvances || 0)))}</span>
      </div>
      <div class="glass-panel stat-card">
        <span class="stat-label">Zarobek do podziału</span>
        <span class="stat-value" style="color: var(--success);">${escapeReportHtml(formatReportCurrency(settlement.profitToSplit || 0))}</span>
      </div>
    </div>
    <p class="reports-note" style="margin-bottom: 1rem;">Zestawienia przygotowano dla ${personsCount} aktywnych osób w miesiącu ${escapeReportHtml(formatMonthLabel(month))}.</p>
  `;
}

function buildReportsPrintDocumentHtml(title, bodyHtml, options = {}) {
  const pageMetrics = getReportsPageMetrics(options.printOrientation || 'portrait');
  const _cs = String(options.contentScale ?? 'auto');
  const _soHtml = [
    ['1.5','150%'],['1.4','140%'],['1.3','130%'],['1.2','120%'],['1.1','110%'],
    ['auto','Auto'],['1','100%'],['0.9','90%'],['0.8','80%'],['0.7','70%'],['0.6','60%'],['0.5','50%']
  ].map(([v,l]) => '<option value="' + v + '"' + (v === _cs ? ' selected' : '') + '>' + l + '</option>').join('');
  const _svgScan = '<svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 7V5a2 2 0 0 1 2-2h2"/><path d="M17 3h2a2 2 0 0 1 2 2v2"/><path d="M21 17v2a2 2 0 0 1-2 2h-2"/><path d="M7 21H5a2 2 0 0 1-2-2v-2"/></svg>';
  const _svgMin2 = '<svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="4 14 10 14 10 20"/><polyline points="20 10 14 10 14 4"/><line x1="10" y1="14" x2="3" y2="21"/><line x1="21" y1="3" x2="14" y2="10"/></svg>';
  const _svgLR = '<svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="17 11 21 7 17 3"/><line x1="21" y1="7" x2="3" y2="7"/><polyline points="7 21 3 17 7 13"/><line x1="3" y1="17" x2="21" y2="17"/></svg>';
  const _svgPrint = '<svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>';
  const _svgX = '<svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>';
  return `<!DOCTYPE html>
<html lang="pl">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeReportHtml(title)}</title>
  <style>
    :root {
      --bg: #ffffff;
      --panel: #ffffff;
      --panel-soft: #f8fafc;
      --panel-muted: #eef2f7;
      --text: #1e293b;
      --text-muted: #64748b;
      --border: rgba(15, 23, 42, 0.12);
      --accent: #4f46e5;
      --accent-soft: #e0e7ff;
      --success: #047857;
      --danger: #b91c1c;
      --warning: #b45309;
      --preview-bg: #cbd5e1;
      --preview-page-width: ${pageMetrics.sheetWidthMm}mm;
      --preview-page-min-height: ${pageMetrics.sheetHeightMm}mm;
      --preview-page-padding: ${pageMetrics.marginMm}mm;
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }
    html,
    body {
      height: 100%;
      margin: 0;
      font-family: Inter, system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
      background: var(--bg);
      color: var(--text);
      line-height: 1.45;
      overflow: hidden;
    }

    .reports-preview-window {
      display: flex;
      flex-direction: column;
      height: 100%;
      overflow: hidden;
    }

    .reports-preview-window-toolbar {
      position: sticky;
      top: 0;
      z-index: 20;
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 0.35rem;
      padding: 0.4rem 0.65rem;
      background: rgba(248, 250, 252, 0.96);
      border-bottom: 1px solid var(--border);
      backdrop-filter: blur(10px);
      flex: 0 0 auto;
    }

    .rptw-g1, .rptw-g2, .rptw-g3, .rptw-g4 {
      display: flex; align-items: center; gap: 0.3rem;
    }
    .rptw-g3 { margin-left: auto; }
    .rptw-g4 { flex: 1 1 auto; min-width: 0; }

    .reports-preview-window-toolbar button {
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 0.3rem 0.4rem;
      cursor: pointer;
      background: var(--panel);
      color: var(--text);
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 0.3rem;
      font: inherit;
      font-size: 0.85rem;
      line-height: 1;
    }

    .reports-preview-window-toolbar button.primary {
      background: var(--accent);
      color: #fff;
      border-color: var(--accent);
      padding: 0.3rem 0.65rem;
    }

    .reports-preview-window-toolbar input[type="range"] {
      width: 110px;
      accent-color: var(--accent);
      font: inherit;
    }

    .rptw-scale-select {
      padding: 0.28rem 0.4rem;
      border: 1px solid var(--border);
      border-radius: 8px;
      background: var(--panel);
      color: var(--text);
      font: inherit;
      font-size: 0.82rem;
      cursor: pointer;
    }

    .reports-preview-window-zoom-value {
      min-width: 2.8rem;
      text-align: right;
      font-weight: 700;
      font-size: 0.85rem;
      color: var(--text);
    }

    @media (max-width: 600px) {
      .rptw-g4 { flex: 0 0 100%; order: 10; }
      .rptw-g4 input[type="range"] { flex: 1 1 0; min-width: 60px; width: auto; }
    }

    .reports-preview-window-viewport {
      flex: 1 1 auto;
      overflow: auto;
      padding: 1rem;
      background: linear-gradient(180deg, var(--preview-bg), #e2e8f0);
    }

    .reports-preview-window-stage {
      position: relative;
      width: max-content;
      height: max-content;
      min-width: 100%;
    }

    .reports-preview-window-document {
      position: relative;
      width: var(--preview-page-width);
      transform-origin: top left;
    }

    .reports-person-card--print {
      margin: 0 0 1rem 0;
      padding: 0;
      background: transparent;
      border: none;
      box-shadow: none;
    }

    table tr.is-holiday td {
      background: #fff7ed !important;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }

    table tr.is-saturday td {
      background: #fefce8 !important;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }

    .reports-preview-document-content {
      transform-origin: top left;
    }

    .reports-person-page {
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 16px;
      padding: var(--preview-page-padding);
      margin-bottom: 1rem;
      width: var(--preview-page-width);
      height: var(--preview-page-min-height);
      break-inside: avoid;
      page-break-inside: avoid;
      box-shadow: 0 20px 40px rgba(15, 23, 42, 0.14);
      box-sizing: border-box;
      overflow: hidden;
    }

    /* Critical: heading line-height must match main app for header height calculation */
    h1, h2, h3, h4, h5, h6 { line-height: 1.2; }
    h2 { margin-bottom: 0.5rem; }

    .reports-page-header,
    .reports-two-column-grid,
    .reports-single-column-grid,
    .settlement-detail-stack,
    .settlement-detail-row,
    .settlement-inline-list,
    .reports-card-toolbar,
    .grid-cards {
      display: block;
    }

    .reports-page-header { margin-bottom: 1rem; }

    .reports-two-column-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 1rem;
      margin-top: 1rem;
    }

    .reports-summary-grid,
    .reports-single-column-grid {
      margin-top: 1rem;
    }

    .reports-summary-inline-table {
      margin-top: 1rem;
    }

    .reports-table-container {
      margin-top: 0.85rem;
      width: fit-content;
      max-width: 100%;
      overflow: visible;
    }

    .reports-detail-table {
      width: max-content;
      min-width: 100%;
    }

    .reports-hours-table {
      width: auto !important;
      min-width: 0 !important;
      max-width: none !important;
      table-layout: auto !important;
      margin-top: 1rem !important;
    }

    /* Critical: day cell line-height MUST match BASE_ROW_HEIGHT=13mm assumption in buildReportPagesHtml */
    .reports-td-day {
      text-align: center !important;
      font-weight: 700 !important;
      line-height: 1.1 !important;
      vertical-align: middle !important;
    }
    .reports-td-weekday {
      font-size: 0.65rem !important;
      font-weight: 500 !important;
      text-transform: lowercase !important;
    }

    .reports-fit-table {
      width: 100% !important;
      min-width: 0 !important;
      max-width: 100% !important;
      table-layout: fixed;
    }

    .reports-detail-table th,
    .reports-detail-table td {
      padding: 0.6rem 0.55rem;
    }

    .settlement-detail-card,
    .glass-panel,
    .stat-card {
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 14px;
      padding: 1rem;
      margin-bottom: 1rem;
      box-shadow: none;
    }

    .grid-cards {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 0.75rem;
    }

    .summary-shortcut-kicker,
    .settlement-detail-meta,
    .reports-person-subtitle,
    .reports-note,
    .stat-label {
      color: var(--text-muted);
    }
    
    .reports-th-subtitle {
      display: inline-block;
      margin-top: 0.2rem;
      color: var(--text-muted);
      font-size: 0.68rem;
      font-weight: 500;
      text-transform: none;
    }
    
    .reports-empty-cell {
      color: var(--text-muted);
      opacity: 0.4;
    }
    
    .reports-hours-table th[colspan="2"] {
      text-align: center;
    }

    .summary-shortcut-kicker {
      display: block;
      margin-bottom: 0.2rem;
      font-size: 0.76rem;
      font-weight: 700;
      letter-spacing: 0.04em;
      text-transform: uppercase;
      color: var(--accent);
    }

    .stat-value,
    h1,
    h2,
    h3,
    h4,
    strong {
      color: var(--text);
    }

    .settlement-detail-row {
      display: flex;
      justify-content: space-between;
      gap: 1rem;
      padding: 0.25rem 0;
    }

    .settlement-inline-list {
      margin: 0.6rem 0 0 1rem;
      padding: 0;
    }

    .settlement-inline-list li + li {
      margin-top: 0.25rem;
    }

    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 0.88rem;
    }

    th,
    td {
      border: 1px solid var(--border);
      padding: 0.55rem 0.6rem;
      vertical-align: top;
      text-align: left;
    }

    th {
      background: var(--panel-muted);
      font-size: 0.78rem;
      text-transform: uppercase;
      letter-spacing: 0.04em;
    }

    .badge {
      display: inline-flex;
      align-items: center;
      padding: 0.2rem 0.55rem;
      border-radius: 999px;
      font-size: 0.74rem;
      font-weight: 700;
      border: 1px solid var(--border);
      background: var(--accent-soft);
      color: var(--accent);
    }

    .badge-employee { background: rgba(16, 185, 129, 0.12); color: var(--success); }
    .badge-working-partner { background: rgba(124, 58, 237, 0.12); color: #6d28d9; }
    .badge-partner { background: var(--accent-soft); color: var(--accent); }

    .summary-preview-compensation-row {
      display: flex;
      justify-content: space-between;
      gap: 0.75rem;
    }

    .summary-preview-compensation-row strong,
    .settlement-accent-positive { color: var(--success); }
    .settlement-accent-negative { color: var(--danger); }
    .settlement-accent-primary { color: var(--accent); }

    @page {
      size: ${options.printOrientation === 'landscape' ? 'A4 landscape' : 'A4'};
      margin: 0;
    }

    @media print {
      html,
      body {
        height: auto;
        overflow: visible;
        background: #fff;
      }

      .reports-preview-window-toolbar {
        display: none;
      }

      .reports-preview-window-viewport {
        overflow: visible;
        padding: 0;
        background: #fff;
      }

      .reports-preview-window-stage {
        width: auto !important;
        height: auto !important;
        min-width: 0;
      }

      .reports-preview-window-document {
        transform: none !important;
        width: var(--preview-page-width);
      }

      .reports-person-card--print { border: none; padding: 0; margin: 0; }
      .reports-person-page {
        break-after: page;
        page-break-after: always;
        width: var(--preview-page-width);
        height: var(--preview-page-min-height);
        padding: var(--preview-page-padding);
        margin-left: auto;
        margin-right: auto;
        margin-bottom: 0;
        border: none;
        border-radius: 0;
        box-shadow: none;
        background: #fff;
        box-sizing: border-box;
        overflow: hidden;
      }
      .reports-person-page:last-child { break-after: auto; page-break-after: auto; }

      table tr.is-holiday td {
        background-color: #fff7ed !important;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }

      table tr.is-saturday td {
        background-color: #fefce8 !important;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }
    }
  </style>
</head>
<body>
  <div class="reports-preview-window">
    <div class="reports-preview-window-toolbar">
      <div class="rptw-g1">
        <button type="button" id="btn-reports-preview-zoom-actual" title="100%">${_svgScan}</button>
        <button type="button" id="btn-reports-preview-zoom-fit-page" title="Cała strona">${_svgMin2}</button>
        <button type="button" id="btn-reports-preview-zoom-fit-width" title="Na szerokość">${_svgLR}</button>
      </div>
      <div class="rptw-g2">
        <select id="reports-preview-scale" class="rptw-scale-select" title="Skala danych">${_soHtml}</select>
      </div>
      <div class="rptw-g3">
        <button type="button" class="primary" id="btn-reports-preview-print" title="Drukuj / zapisz PDF">${_svgPrint} Drukuj</button>
        <button type="button" id="btn-reports-preview-close" title="Zamknij">${_svgX}</button>
      </div>
      <div class="rptw-g4">
        <input type="range" id="reports-preview-zoom-range" min="${REPORTS_PREVIEW_ZOOM_MIN}" max="${REPORTS_PREVIEW_ZOOM_MAX}" step="${REPORTS_PREVIEW_ZOOM_STEP}" value="100">
        <span id="reports-preview-zoom-value" class="reports-preview-window-zoom-value">100%</span>
      </div>
    </div>
    <div id="reports-preview-viewport" class="reports-preview-window-viewport">
      <div id="reports-preview-stage" class="reports-preview-window-stage">
        <div id="reports-preview-document" class="reports-preview-window-document">
          ${bodyHtml}
        </div>
      </div>
    </div>
  </div>
  <script>
    (function () {
      const shouldAutoPrint = ${options.autoPrint === true ? 'true' : 'false'};
      let autoPrintTriggered = false;

      const clampZoom = (value) => {
        const parsed = parseFloat(value);
        if (!Number.isFinite(parsed)) return 100;
        const clamped = Math.max(${REPORTS_PREVIEW_ZOOM_MIN}, Math.min(${REPORTS_PREVIEW_ZOOM_MAX}, parsed));
        return Math.round(clamped / ${REPORTS_PREVIEW_ZOOM_STEP}) * ${REPORTS_PREVIEW_ZOOM_STEP};
      };

      const state = {
        mode: 'fit-width',
        zoomPercent: 100
      };

      const viewport = document.getElementById('reports-preview-viewport');
      const stage = document.getElementById('reports-preview-stage');
      const documentEl = document.getElementById('reports-preview-document');
      const range = document.getElementById('reports-preview-zoom-range');
      const value = document.getElementById('reports-preview-zoom-value');

      const getMetrics = () => {
        const documentWidth = documentEl.offsetWidth || documentEl.scrollWidth || 0;
        const documentHeight = documentEl.scrollHeight || documentEl.offsetHeight || 0;
        if (!(documentWidth > 0) || !(documentHeight > 0)) return null;
        return {
          documentWidth,
          documentHeight,
          viewportWidth: viewport.clientWidth || 0,
          viewportHeight: viewport.clientHeight || 0
        };
      };

      const getResolvedZoom = () => {
        const metrics = getMetrics();
        if (!metrics) return clampZoom(state.zoomPercent);
        const fitWidthZoom = clampZoom((metrics.viewportWidth / metrics.documentWidth) * 100);
        const fitPageZoom = clampZoom(Math.min(
          (metrics.viewportWidth / metrics.documentWidth) * 100,
          (metrics.viewportHeight / metrics.documentHeight) * 100
        ));

        if (state.mode === 'fit-width') return fitWidthZoom;
        if (state.mode === 'fit-page') return fitPageZoom;
        return clampZoom(state.zoomPercent);
      };

      const applyZoom = () => {
        const metrics = getMetrics();
        if (!metrics) return;
        const resolvedZoom = getResolvedZoom();
        const scale = resolvedZoom / 100;

        documentEl.style.transform = 'scale(' + scale + ')';
        stage.style.width = Math.ceil(metrics.documentWidth * scale) + 'px';
        stage.style.height = Math.ceil(metrics.documentHeight * scale) + 'px';
        range.value = String(resolvedZoom);
        value.textContent = resolvedZoom + '%';
      };

      const triggerAutoPrint = () => {
        if (!shouldAutoPrint || autoPrintTriggered) return;
        autoPrintTriggered = true;

        const startPrint = () => {
          window.focus();
          window.setTimeout(() => {
            window.print();
          }, 180);
        };

        if (document.fonts?.ready && typeof document.fonts.ready.then === 'function') {
          document.fonts.ready
            .then(() => {
              window.setTimeout(startPrint, 120);
            })
            .catch(() => {
              window.setTimeout(startPrint, 120);
            });
          return;
        }

        window.setTimeout(startPrint, 180);
      };

      document.getElementById('btn-reports-preview-zoom-actual').addEventListener('click', function () {
        state.mode = 'manual';
        state.zoomPercent = 100;
        applyZoom();
      });

      document.getElementById('btn-reports-preview-zoom-fit-page').addEventListener('click', function () {
        state.mode = 'fit-page';
        applyZoom();
      });

      document.getElementById('btn-reports-preview-zoom-fit-width').addEventListener('click', function () {
        state.mode = 'fit-width';
        applyZoom();
      });

      range.addEventListener('input', function () {
        state.mode = 'manual';
        state.zoomPercent = clampZoom(range.value);
        applyZoom();
      });

      document.getElementById('btn-reports-preview-print').addEventListener('click', function () {
        window.print();
      });

      document.getElementById('btn-reports-preview-close').addEventListener('click', function () {
        window.close();
      });

      document.getElementById('reports-preview-scale').addEventListener('change', function () {
        var newScale = this.value;
        try {
          if (window.opener && !window.opener.closed) {
            if (typeof window.opener.getReportsViewState === 'function') {
              var vs = window.opener.getReportsViewState();
              vs.printContentScale = newScale === 'auto' ? 'auto' : parseFloat(newScale);
            }
            var mainSel = window.opener.document.getElementById('reports-content-scale');
            if (mainSel) mainSel.value = newScale;
            if (typeof window.opener.openReportsPrintPreview === 'function') {
              window.opener.openReportsPrintPreview({ autoPrint: false });
            }
          }
        } catch (e) {}
        window.close();
      });

      window.addEventListener('resize', function () {
        if (state.mode !== 'manual') {
          applyZoom();
          return;
        }
        applyZoom();
      });

      applyZoom();

      if (shouldAutoPrint) {
        window.requestAnimationFrame(() => {
          window.requestAnimationFrame(triggerAutoPrint);
        });
      }
    }());
  </script>
</body>
</html>`;
}

function openReportsPrintWindow(title, bodyHtml, options = {}) {
  const printWindow = window.open('about:blank', '_blank');
  if (!printWindow) return null;

  printWindow.document.open();
  printWindow.document.write(buildReportsPrintDocumentHtml(title, bodyHtml, options));
  printWindow.document.close();
  return printWindow;
}

function shouldUseDedicatedReportsPrintPreview() {
  return /android/i.test(navigator.userAgent || '');
}

function openReportsPrintPreview(options = {}) {
  const selectionData = getReportsSelectedPersonsData(options);
  const { selectedPersons = [] } = selectionData;

  if (!selectedPersons.length) {
    alert('Brak aktywnych osób do przygotowania zestawienia.');
    return false;
  }

  const title = buildReportsPreviewTitle(selectionData);
  const printWindow = openReportsPrintWindow(title, buildReportsPreviewBodyHtml(selectionData), {
    printOrientation: selectionData.viewState.printOrientation,
    contentScale: selectionData.viewState.printContentScale ?? 'auto',
    autoPrint: options.autoPrint === true
  });

  if (!printWindow) {
    alert('Nie udało się otworzyć okna podglądu. Sprawdź blokowanie wyskakujących okienek.');
    return false;
  }

  if (options.autoPrint === true) {
    try {
      printWindow.focus();
    } catch {
      // Ignore focus errors in restricted mobile browsers.
    }
  }

  return true;
}

function initReportsView() {
  const select = document.getElementById('reports-person-filter');
  const printButton = document.getElementById('btn-reports-print');
  if (!select || !printButton || select.dataset.bound === 'true') return;
  select.dataset.bound = 'true';

  // INITIAL POPULATION of person filter select
  const selectionData = getReportsSelectedPersonsData();
  const { activePersons, viewState } = selectionData;
  viewState.personId = syncReportsPersonFilterOptions(select, activePersons, viewState.personId);

  // Person filter
  select.addEventListener('change', () => {
    const personId = select.value || 'all';
    const vs = getReportsViewState();
    vs.personId = personId;
    
    // AUTO-SWITCH VISIBILITY:
    // If selecting an Employee -> Default to 'Dla pracownika' (Worker View)
    // If selecting Partner/Separate Company/All -> Default to 'Pełne dane' (Internal View)
    const currentData = getReportsSelectedPersonsData();
    if (personId !== 'all') {
      const selectedPerson = (currentData.activePersons || []).find(p => p.id === personId);
      if (selectedPerson && selectedPerson.type === 'EMPLOYEE') {
        vs.employeeVisibility = 'employee';
      } else {
        vs.employeeVisibility = 'internal';
      }
    } else {
      vs.employeeVisibility = 'internal';
    }

    renderReports();
  });

  // Employee visibility radio
  document.querySelectorAll('input[name="reports-employee-visibility"]').forEach(input => {
    input.addEventListener('change', () => {
      if (!input.checked) return;
      getReportsViewState().employeeVisibility = input.value;
      renderReports();
    });
  });

  // Orientation
  const orientSel = document.getElementById('reports-orientation');
  orientSel?.addEventListener('change', () => {
    getReportsViewState().printOrientation = orientSel.value;
    renderReports();
  });

  // Content scale
  const scaleSel = document.getElementById('reports-content-scale');
  scaleSel?.addEventListener('change', () => {
    let val = scaleSel.value;
    if (val !== 'auto') val = parseFloat(val) || 1.0;
    getReportsViewState().printContentScale = val;
    renderReports();
  });

  // Zoom controls
  const zoomRange = document.getElementById('reports-zoom-range');
  const zoomVal = document.getElementById('reports-zoom-val');
  const viewport = document.getElementById('reports-viewport');

  function applyZoom(pct, mode = 'manual') {
    const vs = getReportsViewState();
    pct = Math.max(30, Math.min(200, Math.round(pct / 5) * 5));
    vs.previewZoom = pct;
    vs.previewZoomMode = mode;
    if (zoomRange) zoomRange.value = pct;
    if (zoomVal) zoomVal.textContent = pct + '%';
    const docEl = document.getElementById('reports-document');
    const stageEl = document.getElementById('reports-stage');
    if (docEl) {
      const scale = pct / 100;
      docEl.style.transform = `scale(${scale})`;
      if (stageEl) {
        stageEl.style.width = (docEl.scrollWidth * scale) + 'px';
        stageEl.style.height = (docEl.scrollHeight * scale) + 'px';
      }
    }
  }

  const applyFitPage = () => {
    if (!viewport) return;
    const docEl = document.getElementById('reports-document');
    if (!docEl) return;
    docEl.style.transform = 'none';
    const page = docEl.querySelector('.reports-person-page');
    if (!page) return;
    const pageH = page.offsetHeight;
    const availH = viewport.clientHeight - 32;
    const availW = viewport.clientWidth - 32;
    const fitPct = Math.min(availW / page.offsetWidth, availH / pageH) * 100;
    applyZoom(fitPct, 'fit-page');
  };

  const applyFitWidth = () => {
    if (!viewport) return;
    const docEl = document.getElementById('reports-document');
    if (!docEl) return;
    docEl.style.transform = 'none';
    const page = docEl.querySelector('.reports-person-page');
    if (!page) return;
    const availW = viewport.clientWidth - 32;
    const fitPct = (availW / page.offsetWidth) * 100;
    applyZoom(fitPct, 'fit-width');
  };

  window.__reportsApplyZoom = applyZoom;
  window.__reportsApplyFitWidth = applyFitWidth;
  window.__reportsApplyFitPage = applyFitPage;

  zoomRange?.addEventListener('input', () => applyZoom(parseInt(zoomRange.value)));

  document.getElementById('btn-rpt-zoom-100')?.addEventListener('click', () => applyZoom(100));

  document.getElementById('btn-rpt-zoom-fit')?.addEventListener('click', applyFitPage);

  document.getElementById('btn-rpt-zoom-width')?.addEventListener('click', applyFitWidth);

  // Resize listener for active "Fit" modes
  window.addEventListener('resize', () => {
    const vs = getReportsViewState();
    if (vs.previewZoomMode === 'fit-width' && window.__reportsApplyFitWidth) {
      window.__reportsApplyFitWidth();
    } else if (vs.previewZoomMode === 'fit-page' && window.__reportsApplyFitPage) {
      window.__reportsApplyFitPage();
    }
  });

  // Print button — uses window.print() directly
  printButton.addEventListener('click', () => {
    if (shouldUseDedicatedReportsPrintPreview()) {
      const opened = openReportsPrintPreview({ autoPrint: true });
      if (opened) return;
    }

    const selectionData = getReportsSelectedPersonsData();
    if (!selectionData.selectedPersons?.length) {
      alert('Brak aktywnych osób do przygotowania zestawienia.');
      return;
    }
    const originalTitle = document.title;
    document.title = buildReportsPreviewTitle(selectionData);

    const docEl = document.getElementById('reports-document');
    const stageEl = document.getElementById('reports-stage');
    const savedTransform = docEl?.style.transform || '';
    const savedStageW = stageEl?.style.width || '';
    const savedStageH = stageEl?.style.height || '';
    if (docEl) docEl.style.transform = 'none';
    if (stageEl) { stageEl.style.width = ''; stageEl.style.height = ''; }

    document.body.classList.add('printing-reports');
    const orientation = getReportsViewState().printOrientation;
    if (orientation === 'landscape') document.body.classList.add('print-landscape');

    const styleEl = document.createElement('style');
    styleEl.id = 'dynamic-print-orientation';
    styleEl.innerHTML = `@media print { @page { size: A4 ${orientation}; margin: ${REPORTS_PRINT_PAGE_MARGIN_MM}mm; } }`;
    document.head.appendChild(styleEl);
    
    const cleanupPrint = () => {
      document.body.classList.remove('printing-reports');
      document.body.classList.remove('print-landscape');
      document.title = originalTitle;
      const dynStyle = document.getElementById('dynamic-print-orientation');
      if (dynStyle) dynStyle.remove();
      if (docEl) docEl.style.transform = savedTransform;
      if (stageEl) { stageEl.style.width = savedStageW; stageEl.style.height = savedStageH; }
      window.removeEventListener('afterprint', cleanupPrint);
    };

    window.addEventListener('afterprint', cleanupPrint);

    // Short delay gives mobile browsers (Android/Chrome) time to repaint
    // the UI with 'printing-reports' class before locking into the PDF/Print dialog
    setTimeout(() => {
      window.print();
      
      // Fallback for browsers/WebViews where afterprint might not fire reliably
      const isSafari = /^((?!chrome|android).)*safari/i.test(navigator.userAgent);
      if (isSafari || /android/i.test(navigator.userAgent)) {
         setTimeout(cleanupPrint, 3000); // Failsafe cleanup
      }
    }, 150);
  });

  // Toggle extra options on mobile
  document.getElementById('btn-reports-expand')?.addEventListener('click', () => {
    const toolbar = document.querySelector('.reports-toolbar');
    if (toolbar) {
      toolbar.classList.toggle('is-expanded');
    }
  });

  // Store applyZoom for external use
  window.__reportsApplyZoom = applyZoom;

  // Initialize newly added icons
  if (typeof lucide !== 'undefined' && lucide.createIcons) {
    lucide.createIcons();
  }
}

function renderReports() {
  const selectionData = getReportsSelectedPersonsData();
  const { state, month, settlement, invoiceReconciliation, activePersons, viewState, selectedPersons } = selectionData;
  const select = document.getElementById('reports-person-filter');
  const subtitle = document.getElementById('reports-subtitle');
  if (!select) return;

  const orientSel = document.getElementById('reports-orientation');
  const scaleSel = document.getElementById('reports-content-scale');
  const viewport = document.getElementById('reports-viewport');
  const docEl = document.getElementById('reports-document');
  const stageEl = document.getElementById('reports-stage');

  if (orientSel) orientSel.value = viewState.printOrientation;
  if (scaleSel) scaleSel.value = viewState.printContentScale;
  viewState.personId = syncReportsPersonFilterOptions(select, activePersons, viewState.personId);

  document.querySelectorAll('input[name="reports-employee-visibility"]').forEach(input => {
    input.checked = input.value === viewState.employeeVisibility;
  });



  // Set orientation class on viewport
  if (viewport) {
    viewport.classList.toggle('rpt-landscape', viewState.printOrientation === 'landscape');
  }

  // Build A4 pages preview
  if (docEl) {
    docEl.innerHTML = buildReportsUnifiedDocumentHtml(selectionData, {
      contentScale: viewState.printContentScale || 1.0,
      employeeVisibility: viewState.employeeVisibility,
      printOrientation: viewState.printOrientation
    });

    // Theme class
    docEl.classList.remove('rpt-theme-app');
    docEl.classList.add('rpt-theme-print');

    normalizeCurrencySuffixSpacing(docEl);
  }

  // Apply zoom based on current mode
  requestAnimationFrame(() => {
    const mode = viewState.previewZoomMode || 'manual';
    if (mode === 'fit-width' && window.__reportsApplyFitWidth) {
      window.__reportsApplyFitWidth();
    } else if (mode === 'fit-page' && window.__reportsApplyFitPage) {
      window.__reportsApplyFitPage();
    } else if (window.__reportsApplyZoom) {
      window.__reportsApplyZoom(viewState.previewZoom || 100);
    }
  });
}






// Payout functions (getPayoutsStateBundle, buildPayoutEmployeeCardHtml, settlePayoutForEmployee,
// initPayouts, renderPayouts) are defined in payouts.js which is loaded after this file.

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

  setupListSortable(document.getElementById('clients-table-body'), {
    onEnd: () => {
      Store.reorderClients(getSortableRowIds(document.getElementById('clients-table-body')));
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

  state.clients.forEach((c, idx) => {
    const clientRate = parseFloat(c.hourlyRate);
    const isClientActive = Store.isClientActiveInMonth(c.id);
    const tr = document.createElement('tr');
    tr.dataset.id = c.id;
    if (!isClientActive) {
      tr.style.opacity = '0.5';
    }
    tr.innerHTML = `
      <td style="font-weight: 500">${c.name}</td>
      <td>${Number.isFinite(clientRate) && clientRate > 0 ? `${clientRate.toFixed(2)} zł/h` : '<span style="color: var(--danger);">Brak stawki</span>'}</td>
      <td>
        <button class="btn-status btn-status-client ${!isClientActive ? 'inactive' : 'active'}" data-id="${c.id}">
          ${!isClientActive ? 'Nieaktywny' : 'Aktywny'}
        </button>
      </td>
      <td>
        <div style="display: flex; gap: 0.5rem; flex-wrap: wrap;">
          <button class="btn btn-secondary btn-icon btn-edit-client" data-id="${c.id}">
            <i data-lucide="edit-2" style="width:16px;height:16px"></i>
          </button>
          <button class="btn btn-danger btn-icon btn-delete-client" data-id="${c.id}">
            <i data-lucide="trash-2" style="width:16px;height:16px"></i>
          </button>
          <button class="btn btn-secondary btn-icon drag-handle btn-reorder-handle" type="button" title="Przytrzymaj, aby zmienić kolejność" aria-label="Przytrzymaj, aby zmienić kolejność">
            <i data-lucide="grip-horizontal" style="width:16px;height:16px"></i>
          </button>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
    appendMobileCardDividerRow(tbody, 4, idx === state.clients.length - 1);
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
      const client = Store.getState().clients.find(c => c.id === id);
      if (!client) return;
      if (confirm(`Czy na pewno chcesz usunąć klienta "${client.name}"?`)) {
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

  setupListSortable(document.getElementById('people-table-body'), {
    onEnd: () => {
      Store.reorderPersons(getSortableRowIds(document.getElementById('people-table-body')));
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
    tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;color:var(--text-muted)">Brak rekordów</td></tr>`;
    return;
  }

  state.persons.forEach((p, idx) => {
    const isPersonActive = Store.isPersonActiveInMonth(p.id);
    const tr = document.createElement('tr');
    tr.dataset.id = p.id;
    if (!isPersonActive) {
      tr.style.opacity = '0.5';
    }
    tr.innerHTML = `
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
          <button class="btn btn-secondary btn-icon btn-edit-person" data-id="${p.id}">
            <i data-lucide="edit-2" style="width:16px;height:16px"></i>
          </button>
          <button class="btn btn-danger btn-icon btn-delete-person" data-id="${p.id}">
            <i data-lucide="trash-2" style="width:16px;height:16px"></i>
          </button>
          <button class="btn btn-secondary btn-icon drag-handle btn-reorder-handle" type="button" title="Przytrzymaj, aby zmienić kolejność" aria-label="Przytrzymaj, aby zmienić kolejność">
            <i data-lucide="grip-horizontal" style="width:16px;height:16px"></i>
          </button>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
    appendMobileCardDividerRow(tbody, 7, idx === state.persons.length - 1);
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
      if (confirm('Czy na pewno chcesz usunąć tę osobę? Osierocone wpisy tej osoby w aktywnym miesiącu zostaną automatycznie wyczyszczone.')) {
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

  setupListSortable(document.getElementById('hours-sheets-body'), {
    onEnd: () => {
      Store.reorderMonthlySheets(getSortableRowIds(document.getElementById('hours-sheets-body')));
    }
  });

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

  sheetsForMonth.forEach((s, idx) => {
    const visiblePersonIds = new Set(getVisiblePersonsForSheet(state, s).map(person => person.id));
    const totalH = Calculations.getSheetTotalHours(s, visiblePersonIds);

    const sheetRevenue = totalH * Calculations.getSheetClientRate(s, state);

    const tr = document.createElement('tr');
    tr.dataset.id = s.id;
    tr.innerHTML = `
      <td>${s.client || '-'}</td>
      <td>${s.site || '-'}</td>
      <td style="color: var(--primary); font-weight: bold;">${totalH.toFixed(1)}h</td>
      <td style="color: var(--success); font-weight: bold;">${sheetRevenue.toFixed(2)} zł</td>
      <td>
        <div style="display:flex; gap:0.5rem; flex-wrap:wrap;">
          <button class="btn btn-primary btn-sm btn-open-sheet" data-id="${s.id}" style="margin-right:0.5rem;">Otwórz</button>
          <button class="btn btn-secondary btn-icon btn-edit-sheet" data-id="${s.id}" style="margin-right:0.5rem;">
            <i data-lucide="edit-2" style="width:16px;height:16px"></i>
          </button>
          <button class="btn btn-danger btn-icon btn-delete-sheet" data-id="${s.id}">
            <i data-lucide="trash-2" style="width:16px;height:16px"></i>
          </button>
          <button class="btn btn-secondary btn-icon drag-handle btn-reorder-handle" type="button" title="Przytrzymaj, aby zmienić kolejność" aria-label="Przytrzymaj, aby zmienić kolejność">
            <i data-lucide="grip-horizontal" style="width:16px;height:16px"></i>
          </button>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
    appendMobileCardDividerRow(tbody, 5, idx === sheetsForMonth.length - 1);
  });

  document.querySelectorAll('.btn-open-sheet').forEach(btn => {
  btn.addEventListener('click', (e) => {
    const id = e.currentTarget.getAttribute('data-id');
    window.currentSheetId = id;
    activateNavigationTarget('hours-view');
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
  setupListSortable(document.getElementById('hours-sheets-body'), {
    onEnd: () => {
      Store.reorderMonthlySheets(getSortableRowIds(document.getElementById('hours-sheets-body')));
    }
  });
  normalizeCurrencySuffixSpacing(document.getElementById('hours-view'));
  applyArchivedReadOnlyMode(state.isArchived === true);
}

function renderSheetDetail(sheetId) {
  const state = Store.getState();
  const sheet = Store.getMonthlySheet(sheetId);
  const isReadOnly = state.isArchived === true;
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
        <input type="checkbox" class="fill-day-cb" data-day="${d}" ${isChecked} ${isReadOnly ? 'disabled' : ''} style="width: 18px; height: 18px; cursor: ${isReadOnly ? 'default' : 'pointer'};">
      </td>
      <td style="font-weight: bold; line-height: 1.1; text-align: center; padding: 0.35rem 0.3rem;">
        ${d}<br>
        <span style="font-size: 0.6rem; color: var(--text-secondary); font-weight: normal; white-space: nowrap;">${smallText}</span>
      </td>
      <td class="sheet-time-cell" style="font-size: 0.72rem; border: none; padding: 0.35rem 0.25rem;">
        <div class="sheet-time-range">
          <input type="text" class="time-in" data-day="${d}" data-type="start" readonly ${isReadOnly ? 'disabled' : ''} value="${globalStart}" placeholder="--:--" style="width: 44px; padding: 0.05rem; background: transparent; text-align: center; cursor: ${isReadOnly ? 'default' : 'pointer'}; border: 1px solid transparent; border-radius: 4px;">
          <span style="opacity: 0.5">-</span>
          <input type="text" class="time-in" data-day="${d}" data-type="end" readonly ${isReadOnly ? 'disabled' : ''} value="${globalEnd}" placeholder="--:--" style="width: 44px; padding: 0.05rem; background: transparent; text-align: center; cursor: ${isReadOnly ? 'default' : 'pointer'}; border: 1px solid transparent; border-radius: 4px;">
        </div>
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
                <input type="number" step="0.5" class="hour-in ${extraClass}" data-day="${d}" data-id="${p.id}" value="${valToShow}" ${isReadOnly ? 'readonly disabled' : ''} style="${manualStyle} text-align: center; width: 100%; border-radius: 4px; padding: 1px 2px;">
                ${isReadOnly ? '' : `<div class="cell-actions">
                  <button class="action-btn btn-set-zero" data-day="${d}" data-id="${p.id}" title="Ustaw 0 (nieobecność)">X</button>
                  <button class="action-btn btn-reset-manual" data-day="${d}" data-id="${p.id}" title="Przywróć automat">
                    <svg style="width: 12px; height: 12px;" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 0 0-9-9 9.75 9 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/><path d="M3 12a9 9 0 0 0 9 9 9.75 9 0 0 0 6.74-2.74L21 16"/><path d="M16 16h5v5"/></svg>
                  </button>
                </div>`}
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
                   <div class="sheet-footer-earnings-fit" style="color: var(--primary); font-size: 0.8rem; opacity: 0.9;">${formatPolishCurrencyWithSuffix(earnings)}</div>
                 </td>`;
  });
  tfoot.innerHTML = `<tr>${footHtml}</tr>`;

  const totalSheetHours = activePersons.reduce((sum, person) => sum + sums[person.id], 0);
  const totalSheetRevenue = totalSheetHours * sheetClientRate;
  document.getElementById('sheet-total-hours-summary').textContent = `${totalSheetHours.toFixed(1)}h`;
  document.getElementById('sheet-total-revenue-summary').textContent = formatPolishCurrencyWithSuffix(totalSheetRevenue);
  normalizeCurrencySuffixSpacing(document.getElementById('hours-view'));
  scheduleHoursSheetAutoFit();

  // Attach auto-save events
  if (isReadOnly) return;

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
      const dayData = ensureMonthlySheetDayData(s, d);
      
      if (val !== '') {
        dayData.hours[pid] = parseFloat(val);
        dayData.manual[pid] = true;
        if (isMonthlySheetPersonInactiveOnDay(s, pid, dayNumber) || getMonthlySheetPersonActivityOverride(s, d, pid) === 'inactive') {
          setMonthlySheetPersonActivityOverride(s, d, pid, 'active');
        }
        e.target.style.background = 'rgba(255, 255, 255, 0.25)';
        e.target.style.textAlign = 'center';
      } else {
        delete dayData.hours[pid];
        clearMonthlySheetPersonManualFlag(dayData, pid);
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
      const dayData = ensureMonthlySheetDayData(s, d);

      const currentHours = dayData.hours[pid];
      const isManualZero = currentHours === 0 && dayData.manual[pid] === true;
      const isInactive = isMonthlySheetPersonInactiveOnDay(s, pid, dayNumber);
      
      dayData.hours[pid] = 0;
      dayData.manual[pid] = true;

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
      const dayData = ensureMonthlySheetDayData(s, d);

      const todayOverride = getMonthlySheetPersonActivityOverride(s, d, pid);
      const isInactive = isMonthlySheetPersonInactiveOnDay(s, pid, dayNumber);

      clearMonthlySheetPersonManualFlag(dayData, pid);

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
      const dayData = ensureMonthlySheetDayData(s, d);
      
      dayData.isWholeTeamChecked = isChecked;
      
      if (isChecked) {
        const start = dayData.globalStart || '07:00';
        const end = dayData.globalEnd || '17:00';
        dayData.globalStart = start;
        dayData.globalEnd = end;
        
        const calcH = Calculations.calculateHours(start, end);
        if (calcH > 0) {
          activePersons.forEach(p => {
            if (!dayData.manual[p.id]) {
              syncMonthlySheetDayPersonHours(s, d, p.id);
            }
          });
        }
      } else {
        dayData.globalStart = '';
        dayData.globalEnd = '';
        activePersons.forEach(p => {
          if (!dayData.manual[p.id]) {
            delete dayData.hours[p.id];
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
      if (day && day.hours && day.hours[personId]) {
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
const ALL_PARTNERS_EXPENSE_TARGET_ID = 'all_partners';
const BONUS_EXPENSE_PAYER_ID = ALL_PARTNERS_EXPENSE_TARGET_ID;
const DIETA_EXPENSE_PAYER_ID = 'employee_profit_dieta';
const EXPENSE_DIETA_MODE_FIXED = 'FIXED';
const EXPENSE_DIETA_MODE_ACTIVE_DAYS = 'ACTIVE_DAYS';
const EXPENSE_DIETA_MODE_MANUAL_DAYS = 'MANUAL_DAYS';
const EXPENSE_REFUND_FROM_COST_MODE_PARTIAL = 'PARTIAL';
const EXPENSE_REFUND_FROM_COST_MODE_FULL = 'FULL';
const EXPENSE_REFUND_FROM_COST_MODE_CUSTOM = 'CUSTOM';
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

let expenseRefundFromCostContext = null;

function isExpenseSharedWithAllPartners(expense = {}) {
  return (expense?.advanceForId || '').toString().trim() === ALL_PARTNERS_EXPENSE_TARGET_ID;
}

function clearExpenseRefundFromCostContext() {
  expenseRefundFromCostContext = null;
}

function getExpenseRefundFromCostContext() {
  return expenseRefundFromCostContext;
}

function getActiveCostParticipantPersons(state = Store.getState(), month = getSelectedMonthKey()) {
  return Calculations.getActivePersons(state, month).filter(person =>
    Calculations.isPartnerLike(person) && Calculations.personParticipatesInCosts(person)
  );
}

function getDefaultExpenseRefundFromCostFractionDenominator(state = Store.getState(), month = getSelectedMonthKey()) {
  return Math.max(2, getActiveCostParticipantPersons(state, month).length + 1);
}

function normalizeExpenseRefundFromCostFractionDenominator(value, fallback = getDefaultExpenseRefundFromCostFractionDenominator()) {
  const parsed = parseInt((value || '').toString().trim(), 10);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.max(2, parsed);
}

function formatExpenseRefundFromCostFractionDenominator(value) {
  return `${normalizeExpenseRefundFromCostFractionDenominator(value)}`;
}

function clampExpenseRefundAmount(value = 0, maxAmount = 0) {
  const parsedValue = parseFloat(value);
  const parsedMax = Math.max(0, parseFloat(maxAmount) || 0);
  if (!Number.isFinite(parsedValue)) return 0;
  return Math.max(0, Math.min(parsedMax, parsedValue));
}

function getExpenseSelectedRefundFromCostMode() {
  const selected = document.querySelector('input[name="expense-refund-from-cost-mode"]:checked');
  const mode = (selected?.value || '').toString().trim().toUpperCase();
  if (mode === EXPENSE_REFUND_FROM_COST_MODE_FULL || mode === EXPENSE_REFUND_FROM_COST_MODE_CUSTOM) return mode;
  return EXPENSE_REFUND_FROM_COST_MODE_PARTIAL;
}

function setExpenseSelectedRefundFromCostMode(mode = EXPENSE_REFUND_FROM_COST_MODE_PARTIAL) {
  const normalizedMode = [
    EXPENSE_REFUND_FROM_COST_MODE_PARTIAL,
    EXPENSE_REFUND_FROM_COST_MODE_FULL,
    EXPENSE_REFUND_FROM_COST_MODE_CUSTOM
  ].includes(mode)
    ? mode
    : EXPENSE_REFUND_FROM_COST_MODE_PARTIAL;
  const target = document.querySelector(`input[name="expense-refund-from-cost-mode"][value="${normalizedMode}"]`)
    || document.querySelector(`input[name="expense-refund-from-cost-mode"][value="${EXPENSE_REFUND_FROM_COST_MODE_PARTIAL}"]`);
  if (target) target.checked = true;
}

function getExpensePaidByOptionsHtml(options = {}) {
  const {
    selectedId = '',
    excludeIds = [],
    includeClients = true,
    placeholder = '-- Wybierz płatnika --'
  } = options;
  const excludedIds = new Set((excludeIds || []).filter(Boolean));
  const state = Store.getState();
  const personsOptionsHtml = (state.persons || [])
    .filter(person => !excludedIds.has(person.id))
    .map(person => `<option value="${person.id}" ${person.id === selectedId ? 'selected' : ''}>${getPersonDisplayName(person)}</option>`)
    .join('');
  const clientsOptionsHtml = includeClients
    ? (state.clients || []).map(client => `<option value="client_${client.id}" ${`client_${client.id}` === selectedId ? 'selected' : ''}>${client.name}</option>`).join('')
    : '';

  return `<option value="">${placeholder}</option>`
    + (personsOptionsHtml ? `<optgroup label="Zespół">${personsOptionsHtml}</optgroup>` : '')
    + (clientsOptionsHtml ? `<optgroup label="Klienci">${clientsOptionsHtml}</optgroup>` : '');
}

function populateExpensePaidBySelect(options = {}) {
  const selectPaidBy = document.getElementById('expense-paid-by');
  if (!selectPaidBy) return;
  selectPaidBy.innerHTML = getExpensePaidByOptionsHtml(options);
}

function getExpenseRecipientName(expense = {}, state = Store.getState()) {
  if (expense?.type === 'REFUND' && isExpenseSharedWithAllPartners(expense)) {
    return 'Wszyscy wspólnicy';
  }

  return getPersonDisplayName((state?.persons || []).find(person => person.id === expense?.advanceForId)) || '-';
}

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
  const recipients = (Store.getState().persons || []).filter(person => (type !== 'BONUS' && type !== 'DIETA') || person.type === 'EMPLOYEE');
  const placeholder = (type === 'BONUS' || type === 'DIETA')
    ? '-- Wybierz pracownika --'
    : (type === 'REFUND' ? '-- Wybierz odbiorcę zwrotu --' : '-- Wybierz odbiorcę --');

  return `<option value="">${placeholder}</option>` + recipients.map(person => `
    <option value="${person.id}" ${person.id === selectedRecipientId ? 'selected' : ''}>${getPersonDisplayName(person)}</option>
  `).join('');
}

function populateExpenseRecipientSelect(type = 'ADVANCE', selectedRecipientId = '') {
  const selectAdvanceFor = document.getElementById('expense-advance-for');
  if (!selectAdvanceFor) return;

  selectAdvanceFor.innerHTML = getExpenseRecipientOptionsHtml(type, selectedRecipientId);
}

function updateExpenseRefundFromCostAmountUi() {
  const context = getExpenseRefundFromCostContext();
  const container = document.getElementById('expense-refund-from-cost-container');
  const summary = document.getElementById('expense-refund-from-cost-summary');
  const fractionContainer = document.getElementById('expense-refund-from-cost-fraction-container');
  const fractionLabel = document.getElementById('expense-refund-from-cost-fraction-label');
  const fractionInput = document.getElementById('expense-refund-from-cost-fraction');
  const customContainer = document.getElementById('expense-refund-from-cost-custom-container');
  const rangeInput = document.getElementById('expense-refund-from-cost-range');
  const rangeValue = document.getElementById('expense-refund-from-cost-range-value');
  const amountInput = document.getElementById('expense-amount');
  const amountLabel = document.getElementById('expense-amount-label');

  if (!container || !amountInput || !amountLabel) return;

  if (!context) {
    container.style.display = 'none';
    if (summary) summary.textContent = '';
    if (fractionContainer) fractionContainer.style.display = 'none';
    if (customContainer) customContainer.style.display = 'none';
    amountInput.readOnly = false;
    amountInput.max = '';
    return;
  }

  const maxAmount = Math.max(0, parseFloat(context.costAmount) || 0);
  const mode = getExpenseSelectedRefundFromCostMode();
  const denominator = normalizeExpenseRefundFromCostFractionDenominator(
    fractionInput?.value,
    context.defaultFractionDenominator || getDefaultExpenseRefundFromCostFractionDenominator()
  );
  const partialAmount = denominator > 0 ? Number((maxAmount / denominator).toFixed(2)) : 0;

  container.style.display = 'block';
  if (summary) {
    summary.textContent = `Koszt źródłowy: ${context.costName} • zapłacił ${getSettlementEntityName(Store.getState(), context.costPaidById)} • pozostało ${formatSettlementCompactCurrency(maxAmount)}.`;
  }

  if (fractionInput) {
    fractionInput.value = formatExpenseRefundFromCostFractionDenominator(denominator);
  }

  if (fractionContainer) {
    fractionContainer.style.display = mode === EXPENSE_REFUND_FROM_COST_MODE_PARTIAL ? 'block' : 'none';
  }

  if (fractionLabel) {
    fractionLabel.textContent = `Część zwrotu 1/${denominator} (≈ ${formatSettlementCompactCurrency(partialAmount)})`;
  }

  if (customContainer) {
    customContainer.style.display = mode === EXPENSE_REFUND_FROM_COST_MODE_CUSTOM ? 'block' : 'none';
  }

  if (rangeInput) {
    const clampedAmount = clampExpenseRefundAmount(amountInput.value, maxAmount);
    const percent = maxAmount > 0 ? Math.round((clampedAmount / maxAmount) * 100) : 0;

    if (mode === EXPENSE_REFUND_FROM_COST_MODE_CUSTOM) {
      rangeInput.value = `${percent}`;
      if (rangeValue) rangeValue.textContent = `${percent}% • ${formatSettlementCompactCurrency(clampedAmount)}`;
      amountInput.readOnly = false;
      amountInput.value = clampedAmount > 0 ? clampedAmount.toFixed(2) : '';
    } else {
      const resolvedAmount = mode === EXPENSE_REFUND_FROM_COST_MODE_FULL ? maxAmount : partialAmount;
      amountInput.readOnly = true;
      amountInput.value = resolvedAmount > 0 ? resolvedAmount.toFixed(2) : '';
      rangeInput.value = maxAmount > 0 ? `${Math.round((resolvedAmount / maxAmount) * 100)}` : '0';
      if (rangeValue) rangeValue.textContent = `${rangeInput.value}% • ${formatSettlementCompactCurrency(resolvedAmount)}`;
    }
  }

  amountInput.min = '0';
  amountInput.max = `${maxAmount}`;
  amountLabel.textContent = `Kwota zwrotu (max ${formatSettlementCompactCurrency(maxAmount)})`;
}

function openExpenseRefundFromCostForm(costExpense = null) {
  const form = document.getElementById('expense-form');
  const typeSelect = document.getElementById('expense-type');
  const nameInput = document.getElementById('expense-name');
  const recipientSelect = document.getElementById('expense-advance-for');
  const suggestionsButton = document.getElementById('btn-expense-name-suggestions');
  if (!form || !typeSelect || !nameInput || !recipientSelect || !costExpense) return;

  const state = Store.getState();
  if (!costExpense.paidById || costExpense.paidById.startsWith('client_')) {
    alert('Zwrot z kosztu można utworzyć tylko dla kosztu opłaconego przez konkretną osobę.');
    return;
  }

  const availablePayers = (state.persons || []).filter(person => person.id !== costExpense.paidById);
  if (availablePayers.length === 0) {
    alert('Brak innej osoby, która mogłaby wykonać zwrot tego kosztu.');
    return;
  }

  expenseRefundFromCostContext = {
    costExpenseId: costExpense.id,
    costName: costExpense.name || 'Koszt',
    costPaidById: costExpense.paidById,
    costAmount: Math.max(0, parseFloat(costExpense.amount) || 0),
    defaultFractionDenominator: getDefaultExpenseRefundFromCostFractionDenominator(state, getSelectedMonthKey())
  };

  document.getElementById('expense-form-title').textContent = 'Nowy Zwrot z Kosztu';
  document.getElementById('expense-id').value = '';
  document.getElementById('expense-date').value = costExpense.date || getDefaultDateForSelectedMonth();
  nameInput.value = expenseRefundFromCostContext.costName;
  document.getElementById('expense-refund-all-partners').checked = false;
  document.getElementById('expense-dieta-days-adjustment').value = '0';
  document.getElementById('expense-refund-from-cost-fraction').value = `${expenseRefundFromCostContext.defaultFractionDenominator}`;
  document.getElementById('expense-amount').value = '';
  typeSelect.value = 'REFUND';
  typeSelect.disabled = true;
  nameInput.readOnly = true;
  if (suggestionsButton) suggestionsButton.disabled = true;

  recipientSelect.innerHTML = `<option value="${expenseRefundFromCostContext.costPaidById}" selected>${getSettlementEntityName(state, expenseRefundFromCostContext.costPaidById)}</option>`;
  recipientSelect.value = expenseRefundFromCostContext.costPaidById;
  recipientSelect.disabled = true;

  populateExpensePaidBySelect({
    excludeIds: [expenseRefundFromCostContext.costPaidById],
    includeClients: false,
    placeholder: '-- Wybierz osobę zwracającą --'
  });

  setExpenseSelectedRefundFromCostMode(EXPENSE_REFUND_FROM_COST_MODE_PARTIAL);
  updateExpenseFormTypeUI('REFUND', expenseRefundFromCostContext.costPaidById);

  document.querySelector('#expenses-view .table-container').style.display = 'none';
  form.style.display = 'block';
  form.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function parseSignedIntegerInput(value, fallback = 0) {
  const normalizedValue = (value || '').toString().trim().replace(/\s+/g, '');
  if (!/^[+-]?\d+$/.test(normalizedValue)) return fallback;

  const parsed = parseInt(normalizedValue, 10);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function formatSignedIntegerInput(value) {
  const parsed = parseSignedIntegerInput(value, 0);
  if (parsed > 0) return `+${parsed}`;
  return `${parsed}`;
}

function normalizeExpenseDietaModeValue(value) {
  if (value === EXPENSE_DIETA_MODE_ACTIVE_DAYS || value === EXPENSE_DIETA_MODE_MANUAL_DAYS) return value;
  return EXPENSE_DIETA_MODE_FIXED;
}

function getExpenseSelectedDietaMode() {
  const selected = document.querySelector('input[name="expense-dieta-mode"]:checked');
  return normalizeExpenseDietaModeValue(selected?.value);
}

function setExpenseSelectedDietaMode(mode = EXPENSE_DIETA_MODE_FIXED) {
  const normalizedMode = normalizeExpenseDietaModeValue(mode);
  const target = document.querySelector(`input[name="expense-dieta-mode"][value="${normalizedMode}"]`)
    || document.querySelector(`input[name="expense-dieta-mode"][value="${EXPENSE_DIETA_MODE_FIXED}"]`);
  if (target) target.checked = true;
}

function getExpenseDietaModeFromExpense(expense = null) {
  const mode = Calculations.getDietaCalculationMode(expense);
  return normalizeExpenseDietaModeValue(mode);
}

function isExpenseDietaPerDayMode(mode = getExpenseSelectedDietaMode()) {
  return mode === EXPENSE_DIETA_MODE_ACTIVE_DAYS || mode === EXPENSE_DIETA_MODE_MANUAL_DAYS;
}

function parseExpenseManualDaysInput(value, fallback = 0) {
  const parsed = parseInt((value || '').toString().trim(), 10);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.max(0, parsed);
}

function formatExpenseManualDaysInput(value) {
  return `${parseExpenseManualDaysInput(value, 0)}`;
}

function adjustExpenseDietaDays(delta = 0) {
  const adjustmentInput = document.getElementById('expense-dieta-days-adjustment');
  const advanceForSelect = document.getElementById('expense-advance-for');
  if (!adjustmentInput) return;

  const dietaMode = getExpenseSelectedDietaMode();
  if (dietaMode === EXPENSE_DIETA_MODE_MANUAL_DAYS) {
    const nextValue = Math.max(0, parseExpenseManualDaysInput(adjustmentInput.value, 0) + delta);
    adjustmentInput.value = formatExpenseManualDaysInput(nextValue);
  } else {
    const nextValue = parseSignedIntegerInput(adjustmentInput.value, 0) + delta;
    adjustmentInput.value = formatSignedIntegerInput(nextValue);
  }

  updateExpenseDietaDaysAdjustmentLabel(advanceForSelect?.value || '');
  updateExpenseAmountLabel(advanceForSelect?.value || '');
}

function resetExpenseDietaDays() {
  const adjustmentInput = document.getElementById('expense-dieta-days-adjustment');
  const advanceForSelect = document.getElementById('expense-advance-for');
  if (!adjustmentInput) return;

  adjustmentInput.value = getExpenseSelectedDietaMode() === EXPENSE_DIETA_MODE_MANUAL_DAYS
    ? formatExpenseManualDaysInput(0)
    : formatSignedIntegerInput(0);

  updateExpenseDietaDaysAdjustmentLabel(advanceForSelect?.value || '');
  updateExpenseAmountLabel(advanceForSelect?.value || '');
}

function updateExpenseDietaDaysAdjustmentLabel(selectedRecipientId = '') {
  const label = document.getElementById('expense-dieta-days-adjustment-label');
  const typeSelect = document.getElementById('expense-type');
  const adjustmentInput = document.getElementById('expense-dieta-days-adjustment');
  if (!label || !typeSelect || !adjustmentInput) return;

  const dietaMode = getExpenseSelectedDietaMode();

  if (typeSelect.value !== 'DIETA' || dietaMode === EXPENSE_DIETA_MODE_FIXED) {
    label.textContent = 'Koryguj dni';
    return;
  }

  if (dietaMode === EXPENSE_DIETA_MODE_MANUAL_DAYS) {
    label.textContent = 'Liczba dni';
    return;
  }

  if (!selectedRecipientId) {
    label.textContent = 'Koryguj dni';
    return;
  }

  const state = Store.getState();
  const selectedMonth = getSelectedMonthKey();
  const baseDays = Calculations.getPersonActiveWorkDaysCount(selectedRecipientId, state, selectedMonth);
  const adjustment = parseSignedIntegerInput(adjustmentInput.value, 0);
  const effectiveDays = Math.max(0, baseDays + adjustment);

  if (adjustment > 0) {
    label.textContent = `Koryguj dni (${baseDays}+${adjustment} = ${effectiveDays})`;
    return;
  }

  if (adjustment < 0) {
    label.textContent = `Koryguj dni (${baseDays}${adjustment} = ${effectiveDays})`;
    return;
  }

  label.textContent = `Koryguj dni (${baseDays})`;
}

function updateExpenseAmountLabel(selectedRecipientId = '') {
  const amountLabel = document.getElementById('expense-amount-label');
  const typeSelect = document.getElementById('expense-type');
  const amountInput = document.getElementById('expense-amount');
  const adjustmentInput = document.getElementById('expense-dieta-days-adjustment');
  if (!amountLabel || !typeSelect || !amountInput || !adjustmentInput) return;

  const dietaMode = getExpenseSelectedDietaMode();

  if (typeSelect.value !== 'DIETA') {
    amountLabel.textContent = 'Kwota (zł)';
    return;
  }

  if (!isExpenseDietaPerDayMode(dietaMode)) {
    amountLabel.textContent = 'Kwota (zł)';
    return;
  }

  const dayRate = parseFloat(amountInput.value);
  const requiresRecipient = dietaMode === EXPENSE_DIETA_MODE_ACTIVE_DAYS;
  if ((requiresRecipient && !selectedRecipientId) || !Number.isFinite(dayRate) || dayRate <= 0) {
    amountLabel.textContent = 'Kwota za dzień (zł)';
    return;
  }

  const totalAmount = Calculations.getExpenseEffectiveAmount({
    type: 'DIETA',
    dietaCalculationMode: dietaMode,
    dietaByActiveDays: dietaMode === EXPENSE_DIETA_MODE_ACTIVE_DAYS,
    advanceForId: selectedRecipientId,
    amount: dayRate,
    dietDaysAdjustment: dietaMode === EXPENSE_DIETA_MODE_ACTIVE_DAYS
      ? parseSignedIntegerInput(adjustmentInput.value, 0)
      : 0,
    dietaDaysCount: dietaMode === EXPENSE_DIETA_MODE_MANUAL_DAYS
      ? parseExpenseManualDaysInput(adjustmentInput.value, 0)
      : 0
  }, Store.getState(), getSelectedMonthKey());

  amountLabel.textContent = `Kwota za dzień (zł) (Łącznie ${formatSettlementCompactCurrency(totalAmount)})`;
}

function updateExpenseFormTypeUI(type, selectedRecipientId = '') {
  const advanceForContainer = document.getElementById('expense-advance-for-container');
  const advanceForSelect = document.getElementById('expense-advance-for');
  const expenseNameContainer = document.getElementById('expense-name-container');
  const expenseNameInput = document.getElementById('expense-name');
  const expenseNameSuggestionsButton = document.getElementById('btn-expense-name-suggestions');
  const paidBySelect = document.getElementById('expense-paid-by');
  const paidByLabel = document.getElementById('expense-paid-by-label');
  const paidByContainer = paidBySelect ? paidBySelect.closest('.form-group') : null;
  const recipientLabel = document.getElementById('expense-recipient-label');
  const dietaModeContainer = document.getElementById('expense-dieta-mode-container');
  const dietaDaysAdjustmentContainer = document.getElementById('expense-dieta-days-adjustment-container');
  const dietaDaysInput = document.getElementById('expense-dieta-days-adjustment');
  const refundAllPartnersContainer = document.getElementById('expense-refund-all-partners-container');
  const refundAllPartnersCheckbox = document.getElementById('expense-refund-all-partners');
  const refundFromCostContext = getExpenseRefundFromCostContext();
  const isRefundFromCost = type === 'REFUND' && !!refundFromCostContext;
  const typeSelect = document.getElementById('expense-type');
  const isAdvance = type === 'ADVANCE';
  const isBonus = type === 'BONUS';
  const isDieta = type === 'DIETA';
  const isRefund = type === 'REFUND';
  const dietaMode = getExpenseSelectedDietaMode();
  const isRefundSharedWithPartners = isRefund && !isRefundFromCost && refundAllPartnersCheckbox?.checked === true;

  if (typeSelect) {
    typeSelect.disabled = isRefundFromCost;
  }

  if (advanceForContainer) {
    advanceForContainer.style.display = (isAdvance || isBonus || isDieta || (isRefund && !isRefundSharedWithPartners)) ? 'block' : 'none';
  }

  if (expenseNameContainer && expenseNameInput) {
    expenseNameContainer.style.display = (isAdvance || isBonus || isDieta) ? 'none' : 'block';
    expenseNameInput.required = !(isAdvance || isBonus || isDieta);
    expenseNameInput.readOnly = isRefundFromCost;
    if (isBonus) expenseNameInput.value = 'Premia';
    if (isAdvance) expenseNameInput.value = 'Zaliczka';
    if (isDieta) expenseNameInput.value = 'Dieta';
  }

  if (expenseNameSuggestionsButton) {
    expenseNameSuggestionsButton.disabled = isRefundFromCost;
    expenseNameSuggestionsButton.style.display = (isAdvance || isBonus || isDieta) ? 'none' : '';
  }

  if (refundAllPartnersContainer) {
    refundAllPartnersContainer.style.display = isRefund && !isRefundFromCost ? 'block' : 'none';
  }

  if (refundAllPartnersCheckbox && (!isRefund || isRefundFromCost)) {
    refundAllPartnersCheckbox.checked = false;
  }

  if (paidByContainer && paidBySelect) {
    paidByContainer.style.display = (isBonus || isDieta) ? 'none' : 'block';
    paidBySelect.required = !(isBonus || isDieta);
    if (isBonus || isDieta) {
      paidBySelect.value = '';
    } else if (!isRefundFromCost) {
      populateExpensePaidBySelect({ selectedId: paidBySelect.value || '' });
    }
  }

  if (paidByLabel) {
    paidByLabel.textContent = isRefund ? 'Kto zwraca?' : 'Kto płacił?';
  }

  if (recipientLabel) {
    recipientLabel.textContent = isBonus
      ? 'Dla kogo Premia?'
      : (isDieta ? 'Dla kogo dieta?' : (isRefund ? 'Dla kogo zwrot?' : 'Dla kogo zaliczka?'));
  }

  if (dietaModeContainer) {
    dietaModeContainer.style.display = isDieta ? 'block' : 'none';
    if (!isDieta) {
      setExpenseSelectedDietaMode(EXPENSE_DIETA_MODE_FIXED);
    }
  }

  if (dietaDaysAdjustmentContainer) {
    dietaDaysAdjustmentContainer.style.display = isDieta && isExpenseDietaPerDayMode(dietaMode) ? 'block' : 'none';
  }

  if (dietaDaysInput) {
    dietaDaysInput.value = dietaMode === EXPENSE_DIETA_MODE_MANUAL_DAYS
      ? formatExpenseManualDaysInput(dietaDaysInput.value)
      : formatSignedIntegerInput(dietaDaysInput.value);
  }

  if (advanceForSelect) {
    if (isRefundFromCost && refundFromCostContext) {
      advanceForSelect.innerHTML = `<option value="${refundFromCostContext.costPaidById}" selected>${getSettlementEntityName(Store.getState(), refundFromCostContext.costPaidById)}</option>`;
      advanceForSelect.value = refundFromCostContext.costPaidById;
      advanceForSelect.disabled = true;
    } else {
      advanceForSelect.disabled = false;
      populateExpenseRecipientSelect(type, selectedRecipientId);
    }
  }

  updateExpenseDietaDaysAdjustmentLabel(selectedRecipientId);
  updateExpenseAmountLabel(selectedRecipientId);
  updateExpenseRefundFromCostAmountUi();
}

function openExpenseFormForCreate() {
  const form = document.getElementById('expense-form');
  const typeSelect = document.getElementById('expense-type');
  if (!form || !typeSelect) return;

   clearExpenseRefundFromCostContext();

  document.getElementById('expense-form-title').textContent = 'Nowy Dowód/Koszt';
  document.getElementById('expense-id').value = '';
  document.getElementById('expense-date').value = getDefaultDateForSelectedMonth();
  document.getElementById('expense-name').value = '';
  document.getElementById('expense-amount').value = '';
  populateExpensePaidBySelect();
  document.getElementById('expense-paid-by').value = '';
  document.getElementById('expense-advance-for').value = '';
  document.getElementById('expense-refund-all-partners').checked = false;
  setExpenseSelectedDietaMode(EXPENSE_DIETA_MODE_FIXED);
  document.getElementById('expense-dieta-days-adjustment').value = '0';
  document.getElementById('expense-refund-from-cost-fraction').value = `${getDefaultExpenseRefundFromCostFractionDenominator()}`;
  document.getElementById('expense-type').disabled = false;
  document.getElementById('expense-name').readOnly = false;
  const expenseNameSuggestionsButton = document.getElementById('btn-expense-name-suggestions');
  if (expenseNameSuggestionsButton) expenseNameSuggestionsButton.disabled = false;
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
  const amountInput = document.getElementById('expense-amount');
  const refundFromCostModeInputs = document.querySelectorAll('input[name="expense-refund-from-cost-mode"]');
  const refundFromCostFractionInput = document.getElementById('expense-refund-from-cost-fraction');
  const refundFromCostRangeInput = document.getElementById('expense-refund-from-cost-range');
  const dietaModeInputs = document.querySelectorAll('input[name="expense-dieta-mode"]');
  const dietaDaysAdjustmentInput = document.getElementById('expense-dieta-days-adjustment');
  const advanceForSelect = document.getElementById('expense-advance-for');
  const refundAllPartnersCheckbox = document.getElementById('expense-refund-all-partners');
  const refundFractionDecreaseButton = document.getElementById('btn-expense-refund-fraction-decrease');
  const refundFractionIncreaseButton = document.getElementById('btn-expense-refund-fraction-increase');
  const refundFractionResetButton = document.getElementById('btn-expense-refund-fraction-reset');
  const dietaDaysDecreaseButton = document.getElementById('btn-expense-dieta-days-decrease');
  const dietaDaysIncreaseButton = document.getElementById('btn-expense-dieta-days-increase');
  const dietaDaysResetButton = document.getElementById('btn-expense-dieta-days-reset');

  setupListSortable(document.getElementById('expenses-table-body'), {
    onMove: (evt) => allowSortOnlyWithinSameDate(evt),
    onEnd: (evt) => {
      const date = evt.item?.dataset?.sortDate || '';
      if (!date) return;
      const selector = `tr[data-id][data-sort-date="${date}"]`;
      Store.reorderExpensesForDate(date, getSortableRowIds(document.getElementById('expenses-table-body'), selector));
    }
  });

  initExpenseNameSuggestions();

  document.getElementById('expense-date').value = getDefaultDateForSelectedMonth();

  btnAdd.addEventListener('click', openExpenseFormForCreate);

  btnCancel.addEventListener('click', () => {
    clearExpenseRefundFromCostContext();
    typeSelect.disabled = false;
    document.getElementById('expense-name').readOnly = false;
    const expenseNameSuggestionsButton = document.getElementById('btn-expense-name-suggestions');
    if (expenseNameSuggestionsButton) expenseNameSuggestionsButton.disabled = false;
    form.style.display = 'none';
    document.querySelector('#expenses-view .table-container').style.display = 'block';
  });

  typeSelect.addEventListener('change', (e) => {
    updateExpenseFormTypeUI(e.target.value);
  });

  if (advanceForSelect) {
    advanceForSelect.addEventListener('change', () => {
      updateExpenseDietaDaysAdjustmentLabel(advanceForSelect.value || '');
      updateExpenseAmountLabel(advanceForSelect.value || '');
    });
  }

  if (refundAllPartnersCheckbox) {
    refundAllPartnersCheckbox.addEventListener('change', () => {
      updateExpenseFormTypeUI(typeSelect.value, advanceForSelect?.value || '');
    });
  }

  if (amountInput) {
    amountInput.addEventListener('input', () => {
      if (getExpenseRefundFromCostContext()) {
        updateExpenseRefundFromCostAmountUi();
      }
      updateExpenseAmountLabel(advanceForSelect?.value || '');
    });
  }

  refundFromCostModeInputs.forEach(input => {
    input.addEventListener('change', () => {
      updateExpenseRefundFromCostAmountUi();
    });
  });

  if (refundFromCostFractionInput) {
    refundFromCostFractionInput.addEventListener('input', () => {
      refundFromCostFractionInput.value = refundFromCostFractionInput.value.replace(/[^\d]/g, '');
      updateExpenseRefundFromCostAmountUi();
    });

    refundFromCostFractionInput.addEventListener('blur', () => {
      refundFromCostFractionInput.value = formatExpenseRefundFromCostFractionDenominator(refundFromCostFractionInput.value);
      updateExpenseRefundFromCostAmountUi();
    });
  }

  if (refundFractionDecreaseButton) {
    refundFractionDecreaseButton.addEventListener('click', () => {
      const nextValue = normalizeExpenseRefundFromCostFractionDenominator(refundFromCostFractionInput?.value, getDefaultExpenseRefundFromCostFractionDenominator()) - 1;
      if (refundFromCostFractionInput) refundFromCostFractionInput.value = formatExpenseRefundFromCostFractionDenominator(nextValue);
      updateExpenseRefundFromCostAmountUi();
    });
  }

  if (refundFractionIncreaseButton) {
    refundFractionIncreaseButton.addEventListener('click', () => {
      const nextValue = normalizeExpenseRefundFromCostFractionDenominator(refundFromCostFractionInput?.value, getDefaultExpenseRefundFromCostFractionDenominator()) + 1;
      if (refundFromCostFractionInput) refundFromCostFractionInput.value = formatExpenseRefundFromCostFractionDenominator(nextValue);
      updateExpenseRefundFromCostAmountUi();
    });
  }

  if (refundFractionResetButton) {
    refundFractionResetButton.addEventListener('click', () => {
      if (refundFromCostFractionInput) refundFromCostFractionInput.value = `${getDefaultExpenseRefundFromCostFractionDenominator()}`;
      updateExpenseRefundFromCostAmountUi();
    });
  }

  if (refundFromCostRangeInput) {
    refundFromCostRangeInput.addEventListener('input', () => {
      const context = getExpenseRefundFromCostContext();
      if (!context || !amountInput) return;
      const maxAmount = Math.max(0, parseFloat(context.costAmount) || 0);
      const percent = Math.max(0, Math.min(100, parseInt(refundFromCostRangeInput.value, 10) || 0));
      const nextAmount = clampExpenseRefundAmount((maxAmount * percent) / 100, maxAmount);
      amountInput.value = nextAmount > 0 ? nextAmount.toFixed(2) : '';
      updateExpenseRefundFromCostAmountUi();
    });
  }

  dietaModeInputs.forEach(input => {
    input.addEventListener('change', () => {
      updateExpenseFormTypeUI(typeSelect.value, document.getElementById('expense-advance-for')?.value || '');
    });
  });

  if (dietaDaysAdjustmentInput) {
    dietaDaysAdjustmentInput.addEventListener('input', () => {
      const dietaMode = getExpenseSelectedDietaMode();
      if (dietaMode === EXPENSE_DIETA_MODE_MANUAL_DAYS) {
        dietaDaysAdjustmentInput.value = dietaDaysAdjustmentInput.value.replace(/[^\d]/g, '');
      }
      updateExpenseDietaDaysAdjustmentLabel(advanceForSelect?.value || '');
      updateExpenseAmountLabel(advanceForSelect?.value || '');
    });

    dietaDaysAdjustmentInput.addEventListener('blur', () => {
      const dietaMode = getExpenseSelectedDietaMode();
      dietaDaysAdjustmentInput.value = dietaMode === EXPENSE_DIETA_MODE_MANUAL_DAYS
        ? formatExpenseManualDaysInput(dietaDaysAdjustmentInput.value)
        : formatSignedIntegerInput(dietaDaysAdjustmentInput.value);
      updateExpenseDietaDaysAdjustmentLabel(advanceForSelect?.value || '');
      updateExpenseAmountLabel(advanceForSelect?.value || '');
    });
  }

  if (dietaDaysDecreaseButton) {
    dietaDaysDecreaseButton.addEventListener('click', () => {
      adjustExpenseDietaDays(-1);
    });
  }

  if (dietaDaysIncreaseButton) {
    dietaDaysIncreaseButton.addEventListener('click', () => {
      adjustExpenseDietaDays(1);
    });
  }

  if (dietaDaysResetButton) {
    dietaDaysResetButton.addEventListener('click', () => {
      resetExpenseDietaDays();
    });
  }

  form.addEventListener('submit', (e) => {
    e.preventDefault();
    const id = document.getElementById('expense-id').value;
    const type = typeSelect.value;
    const isAdvance = type === 'ADVANCE';
    const isBonus = type === 'BONUS';
    const isDieta = type === 'DIETA';
    const isRefund = type === 'REFUND';
    const dietaCalculationMode = isDieta ? getExpenseSelectedDietaMode() : EXPENSE_DIETA_MODE_FIXED;
    const date = document.getElementById('expense-date').value;
    const name = isAdvance ? 'Zaliczka' : (isBonus ? 'Premia' : (isDieta ? 'Dieta' : document.getElementById('expense-name').value));
    const amount = parseFloat(document.getElementById('expense-amount').value);
    const paidById = isBonus ? BONUS_EXPENSE_PAYER_ID : (isDieta ? DIETA_EXPENSE_PAYER_ID : document.getElementById('expense-paid-by').value);
    const advanceForId = (isAdvance || isBonus || isDieta)
      ? document.getElementById('expense-advance-for').value
      : (isRefund
          ? (refundAllPartnersCheckbox?.checked ? ALL_PARTNERS_EXPENSE_TARGET_ID : document.getElementById('expense-advance-for').value)
          : '');
    const dietaByActiveDays = isDieta && dietaCalculationMode === EXPENSE_DIETA_MODE_ACTIVE_DAYS;
    const dietDaysAdjustment = isDieta && dietaCalculationMode === EXPENSE_DIETA_MODE_ACTIVE_DAYS
      ? parseSignedIntegerInput(document.getElementById('expense-dieta-days-adjustment').value, 0)
      : 0;
    const dietaDaysCount = isDieta && dietaCalculationMode === EXPENSE_DIETA_MODE_MANUAL_DAYS
      ? parseExpenseManualDaysInput(document.getElementById('expense-dieta-days-adjustment').value, 0)
      : 0;
    const selectedMonth = getSelectedMonthKey();
    const refundFromCostContext = getExpenseRefundFromCostContext();

    if (!isBonus && !isDieta && !paidById) {
      alert('Wybierz osobę dokonującą płatności');
      return;
    }
    if ((isAdvance || isBonus || isDieta) && !advanceForId) {
      alert(isBonus
        ? 'Wybierz pracownika, który otrzymuje premię.'
        : (isDieta ? 'Wybierz pracownika, który otrzymuje dietę.' : 'Wybierz osobę która otrzymuje zaliczkę'));
      return;
    }
    if (isRefund && !advanceForId) {
      alert('Wybierz odbiorcę zwrotu lub zaznacz opcję dla wszystkich wspólników.');
      return;
    }
    if (refundFromCostContext && paidById === refundFromCostContext.costPaidById) {
      alert('Nie można tworzyć zwrotu od tej samej osoby, która opłaciła koszt.');
      return;
    }
    if (!date || !date.startsWith(`${selectedMonth}-`)) {
      alert('Data wpisu musi należeć do wybranego miesiąca.');
      document.getElementById('expense-date').focus();
      return;
    }

    if (refundFromCostContext) {
      const refundAmount = clampExpenseRefundAmount(amount, refundFromCostContext.costAmount);
      if (!(refundAmount > 0)) {
        alert('Kwota zwrotu musi być większa od 0.');
        return;
      }
      if (refundAmount - refundFromCostContext.costAmount > 0.005) {
        alert('Kwota zwrotu nie może być wyższa od pozostałej kwoty kosztu.');
        return;
      }

      const refundCreated = Store.addExpense({
        type: 'REFUND',
        date,
        name: refundFromCostContext.costName,
        amount: refundAmount,
        paidById,
        advanceForId: refundFromCostContext.costPaidById,
        dietaCalculationMode: EXPENSE_DIETA_MODE_FIXED,
        dietaByActiveDays: false,
        dietDaysAdjustment: 0,
        dietaDaysCount: 0
      });
      if (!refundCreated) {
        alert('Nie udało się dodać zwrotu.');
        return;
      }

      const remainingCostAmount = Number((refundFromCostContext.costAmount - refundAmount).toFixed(2));
      if (remainingCostAmount <= 0.005) {
        Store.deleteExpense(refundFromCostContext.costExpenseId);
      } else {
        Store.updateExpense(refundFromCostContext.costExpenseId, {
          amount: remainingCostAmount
        });
      }

      clearExpenseRefundFromCostContext();
      typeSelect.disabled = false;
      document.getElementById('expense-name').readOnly = false;
      const expenseNameSuggestionsButton = document.getElementById('btn-expense-name-suggestions');
      if (expenseNameSuggestionsButton) expenseNameSuggestionsButton.disabled = false;
      const form = document.getElementById('expense-form');
      if (form) form.style.display = 'none';
      const tableContainer = document.querySelector('#expenses-view .table-container');
      if (tableContainer) tableContainer.style.display = 'block';
      return;
    }

    if (id) {
      Store.updateExpense(id, { type, date, name, amount, paidById, advanceForId, dietaCalculationMode, dietaByActiveDays, dietDaysAdjustment, dietaDaysCount });
    } else {
      Store.addExpense({ type, date, name, amount, paidById, advanceForId, dietaCalculationMode, dietaByActiveDays, dietDaysAdjustment, dietaDaysCount });
    }
    const form = document.getElementById('expense-form');
    if (form) form.style.display = 'none';
    const tableContainer = document.querySelector('#expenses-view .table-container');
    if (tableContainer) tableContainer.style.display = 'block';
  });

  updateExpenseFormTypeUI(typeSelect.value);
}

function renderExpenses() {
  const state = Store.getState();

  document.getElementById('expense-form').style.display = 'none';
  const tableContainer = document.querySelector('#expenses-view .table-container');
  if (tableContainer) tableContainer.style.display = 'block';

  const selectedMonth = getSelectedMonthKey();
  const currentExpenseType = document.getElementById('expense-type')?.value || 'COST';

  populateExpensePaidBySelect();
  populateExpenseRecipientSelect(currentExpenseType);

  const tbody = document.getElementById('expenses-table-body');
  tbody.innerHTML = '';

  const expensesForMonth = state.expenses.filter(expense => expense.date && expense.date.startsWith(`${selectedMonth}-`));
  const expenseDateCounts = expensesForMonth.reduce((counts, expense) => {
    const key = expense?.date || '';
    counts[key] = (counts[key] || 0) + 1;
    return counts;
  }, {});

  if (expensesForMonth.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;color:var(--text-muted)">Brak kosztów, zaliczek, zwrotów, premii i diet.</td></tr>`;
    return;
  }

  // Sort by date, but keep manual order inside the same day
  const expenseOrder = new Map((state.expenses || []).map((expense, index) => [expense.id, index]));
  const sortedExpenses = [...expensesForMonth].sort((a, b) => {
    const dateDiff = new Date(a.date) - new Date(b.date);
    if (dateDiff !== 0) return dateDiff;
    return (expenseOrder.get(a.id) ?? 0) - (expenseOrder.get(b.id) ?? 0);
  });

  const getPayerName = (id) => {
    if (!id) return 'Nieznany';
    if (id === BONUS_EXPENSE_PAYER_ID) return 'Od Wszystkich Wspólników';
    if (id === DIETA_EXPENSE_PAYER_ID) return 'Z zysku pracownika';
    if (id.startsWith('client_')) {
      const cId = id.substring(7);
      return (state.clients.find(c => c.id === cId)?.name || 'Nieznany') + ' (Klient)';
    }
    return getPersonDisplayName(state.persons.find(p => p.id === id)) || 'Nieznany';
  };

  const getPersonName = (id) => getPersonDisplayName(state.persons.find(p => p.id === id)) || 'Nieznany';

  const getExpenseBadgeMeta = (type) => {
    if (type === 'REFUND') {
      return { badgeClass: 'badge-refund', label: 'Zwrot' };
    }
    if (type === 'BONUS') {
      return { badgeClass: 'badge-bonus', label: 'Premia' };
    }
    if (type === 'DIETA') {
      return { badgeClass: 'badge-bonus', label: 'Dieta' };
    }
    if (type === 'ADVANCE') {
      return { badgeClass: 'badge-employee', label: 'Zaliczka' };
    }
    return { badgeClass: 'badge-partner', label: 'Koszt' };
  };

  const getExpenseResolvedAmount = (expense) => Calculations.getExpenseEffectiveAmount(expense, state, selectedMonth);

  const getDietaPayerLabel = (expense) => {
    if (expense.type !== 'DIETA') return getPayerName(expense.paidById);
    return getSettlementDietaMeta(expense, state, selectedMonth);
  };

  const getExpenseRecipientLabel = (expense) => {
    if (expense.type === 'ADVANCE' || expense.type === 'BONUS' || expense.type === 'DIETA' || expense.type === 'REFUND') {
      return getExpenseRecipientName(expense, state);
    }
    return '-';
  };

  sortedExpenses.forEach((e, idx) => {
    const badgeMeta = getExpenseBadgeMeta(e.type);
    const resolvedAmount = getExpenseResolvedAmount(e);
    const expenseRecipientLabel = getExpenseRecipientLabel(e);
    const expenseTitle = e.type === 'ADVANCE'
      ? 'Zaliczka'
      : (e.type === 'BONUS'
          ? 'Premia'
          : (e.type === 'DIETA'
              ? 'Dieta'
              : (e.type === 'REFUND' ? `Zwrot za ${e.name || 'Zwrot'}` : e.name)));
    const mobileExpenseTitle = e.type === 'REFUND'
      ? `${expenseTitle} dla ${expenseRecipientLabel}`
      : expenseTitle;
    const canReorderExpense = (expenseDateCounts[e.date || ''] || 0) > 1;
    const canCreateRefundFromCost = e.type === 'COST'
      && !!e.paidById
      && !e.paidById.startsWith('client_')
      && (parseFloat(e.amount) || 0) > 0
      && (state.persons || []).some(person => person.id !== e.paidById);
    const tr = document.createElement('tr');
    tr.dataset.id = e.id;
    tr.dataset.sortDate = e.date || '';
    tr.dataset.expenseType = e.type || 'COST';
    tr.innerHTML = `
      <td>${e.date}</td>
      <td style="font-weight: 500"><span class="expense-title-desktop">${expenseTitle}</span><span class="expense-title-mobile">${mobileExpenseTitle}</span></td>
      <td>
        <span class="badge ${badgeMeta.badgeClass}">
          ${badgeMeta.label}
        </span>
      </td>
      <td style="color: ${e.type === 'ADVANCE' || e.type === 'REFUND' ? 'var(--warning)' : ((e.type === 'BONUS' || e.type === 'DIETA') ? '#fbbf24' : 'var(--danger)')}">
        ${resolvedAmount.toFixed(2)} zł
      </td>
      <td>${e.type === 'BONUS' ? 'Od Wszystkich Wspólników' : getDietaPayerLabel(e)}</td>
      <td>${expenseRecipientLabel}</td>
      <td>
        <div class="expense-actions-row" style="display: flex; gap: 0.5rem; align-items: center; width: 100%;">
          <span class="expense-mobile-date" aria-hidden="true">
            <span class="expense-mobile-date-label">Data:</span>
            <span class="expense-mobile-date-value">${e.date}</span>
          </span>
          <span class="expense-actions-buttons" style="display: flex; gap: 0.5rem; margin-left: auto;">
            <button class="btn btn-secondary btn-icon btn-edit-expense" data-id="${e.id}">
              <i data-lucide="edit-2" style="width:16px;height:16px"></i>
            </button>
            <button class="btn btn-danger btn-icon btn-delete-expense" data-id="${e.id}">
              <i data-lucide="trash-2" style="width:16px;height:16px"></i>
            </button>
            ${canCreateRefundFromCost ? `<button class="btn btn-secondary btn-icon btn-create-refund-from-cost" type="button" data-id="${e.id}" title="Dodaj zwrot z tego kosztu" aria-label="Dodaj zwrot z tego kosztu"><i data-lucide="corner-down-left" style="width:16px;height:16px"></i></button>` : ''}
            ${canReorderExpense ? `<button class="btn btn-secondary btn-icon drag-handle btn-reorder-handle" type="button" title="Przytrzymaj, aby zmienić kolejność wpisów z tej samej daty" aria-label="Przytrzymaj, aby zmienić kolejność wpisów z tej samej daty">
              <i data-lucide="grip-horizontal" style="width:16px;height:16px"></i>
            </button>` : ''}
          </span>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
    // Insert mobile-only divider row between entries (no spacing/margins)
    if (idx !== sortedExpenses.length - 1) {
      const dividerTr = document.createElement('tr');
      dividerTr.className = 'expense-mobile-divider-row';
      dividerTr.innerHTML = `<td colspan="7" class="expense-mobile-divider-cell"><div class="expense-mobile-divider" aria-hidden="true"></div></td>`;
      tbody.appendChild(dividerTr);
    }
  });

  lucide.createIcons();
  setupListSortable(document.getElementById('expenses-table-body'), {
    onMove: (evt) => allowSortOnlyWithinSameDate(evt),
    onEnd: (evt) => {
      const date = evt.item?.dataset?.sortDate || '';
      if (!date) return;
      const selector = `tr[data-id][data-sort-date="${date}"]`;
      Store.reorderExpensesForDate(date, getSortableRowIds(document.getElementById('expenses-table-body'), selector));
    }
  });

  document.querySelectorAll('.btn-edit-expense').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = e.currentTarget.getAttribute('data-id');
      const expense = state.expenses.find(x => x.id === id);
      if (expense) {
        clearExpenseRefundFromCostContext();
        const expenseNameSuggestionsButton = document.getElementById('btn-expense-name-suggestions');
        if (expenseNameSuggestionsButton) expenseNameSuggestionsButton.disabled = false;
        document.getElementById('expense-form-title').textContent = 'Edytuj Wpis';
        document.getElementById('expense-id').value = expense.id;
        document.getElementById('expense-date').value = expense.date;
        document.getElementById('expense-name').value = expense.name;
        document.getElementById('expense-amount').value = expense.amount;
        document.getElementById('expense-paid-by').value = expense.paidById;
        document.getElementById('expense-refund-all-partners').checked = expense.type === 'REFUND' && isExpenseSharedWithAllPartners(expense);
        setExpenseSelectedDietaMode(getExpenseDietaModeFromExpense(expense));
        document.getElementById('expense-dieta-days-adjustment').value = getExpenseDietaModeFromExpense(expense) === EXPENSE_DIETA_MODE_MANUAL_DAYS
          ? formatExpenseManualDaysInput(expense.dietaDaysCount ?? 0)
          : formatSignedIntegerInput(expense.dietDaysAdjustment ?? 0);
        const typeSelect = document.getElementById('expense-type');
        typeSelect.value = expense.type;
        updateExpenseFormTypeUI(expense.type, expense.type === 'REFUND' && isExpenseSharedWithAllPartners(expense) ? '' : (expense.advanceForId || ''));
        if (expense.type !== 'BONUS' && expense.type !== 'DIETA') {
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

  document.querySelectorAll('.btn-create-refund-from-cost').forEach(btn => {
    btn.addEventListener('click', (event) => {
      const id = event.currentTarget.getAttribute('data-id');
      const expense = state.expenses.find(item => item.id === id && item.type === 'COST');
      if (!expense) return;
      openExpenseRefundFromCostForm(expense);
    });
  });
}

// ==========================================
// SETTLEMENT VIEW
// ==========================================
function formatSettlementCurrency(value) {
  return formatPolishCurrencyWithSuffix(value, 'zł', 2, 2);
}

function formatSettlementHours(value) {
  const amount = parseFloat(value);
  return `${(Number.isFinite(amount) ? amount : 0).toFixed(1)}h`;
}

function formatSettlementCompactCurrency(value) {
  return formatPolishCurrencyWithSuffix(value, 'zł', 2, 2);
}

function getSettlementDietaLabel(expense, state, month = getSelectedMonthKey()) {
  if (!expense || expense.type !== 'DIETA') return 'Dieta';
  if (!Calculations.isDietaCountedByActiveDays(expense) && !Calculations.isDietaCountedByManualDays(expense)) return 'Dieta';

  const effectiveDays = Calculations.getExpenseEffectiveDays(expense, state, month);
  const dayRate = parseFloat(expense.amount) || 0;
  return `Dieta (za ${effectiveDays} dni po ${formatSettlementCompactCurrency(dayRate)}/dzień)`;
}

function getSettlementDietaMeta(expense, state, month = getSelectedMonthKey()) {
  if (!expense || expense.type !== 'DIETA') return 'stała kwota';
  if (!Calculations.isDietaCountedByActiveDays(expense) && !Calculations.isDietaCountedByManualDays(expense)) return 'stała kwota';

  const effectiveDays = Calculations.getExpenseEffectiveDays(expense, state, month);
  const dayRate = parseFloat(expense.amount) || 0;
  return `za ${effectiveDays} dni po ${formatSettlementCompactCurrency(dayRate)}/dzień`;
}

function getSettlementPersonDietaSummaryLabel(personId, expenses, state, month = getSelectedMonthKey()) {
  const personDietas = (expenses || []).filter(expense => expense.type === 'DIETA' && expense.advanceForId === personId);
  if (personDietas.length === 1) {
    return getSettlementDietaLabel(personDietas[0], state, month);
  }
  return 'Dieta';
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
  if (id === ALL_PARTNERS_EXPENSE_TARGET_ID) return 'Wszyscy wspólnicy';
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
      if (!day?.hours) return;
      if (isMonthlySheetPersonInactiveOnDay(sheet, person.id, parseInt(dayKey, 10))) return;
      if (day.hours[person.id] !== undefined && day.hours[person.id] !== '') {
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
      amount: Calculations.getExpenseEffectiveAmount(expense, state),
      meta: 'Premia od wszystkich wspólników'
    }));

  const dietasReceived = selectedExpenses
    .filter(expense => expense.type === 'DIETA' && expense.advanceForId === personId)
    .map(expense => ({
      label: `${expense.date} • ${getSettlementDietaLabel(expense, state)}`,
      amount: Calculations.getExpenseEffectiveAmount(expense, state),
      meta: getSettlementDietaMeta(expense, state)
    }));

  const refundsPaid = selectedExpenses
    .filter(expense => expense.type === 'REFUND' && expense.paidById === personId)
    .map(expense => ({
      label: `${expense.date} • Zwrot za ${expense.name || 'Zwrot'}`,
      shortLabel: `Zwrot za ${expense.name || 'Zwrot'}`,
      amount: Calculations.getExpenseEffectiveAmount(expense, state),
      meta: isExpenseSharedWithAllPartners(expense)
        ? 'Zwrot dla wszystkich wspólników'
        : `Zwrot dla ${getSettlementEntityName(state, expense.advanceForId)}`
    }));

  const refundsReceived = selectedExpenses
    .filter(expense => expense.type === 'REFUND')
    .map(expense => {
      const amount = Calculations.getRefundReceivedAmountForPerson
        ? Calculations.getRefundReceivedAmountForPerson(expense, personId, state, getSelectedMonthKey())
        : 0;
      if (!(amount > 0)) return null;

      const payerName = getSettlementEntityName(state, expense.paidById);
      return {
        label: `${expense.date} • Zwrot za ${expense.name || 'Zwrot'} od ${payerName}`,
        shortLabel: `Zwrot za ${expense.name || 'Zwrot'} od ${payerName}`,
        amount,
        meta: isExpenseSharedWithAllPartners(expense)
          ? 'Kwota podzielona po równo między wszystkich wspólników'
          : 'Zwrot od konkretnej osoby'
      };
    })
    .filter(Boolean);

  return { paidCosts, paidAdvances, advancesTaken, bonusesReceived, dietasReceived, refundsPaid, refundsReceived };
}

function buildSettlementRefundRowsHtml(refundsReceived = [], refundsPaid = []) {
  return [
    ...refundsReceived.map(item => `
      <div class="settlement-detail-row"><span>${item.shortLabel}</span><strong class="settlement-accent-positive">+${formatSettlementCurrency(item.amount)}</strong></div>
    `),
    ...refundsPaid.map(item => `
      <div class="settlement-detail-row"><span>${item.shortLabel}</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(item.amount)}</strong></div>
    `)
  ].join('');
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
            ? `${formatSettlementHours(metrics.totalLoggedHours)} × ${replaceCurrencySuffix(formatSettlementCurrency(metrics.clientRate), 'zł/h')}`
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
  const totalEmployeeDietas = res.employees.reduce((sum, employee) => sum + (employee.dietaAmount || 0), 0);
  const totalCommonCosts = selectedExpenses
    .filter(expense => expense.type === 'COST')
    .reduce((sum, expense) => sum + expense.amount, 0);
  const clientAdvanceItems = selectedExpenses.filter(expense => expense.type === 'ADVANCE' && expense.paidById && expense.paidById.startsWith('client_'));
  const otherAdvanceItems = selectedExpenses.filter(expense => expense.type === 'ADVANCE' && (!expense.paidById || !expense.paidById.startsWith('client_')));
  const bonusItems = selectedExpenses.filter(expense => expense.type === 'BONUS');
  const dietaItems = selectedExpenses.filter(expense => expense.type === 'DIETA');
  const refundItems = selectedExpenses.filter(expense => expense.type === 'REFUND');

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
    totalEmployeeDietas,
    totalGrossPayouts: [...res.partners, ...(res.separateCompanies || []), ...res.workingPartners, ...res.employees].reduce((sum, person) => sum + (person.toPayout || 0), 0),
    totalCommonCosts,
    activeCostShareCount,
    costShare: activeCostShareCount > 0 ? totalCommonCosts / activeCostShareCount : 0,
    clientAdvanceItems,
    otherAdvanceItems,
    bonusItems,
    dietaItems,
    refundItems,
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

function buildSettlementEmployeePayoutsHtml(personId, state, payoutMonth) {
  try {
    const payoutSettings = Calculations.getPayoutSettings(state, payoutMonth);
    const record = payoutSettings?.employees?.[personId];
    if (!record || !Array.isArray(record.payouts) || record.payouts.length === 0) return '';

    const totalCash = record.payouts.reduce((s, p) => s + (p.cashAmount || 0), 0);
    const totalAdv = record.payouts.reduce((s, p) => s + p.deductedAdvances.filter(a => !a.restoredToCosts).reduce((ss, a) => ss + a.amount, 0), 0);
    const total = totalCash + totalAdv;

    const allPersons = state?.common?.persons || state?.persons || [];
    const entriesHtml = record.payouts.map(p => {
      const advTotal = p.deductedAdvances.filter(a => !a.restoredToCosts).reduce((s, a) => s + a.amount, 0);
      const entryTotal = (p.cashAmount || 0) + advTotal;
      const typeLabel = p.type === 'weekly' ? 'Tygodniówka' : p.type === 'custom' ? 'Niestandardowa' : 'Miesięczna';
      const label = p.label || typeLabel;
      const payer = p.paidByPartnerId ? allPersons.find(x => x.id === p.paidByPartnerId) : null;
      const payerName = payer ? getPersonDisplayName(payer) : '';
      const dateStr = p.payoutDate ? ` • ${p.payoutDate.split('-').reverse().join('.')}` : '';
      return `<div class="settlement-payout-entry">
        <span>${escapeReportHtml(label)}${escapeReportHtml(dateStr)}${payerName ? ` • ${escapeReportHtml(payerName)}` : ''}</span>
        <strong>${formatSettlementCurrency(entryTotal)}</strong>
      </div>`;
    }).join('');

    return `
      <div class="settlement-detail-item settlement-payout-section">
        <div class="settlement-detail-item-header">
          <div><h4>Wypłaty w tym miesiącu</h4></div>
          <strong>${formatSettlementCurrency(total)}</strong>
        </div>
        <div>${entriesHtml}</div>
      </div>`;
  } catch (e) {
    return '';
  }
}

function buildSettlementPersonCardHtml(entry, role, state, details) {
  const personSources = getSettlementPersonSources(entry.person, state);
  const expenseDetails = getSettlementPersonExpenseDetails(entry.person.id, details.selectedExpenses, state);
  const employeeGeneratedProfit = role === 'employee' ? getEmployeeGeneratedProfitDisplay(entry) : null;
  const ownGrossAmount = parseFloat(entry.ownGrossAmount) || 0;
  const totalTaxAmount = parseFloat(entry.taxAmount) || 0;
  const sharedCompanyTaxAmount = parseFloat(entry.sharedCompanyTaxAmount) || 0;
  const bonusAmount = parseFloat(entry.bonusAmount) || 0;
  const dietaAmount = parseFloat(entry.dietaAmount) || 0;
  const refundsReceivedAmount = parseFloat(entry.refundsReceived) || 0;
  const refundsPaidAmount = parseFloat(entry.refundsPaid) || 0;
  const dietaSummaryLabel = getSettlementPersonDietaSummaryLabel(entry.person.id, details.selectedExpenses, state, details.selectedMonth);
  const sourcesHtml = personSources.sources.length > 0
    ? personSources.sources.map(source => `<li>${source.label}: ${formatSettlementHours(source.hours)} × ${replaceCurrencySuffix(formatSettlementCurrency(source.rate), 'zł/h')} = ${formatSettlementCurrency(source.salary)}</li>`).join('')
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
    grossFormula += ` + ${formatSettlementCurrency(entry.paidCosts)} + ${formatSettlementCurrency(entry.paidAdvances)}${refundsReceivedAmount > 0 ? ` + ${formatSettlementCurrency(refundsReceivedAmount)}` : ''}${refundsPaidAmount > 0 ? ` - ${formatSettlementCurrency(refundsPaidAmount)}` : ''} - ${formatSettlementCurrency(entry.costShareApplied || 0)} - ${formatSettlementCurrency(entry.advancesTaken)}`;
    grossText = `Przychód (Brutto) = przychód własny(Brutto) + zwrot kosztów + zwrot zaliczek${refundsReceivedAmount > 0 ? ' + otrzymane zwroty' : ''}${refundsPaidAmount > 0 ? ' - wypłacone zwroty' : ''} - udział w kosztach - pobrane zaliczki`;
    netFormulaHtml = `<div style="margin-top: 0.45rem;">Netto = ${formatSettlementCurrency(entry.toPayout)} - podatek własny ${formatSettlementCurrency(entry.ownTaxAmount || 0)}${sharedCompanyTaxAmount > 0 ? ` - podatek wspólny ${formatSettlementCurrency(sharedCompanyTaxAmount)}` : ''} - ZUS ${formatSettlementCurrency(entry.zusAmount)} = <strong>${formatSettlementCurrency(entry.netAfterAccounting)}</strong></div>`;
  } else if (role === 'separateCompany') {
    ownGrossFormula = `${formatSettlementCurrency(entry.salary)} + ${formatSettlementCurrency(entry.revenueShare)} + ${formatSettlementCurrency(entry.worksShare)}`;
    ownGrossText = 'Przychód własny(Brutto) = własne godziny + udział z godzin pracowników + udział z wykonanych prac';
    grossFormula += ` + ${formatSettlementCurrency(entry.paidCosts)} + ${formatSettlementCurrency(entry.paidAdvances)}${refundsReceivedAmount > 0 ? ` + ${formatSettlementCurrency(refundsReceivedAmount)}` : ''}${refundsPaidAmount > 0 ? ` - ${formatSettlementCurrency(refundsPaidAmount)}` : ''} - ${formatSettlementCurrency(entry.costShareApplied || 0)} - ${formatSettlementCurrency(entry.advancesTaken)}`;
    grossText = `Przychód (Brutto) = przychód własny(Brutto) + zwrot kosztów + zwrot zaliczek${refundsReceivedAmount > 0 ? ' + otrzymane zwroty' : ''}${refundsPaidAmount > 0 ? ' - wypłacone zwroty' : ''} - udział w kosztach - pobrane zaliczki`;
  } else if (role === 'workingPartner') {
    ownGrossFormula = `${formatSettlementCurrency(entry.salary)} + ${formatSettlementCurrency(entry.worksShare)}`;
    ownGrossText = 'Przychód własny(Brutto) = własne godziny + udział z wykonanych prac';
    grossFormula += ` + ${formatSettlementCurrency(entry.paidCosts)} + ${formatSettlementCurrency(entry.paidAdvances)}${refundsReceivedAmount > 0 ? ` + ${formatSettlementCurrency(refundsReceivedAmount)}` : ''}${refundsPaidAmount > 0 ? ` - ${formatSettlementCurrency(refundsPaidAmount)}` : ''} - ${formatSettlementCurrency(entry.costShareApplied || 0)} - ${formatSettlementCurrency(entry.advancesTaken)}`;
    grossText = `Przychód (Brutto) = przychód własny(Brutto) + zwrot kosztów + zwrot zaliczek${refundsReceivedAmount > 0 ? ' + otrzymane zwroty' : ''}${refundsPaidAmount > 0 ? ' - wypłacone zwroty' : ''} - udział w kosztach - pobrane zaliczki`;
    const deductions = [];
    if ((entry.ownTaxAmount || 0) > 0) deductions.push(`- podatek własny ${formatSettlementCurrency(entry.ownTaxAmount || 0)}`);
    if ((entry.deductedContractTaxAmount || 0) > 0) deductions.push(`- podatek UZ ${formatSettlementCurrency(entry.deductedContractTaxAmount || 0)}`);
    if ((entry.deductedContractZusAmount || 0) > 0) deductions.push(`- ZUS UZ ${formatSettlementCurrency(entry.deductedContractZusAmount || 0)}`);
    netFormulaHtml = `<div style="margin-top: 0.45rem;">Netto = ${formatSettlementCurrency(entry.toPayout)}${deductions.length ? ` ${deductions.join(' ')}` : ''} = <strong>${formatSettlementCurrency(entry.netAfterAccounting)}</strong></div>`;
  } else {
    if (bonusAmount > 0) {
      grossFormula += ` + ${formatSettlementCurrency(bonusAmount)}`;
    }
    if (dietaAmount > 0) {
      grossFormula += ` + ${formatSettlementCurrency(dietaAmount)}`;
    }
    grossFormula += ` + ${formatSettlementCurrency(entry.paidCosts)} + ${formatSettlementCurrency(entry.paidAdvances)}${refundsReceivedAmount > 0 ? ` + ${formatSettlementCurrency(refundsReceivedAmount)}` : ''}${refundsPaidAmount > 0 ? ` - ${formatSettlementCurrency(refundsPaidAmount)}` : ''} - ${formatSettlementCurrency(entry.advancesTaken)}`;
    if ((entry.deductedContractTaxAmount || 0) > 0 || (entry.deductedContractZusAmount || 0) > 0) {
      grossFormula += `${(entry.deductedContractTaxAmount || 0) > 0 ? ` - ${formatSettlementCurrency(entry.deductedContractTaxAmount || 0)}` : ''}${(entry.deductedContractZusAmount || 0) > 0 ? ` - ${formatSettlementCurrency(entry.deductedContractZusAmount || 0)}` : ''}`;
      grossText = `Do wypłaty = zarobek${bonusAmount > 0 ? ' + premia' : ''}${dietaAmount > 0 ? ' + dieta' : ''} + zwroty${refundsReceivedAmount > 0 ? ' + otrzymane zwroty' : ''}${refundsPaidAmount > 0 ? ' - wypłacone zwroty' : ''} - pobrane zaliczki - podatek UZ - ZUS UZ`;
    } else {
      grossText = `Do wypłaty = zarobek${bonusAmount > 0 ? ' + premia' : ''}${dietaAmount > 0 ? ' + dieta' : ''} + zwroty${refundsReceivedAmount > 0 ? ' + otrzymane zwroty' : ''}${refundsPaidAmount > 0 ? ' - wypłacone zwroty' : ''} - pobrane zaliczki`;
    }
  }

  return `
    <div class="settlement-person-card">
      <div class="settlement-person-header">
        <div>
          <div class="settlement-person-group">${getSettlementPersonTypeLabel(entry.person.type)}</div>
          <h4>${getSettlementPersonDisplayNameHtml(entry.person, state)}</h4>
          <div class="settlement-person-summary">Godziny: ${formatSettlementHours(entry.hours)}${entry.effectiveRate > 0 ? ` • efektywna stawka ${replaceCurrencySuffix(formatSettlementCurrency(entry.effectiveRate), 'zł/h')}` : ''}</div>
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
        ${dietaAmount > 0 ? `<div class="settlement-detail-row"><span>${dietaSummaryLabel}</span><strong style="color: #fbbf24;">${formatSettlementCurrency(dietaAmount)}</strong></div>` : ''}
        <div class="settlement-detail-row"><span>Zwrot kosztów</span><strong>${formatSettlementCurrency(entry.paidCosts)}</strong></div>
        <div class="settlement-detail-row"><span>Zwrot wypłaconych zaliczek</span><strong>${formatSettlementCurrency(entry.paidAdvances)}</strong></div>
        ${buildSettlementRefundRowsHtml(expenseDetails.refundsReceived, expenseDetails.refundsPaid)}
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
              <h4>Dieta</h4>
            </div>
            <strong>${formatSettlementCurrency(dietaAmount)}</strong>
          </div>
          <ul class="settlement-inline-list">${expenseDetails.dietasReceived.length ? expenseDetails.dietasReceived.map(item => `<li>${item.label}: ${formatSettlementCurrency(item.amount)} <span style="color: var(--text-secondary);">(${item.meta})</span></li>`).join('') : '<li>Brak pozycji.</li>'}</ul>
        </div>
        <div class="settlement-detail-item">
          <div class="settlement-detail-item-header">
            <div>
              <h4>Zwroty otrzymane</h4>
            </div>
            <strong>${formatSettlementCurrency(refundsReceivedAmount)}</strong>
          </div>
          <ul class="settlement-inline-list">${expenseDetails.refundsReceived.length ? expenseDetails.refundsReceived.map(item => `<li>${item.label}: ${formatSettlementCurrency(item.amount)} <span style="color: var(--text-secondary);">(${item.meta})</span></li>`).join('') : '<li>Brak pozycji.</li>'}</ul>
        </div>
        <div class="settlement-detail-item">
          <div class="settlement-detail-item-header">
            <div>
              <h4>Zwroty wypłacone</h4>
            </div>
            <strong>${formatSettlementCurrency(refundsPaidAmount)}</strong>
          </div>
          <ul class="settlement-inline-list">${expenseDetails.refundsPaid.length ? expenseDetails.refundsPaid.map(item => `<li>${item.label}: ${formatSettlementCurrency(item.amount)} <span style="color: var(--text-secondary);">(${item.meta})</span></li>`).join('') : '<li>Brak pozycji.</li>'}</ul>
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
      ${role === 'employee' ? buildSettlementEmployeePayoutsHtml(entry.person.id, state, details.selectedMonth) : ''}
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
          <div style="text-align: right;">
            <strong>${formatSettlementCurrency(targetTaxAmount)}</strong>
            <div class="settlement-detail-meta" style="margin-top: 0.15rem;">Docelowy podatek do poniesienia</div>
          </div>
        </div>
        <div class="settlement-detail-stack" style="margin-top: 0.75rem;">
          <div class="settlement-detail-row"><span>Rzeczywiście wystawione faktury</span><strong>${formatSettlementCurrency(issuedAmount)}</strong></div>
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Opłaca w urzędzie od swoich faktur</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(officeTaxAmount)}</strong></div>
          <div class="settlement-detail-meta" style="margin-top: -0.1rem;">${formatSettlementCurrency(issuedAmount)} × ${taxRatePercent.toFixed(2)}% = ${formatSettlementCurrency(officeTaxAmount)}</div>
          ${separateCompanyReimbursements > 0 ? `<div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Zwrot podatku od osobnych firm</span><strong>${formatSettlementCurrency(separateCompanyReimbursements)}</strong></div>` : ''}
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Podatek własny</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(ownTaxAmount)}</strong></div>
          <div class="settlement-detail-meta" style="margin-top: -0.1rem;">Liczony tylko od pola „Przychód własny(Brutto)”.</div>
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Podatek wspólny</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(sharedCompanyTaxAmount)}</strong></div>
          <div class="settlement-detail-meta" style="margin-top: -0.1rem;">Równa część podatku spółki: zwrot zaliczek + pensje pracowników + zwrot Podatku i ZUS za pracowników.</div>
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Faktyczny zapłacony podatek</span><strong>${formatSettlementCurrency(actualTaxBurdenAmount)}</strong></div>
          <div class="settlement-detail-row" style="margin-top: 0.6rem;"><span>Docelowy podatek do poniesienia</span><strong>${formatSettlementCurrency(targetTaxAmount)}</strong></div>
          <div class="settlement-detail-meta" style="margin-top: -0.1rem;">(${formatSettlementCurrency(actualTaxBurdenAmount)} - ${formatSettlementCurrency(targetTaxAmount)} = ${formatSettlementCurrency(actualTaxBurdenAmount - targetTaxAmount)})</div>
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
    const totalEmployeeAccountingRefunds = partnerEntries.reduce((sum, entry) => sum + (parseFloat(entry.employeeAccountingRefund) || 0), 0);
    const totalEmployeeReceivables = partnerEntries.reduce((sum, entry) => sum + (parseFloat(entry.employeeReceivables) || 0), 0);
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
          <div class="settlement-detail-row"><span>Pensje pracowników</span><strong>${formatSettlementCurrency(totalEmployeeSalaries)}</strong></div>
          <div class="settlement-detail-row"><span>Zwrot Podatku i ZUS za pracowników</span><strong>${formatSettlementCurrency(totalEmployeeAccountingRefunds)}</strong></div>
          ${totalEmployeeReceivables > 0 ? `<div class="settlement-detail-row"><span>Do odebrania od pracowników</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(totalEmployeeReceivables)}</strong></div>` : ''}
        </div>
        <div class="settlement-detail-formula" style="margin-top: 0.85rem;">
          ${formatSettlementCurrency(totalAdvanceRefunds)} + ${formatSettlementCurrency(totalEmployeeSalaries)} + ${formatSettlementCurrency(totalEmployeeAccountingRefunds)}${totalEmployeeReceivables > 0 ? ` - ${formatSettlementCurrency(totalEmployeeReceivables)}` : ''} = <strong>${formatSettlementCurrency(taxBase)}</strong>
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
  const employeeDietas = details.totalEmployeeDietas || 0;
  const basicEmployeeProfitBeforeEmployerCharges = res.employeeRevenue - details.totalEmployeesSalaryFromHours;
  const basicEmployeeProfit = basicEmployeeProfitBeforeEmployerCharges - employeeBonuses - employeeDietas - employerPaidContractCharges;
  const profitFromEmployees = res.employeeProfitShared || 0;
  const extraEmployeeProfitShare = profitFromEmployees - basicEmployeeProfit;
  const accountingPeopleCount = res.partners.length + res.workingPartners.length;
  const clientAdvanceSum = details.clientAdvanceItems.reduce((sum, item) => sum + item.amount, 0);
  const otherAdvanceSum = details.otherAdvanceItems.reduce((sum, item) => sum + item.amount, 0);
  const totalBonusSum = details.bonusItems.reduce((sum, item) => sum + item.amount, 0);
  const totalDietaSum = details.dietaItems.reduce((sum, item) => sum + Calculations.getExpenseEffectiveAmount(item, state, details.selectedMonth), 0);
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
          ${employeeDietas > 0 ? `<div class="settlement-detail-row"><span>Diety pracowników</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(employeeDietas)}</strong></div>` : ''}
          ${employerPaidContractCharges > 0 ? `<div class="settlement-detail-row"><span>Podatek i ZUS UZ opłacony przez pracodawcę</span><strong class="settlement-accent-negative">-${formatSettlementCurrency(employerPaidContractCharges)}</strong></div>` : ''}
          ${Math.abs(extraEmployeeProfitShare) >= 0.005 ? `<div class="settlement-detail-row"><span>Dodatkowy podział zysków firm i ich pracowników</span><strong>${formatSettlementCurrency(extraEmployeeProfitShare)}</strong></div>` : ''}
        </div>
        <div class="settlement-detail-formula">${formatSettlementCurrency(res.employeeRevenue)} - ${formatSettlementCurrency(details.totalEmployeesSalaryFromHours)}${employeeBonuses > 0 ? ` - ${formatSettlementCurrency(employeeBonuses)}` : ''}${employeeDietas > 0 ? ` - ${formatSettlementCurrency(employeeDietas)}` : ''}${employerPaidContractCharges > 0 ? ` - ${formatSettlementCurrency(employerPaidContractCharges)}` : ''}${Math.abs(extraEmployeeProfitShare) >= 0.005 ? ` ${extraEmployeeProfitShare >= 0 ? '+' : '-'} ${formatSettlementCurrency(Math.abs(extraEmployeeProfitShare))}` : ''} = <strong>${formatSettlementCurrency(profitFromEmployees)}</strong></div>
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
        <h3>6. Zaliczki, premie i diety</h3>
        <div class="settlement-detail-stack">
          <div class="settlement-detail-row"><span>Suma zaliczek od klientów</span><strong>${formatSettlementCurrency(clientAdvanceSum)}</strong></div>
          <div class="settlement-detail-row"><span>Suma pozostałych zaliczek</span><strong>${formatSettlementCurrency(otherAdvanceSum)}</strong></div>
          <div class="settlement-detail-row"><span>Suma premii pracowników</span><strong>${formatSettlementCurrency(totalBonusSum)}</strong></div>
          <div class="settlement-detail-row"><span>Suma diet pracowników</span><strong>${formatSettlementCurrency(totalDietaSum)}</strong></div>
          <div class="settlement-detail-row"><span>Suma wszystkich zaliczek</span><strong>${formatSettlementCurrency(totalAdvanceSum)}</strong></div>
        </div>
        <div class="settlement-detail-formula">Zaliczki: ${formatSettlementCurrency(clientAdvanceSum)} + ${formatSettlementCurrency(otherAdvanceSum)} = <strong>${formatSettlementCurrency(totalAdvanceSum)}</strong>${totalBonusSum > 0 ? `<div style="margin-top: 0.45rem;">Premie pracowników: <strong>${formatSettlementCurrency(totalBonusSum)}</strong></div>` : ''}${totalDietaSum > 0 ? `<div style="margin-top: 0.45rem;">Diety pracowników: <strong>${formatSettlementCurrency(totalDietaSum)}</strong></div>` : ''}</div>
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
            description: `${formatSettlementHours(item.hours)} × ${replaceCurrencySuffix(formatSettlementCurrency(item.clientRate), 'zł/h')}`
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
      <h3>Koszty, zaliczki i zwroty z miesiąca ${formatMonthLabel(details.selectedMonth)}</h3>
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
                amount: Calculations.getExpenseEffectiveAmount(item, state, details.selectedMonth),
                meta: `Dla: ${getSettlementEntityName(state, item.advanceForId)} • płaci: Od Wszystkich Wspólników`
              })),
              'amount',
              'Brak premii w wybranym miesiącu.'
            )}
          </div>
        </div>
        <div>
          <h4 style="margin-bottom: 0.75rem;">Diety</h4>
          <div class="settlement-detail-list">
            ${buildSettlementEntryListHtml(
              details.dietaItems.map(item => ({
                label: `${item.date} • ${getSettlementDietaLabel(item, state, details.selectedMonth)}`,
                amount: Calculations.getExpenseEffectiveAmount(item, state, details.selectedMonth),
                meta: `Dla: ${getSettlementEntityName(state, item.advanceForId)} • ${getSettlementDietaMeta(item, state, details.selectedMonth)}`
              })),
              'amount',
              'Brak diet w wybranym miesiącu.'
            )}
          </div>
        </div>
        <div>
          <h4 style="margin-bottom: 0.75rem;">Zwroty</h4>
          <div class="settlement-detail-list">
            ${buildSettlementEntryListHtml(
              details.refundItems.map(item => ({
                label: `${item.date} • Zwrot za ${item.name || 'Zwrot'}`,
                amount: Calculations.getExpenseEffectiveAmount(item, state, details.selectedMonth),
                meta: `Zwraca: ${getSettlementEntityName(state, item.paidById)} • dla: ${getExpenseRecipientName(item, state)}`
              })),
              'amount',
              'Brak zwrotów w wybranym miesiącu.'
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
  const selectedSettlementExpenses = (state.expenses || []).filter(expense => expense.date && expense.date.startsWith(`${selectedMonth}-`));
  const getSettlementRefundRowsForBlock = (personId = '') => {
    const expenseDetails = getSettlementPersonExpenseDetails(personId, selectedSettlementExpenses, state);
    return buildSettlementRefundRowsHtml(expenseDetails.refundsReceived, expenseDetails.refundsPaid);
  };
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

  const getSeparateCompanyEmployeeProfitLabel = (companyPerson) => {
    if (companyPerson?.sharesEmployeeProfits === true || companyPerson?.receivesCompanyEmployeeProfits === true) {
      return 'Podział zysku z pracowników:';
    }
    return 'Zysk z pracowników:';
  };

  const renderPartnerBlockHtml = (p) => `
        <div style="margin-bottom: 1.5rem; padding-bottom: 1rem; border-bottom: 1px dashed var(--border-color);">
          <h4 class="settlement-person-name--large">${getSettlementPersonDisplayNameHtml(p.person, state)}</h4>
          <div style="display: flex; flex-direction: column; gap: 0.3rem; font-size: 0.9rem; margin-top: 0.5rem;">
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Godziny:</span>
              <span>${p.hours.toFixed(1)}h</span>
            </div>
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Zarobek z własnych godzin:</span>
              <span>${p.salary.toFixed(2)} zł</span>
            </div>
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">Podział zysku z pracowników:</span>
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
            ${getSettlementRefundRowsForBlock(p.person.id)}
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
              <span style="color: var(--text-secondary);">Zarobek z własnych godzin:</span>
              <span>${company.salary.toFixed(2)} zł</span>
            </div>
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">${getSeparateCompanyEmployeeProfitLabel(company.person)}</span>
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
            ${getSettlementRefundRowsForBlock(company.person.id)}
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
            ${getSettlementRefundRowsForBlock(wp.person.id)}
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
    const dietaSummaryLabel = getSettlementPersonDietaSummaryLabel(e.person.id, state.expenses, state, getSelectedMonthKey());
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
            ${(e.dietaAmount || 0) > 0 ? `
            <div style="display: flex; justify-content: space-between;">
              <span style="color: var(--text-secondary);">${dietaSummaryLabel}:</span>
              <span style="color: #fbbf24;">+${(e.dietaAmount || 0).toFixed(2)} zł</span>
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
            ${getSettlementRefundRowsForBlock(e.person.id)}
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

function formatInvoiceIssuedAt(issuedAt) {
  if (!issuedAt) return '';
  const date = new Date(issuedAt);
  if (Number.isNaN(date.getTime())) return '';
  return date.toLocaleString('pl-PL');
}

function buildIssuedInvoiceSnapshot(invoiceData) {
  if (!invoiceData || typeof invoiceData !== 'object') return null;
  return {
    month: invoiceData.month || '',
    issueDate: invoiceData.issueDate || '',
    issuers: Array.isArray(invoiceData.issuers) ? invoiceData.issuers : [],
    clientInvoices: Array.isArray(invoiceData.clientInvoices) ? invoiceData.clientInvoices : [],
    extraInvoices: Array.isArray(invoiceData.extraInvoices) ? invoiceData.extraInvoices : [],
    issuerSummaries: Array.isArray(invoiceData.issuerSummaries) ? invoiceData.issuerSummaries : [],
    totalRevenue: parseFloat(invoiceData.totalRevenue) || 0,
    totalInvoices: parseFloat(invoiceData.totalInvoices) || 0,
    difference: Number.isFinite(parseFloat(invoiceData.difference))
      ? parseFloat(invoiceData.difference)
      : ((parseFloat(invoiceData.totalInvoices) || 0) - (parseFloat(invoiceData.totalRevenue) || 0)),
    emailText: (invoiceData.emailText || '').toString()
  };
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
  const issuedToggleButton = document.getElementById('btn-toggle-invoices-issued');
  const extraForm = document.getElementById('invoice-extra-form');
  const extraClientSelect = document.getElementById('invoice-extra-client-select');
  const extraClientNameInput = document.getElementById('invoice-extra-client-name');
  const extraCancelButton = document.getElementById('btn-cancel-invoice-extra');
  const extraPanelToggleButton = document.getElementById('btn-toggle-invoice-extra-panel');

  if (issueDateInput) {
    issueDateInput.addEventListener('change', () => {
      if (Store.getMonthSettings(getSelectedMonthKey())?.invoices?.issued === true) {
        alert('Faktury zostały oznaczone jako wystawione. Najpierw przywróć edycję faktur.');
        renderInvoices();
        return;
      }
      Store.updateInvoiceMonthConfig({ issueDate: issueDateInput.value, emailIntro: emailIntroInput ? emailIntroInput.value.trim() : '' });
    });
  }

  if (issuedToggleButton) {
    issuedToggleButton.addEventListener('click', () => {
      const selectedMonth = getSelectedMonthKey();
      const invoiceSettings = Store.getMonthSettings(selectedMonth)?.invoices || {};

      if (invoiceSettings.issued === true) {
        if (!confirm('Przywrócić edycję faktur dla tego miesiąca?')) return;
        Store.restoreInvoiceEditing(selectedMonth);
        renderInvoices();
        return;
      }

      const invoiceData = Calculations.calculateInvoices(Store.getState(), selectedMonth, { ignoreIssuedSnapshot: true });
      const issuedSnapshot = buildIssuedInvoiceSnapshot(invoiceData);
      if (!issuedSnapshot) return;

      if (!confirm('Oznaczyć faktury jako wystawione, zapamiętać ich bieżące wartości i zablokować edycję?')) return;

      Store.setInvoicesIssued(issuedSnapshot, selectedMonth);
      resetInvoiceExtraForm();
      renderInvoices();
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
      if (Store.getMonthSettings(getSelectedMonthKey())?.invoices?.issued === true) {
        alert('Faktury zostały oznaczone jako wystawione. Najpierw przywróć edycję faktur.');
        renderInvoices();
        return;
      }
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

      if (Store.getMonthSettings(getSelectedMonthKey())?.invoices?.issued === true) {
        alert('Faktury zostały oznaczone jako wystawione. Najpierw przywróć edycję faktur.');
        renderInvoices();
        return;
      }

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
  const issuedStatus = document.getElementById('invoices-issued-status');
  const totalRevenueEl = document.getElementById('invoices-total-revenue');
  const totalIssuedEl = document.getElementById('invoices-total-issued');
  const totalDifferenceEl = document.getElementById('invoices-total-difference');
  const issueDateInput = document.getElementById('invoices-issue-date');
  const emailIntroInput = document.getElementById('invoices-email-intro');
  const issuedToggleButton = document.getElementById('btn-toggle-invoices-issued');
  const configList = document.getElementById('invoices-config-list');
  const issuerSummary = document.getElementById('invoices-issuer-summary');
  const clientSummary = document.getElementById('invoices-client-summary');
  const emailText = document.getElementById('invoices-email-text');
  const extraClientSelect = document.getElementById('invoice-extra-client-select');
  const extraClientNameInput = document.getElementById('invoice-extra-client-name');
  const extraIssuerSelect = document.getElementById('invoice-extra-issuer');
  const extraAmountInput = document.getElementById('invoice-extra-amount');
  const extraSaveButton = document.getElementById('btn-save-invoice-extra');
  const extraCancelButton = document.getElementById('btn-cancel-invoice-extra');
  const extraList = document.getElementById('invoice-extra-list');
  const extraPanelToggleButton = document.getElementById('btn-toggle-invoice-extra-panel');
  if (!configList || !issuerSummary || !clientSummary || !emailText || !extraClientSelect || !extraIssuerSelect || !extraList || !extraPanelToggleButton || !issuedToggleButton) return;

  const state = Store.getState();
  const selectedMonth = getSelectedMonthKey();
  const monthSettings = Store.getMonthSettings(selectedMonth);
  const invoiceSettings = monthSettings?.invoices || { issueDate: '', emailIntro: '', issued: false, issuedAt: '', issuedSnapshot: null, clients: {}, extraInvoices: [] };
  const isInvoicesIssued = invoiceSettings.issued === true;
  const invoiceData = Calculations.calculateInvoices(state, selectedMonth);
  const extraEligibleIssuers = getInvoiceExtraEligibleIssuers(state);
  const issuedAtLabel = formatInvoiceIssuedAt(invoiceSettings.issuedAt);

  extraClientSelect.innerHTML = '<option value="">-- Wybierz klienta --</option>'
    + (state.clients || []).map(client => `<option value="${client.id}">${client.name}</option>`).join('');
  extraIssuerSelect.innerHTML = '<option value="">-- Wybierz wspólnika --</option>'
    + extraEligibleIssuers.map(person => `<option value="${person.id}">${getPersonDisplayName(person)}</option>`).join('');

  issuedToggleButton.textContent = isInvoicesIssued ? 'Przywróć edycję faktur' : 'Zostały wystawione';
  issuedToggleButton.className = isInvoicesIssued ? 'btn btn-secondary' : 'btn btn-primary';

  if (issuedStatus) {
    if (isInvoicesIssued) {
      issuedStatus.style.display = 'block';
      issuedStatus.textContent = `Faktury zostały oznaczone jako wystawione${issuedAtLabel ? ` (${issuedAtLabel})` : ''}. Pokazywane są zapamiętane wartości i edycja jest zablokowana.`;
    } else {
      issuedStatus.style.display = 'none';
      issuedStatus.textContent = '';
    }
  }

  if (subtitle) {
    subtitle.textContent = `Podział faktur i mail do księgowej dla ${formatMonthLabel(selectedMonth)}`;
  }

  if (totalRevenueEl) totalRevenueEl.textContent = Calculations.formatInvoiceCurrency(invoiceData.totalRevenue);
  if (totalIssuedEl) totalIssuedEl.textContent = Calculations.formatInvoiceCurrency(invoiceData.totalInvoices);
  if (totalDifferenceEl) {
    totalDifferenceEl.textContent = Calculations.formatInvoiceCurrency(invoiceData.difference);
    totalDifferenceEl.style.color = Math.abs(invoiceData.difference) < 0.005 ? 'var(--success)' : 'var(--danger)';
  }

  if (issueDateInput) {
    issueDateInput.value = invoiceData.issueDate;
    issueDateInput.disabled = isInvoicesIssued;
    updateMonthPickerRestrictions(issueDateInput, selectedMonth);
  }
  if (emailIntroInput) {
    emailIntroInput.value = invoiceSettings.emailIntro || '';
    emailIntroInput.disabled = isInvoicesIssued;
  }
  extraClientSelect.disabled = isInvoicesIssued;
  if (extraClientNameInput) extraClientNameInput.disabled = isInvoicesIssued;
  extraIssuerSelect.disabled = isInvoicesIssued;
  if (extraAmountInput) extraAmountInput.disabled = isInvoicesIssued;
  if (extraSaveButton) extraSaveButton.disabled = isInvoicesIssued;
  if (extraCancelButton) extraCancelButton.disabled = isInvoicesIssued;
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
          ${isInvoicesIssued
            ? '<span style="color: var(--warning); font-size: 0.82rem;">Edycja zablokowana</span>'
            : `<div style="display:flex; gap:0.5rem;">
                <button class="btn btn-secondary btn-icon btn-edit-invoice-extra" data-id="${invoice.id}">
                  <i data-lucide="edit-2" style="width:16px;height:16px"></i>
                </button>
                <button class="btn btn-danger btn-icon btn-delete-invoice-extra" data-id="${invoice.id}">
                  <i data-lucide="trash-2" style="width:16px;height:16px"></i>
                </button>
              </div>`}
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
    if (!isInvoicesIssued) {
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
    }

    lucide.createIcons();
    applyArchivedReadOnlyMode(state.isArchived === true);
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
    const frozenNote = isInvoicesIssued
      ? `<div style="margin-bottom: 0.85rem; color: var(--warning); font-size: 0.86rem;">Faktury zostały wystawione — pokazane są zapamiętane wartości, bez przeliczania.</div>`
      : '';

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
        ${frozenNote}
        <div style="display:grid; grid-template-columns: minmax(220px, 260px) 1fr; gap: 1rem; align-items:start; margin-bottom: 0.85rem;">
          <div class="form-group" style="margin-bottom: 0;">
            <label>Tryb podziału</label>
            <select class="invoice-client-mode" ${isInvoicesIssued ? 'disabled' : ''}>
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
            <input type="text" class="invoice-client-notes" value="${(clientConfig.notes || '').replace(/"/g, '&quot;')}" placeholder="Np. faktury na spółkę tylko od wybranych wspólników" ${isInvoicesIssued ? 'disabled' : ''}>
          </div>
        </div>
        <div style="margin-bottom: 0.85rem; display:flex; justify-content:space-between; gap:1rem; align-items:center; flex-wrap:wrap;">
          <label style="display:flex; align-items:center; gap:0.6rem; cursor:pointer; color: var(--text-secondary); margin: 0;">
            <input type="checkbox" class="invoice-deduct-advances-toggle" ${clientInvoice.deductClientAdvances ? 'checked' : ''} style="width:auto;" ${isInvoicesIssued ? 'disabled' : ''}>
            Odlicz zaliczki klienta od Przychodu klienta z arkuszy
          </label>
          ${(clientInvoice.clientAdvances || 0) > 0 ? `<div style="font-size:0.82rem; color: var(--text-secondary);">Zaliczki klienta w miesiącu: <strong style="color: var(--warning);">${Calculations.formatInvoiceCurrency(clientInvoice.clientAdvances || 0)}</strong></div>` : ''}
        </div>
        <div class="invoice-include-costs-row" style="margin-bottom: 0.85rem; display:${isSettlementRevenueMode ? 'flex' : 'none'}; justify-content:space-between; gap:1rem; align-items:center; flex-wrap:wrap;">
          <label style="display:flex; align-items:center; gap:0.6rem; cursor:pointer; color: var(--text-secondary); margin: 0;">
            <input type="checkbox" class="invoice-include-costs-toggle" ${clientInvoice.includeClientCosts !== false ? 'checked' : ''} style="width:auto;" ${isInvoicesIssued ? 'disabled' : ''}>
            Czy liczyć koszty w fakturach dla tego klienta
          </label>
          ${clientInvoice.includeClientCosts !== false && (clientInvoice.clientCostShareRatio || 0) > 0 ? `<div style="font-size:0.82rem; color: var(--text-secondary);">Udział kosztów w fakturach: <strong style="color: var(--warning);">${((clientInvoice.clientCostShareRatio || 0) * 100).toFixed(2)}%</strong></div>` : ''}
        </div>
        <div class="invoice-randomize-equal-split-row" style="margin-bottom: 0.85rem; display:${isEqualSplitMode ? 'flex' : 'none'}; justify-content:space-between; gap:1rem; align-items:center; flex-wrap:wrap;">
          <label style="display:flex; align-items:center; gap:0.6rem; cursor:pointer; color: var(--text-secondary); margin: 0;">
            <input type="checkbox" class="invoice-randomize-equal-split-toggle" ${clientInvoice.randomizeEqualSplitInvoices === true ? 'checked' : ''} style="width:auto;" ${isInvoicesIssued ? 'disabled' : ''}>
            Zróżnicuj faktury o:
          </label>
          <div style="display:flex; align-items:center; gap:0.5rem;">
            <input type="number" step="0.01" min="0" class="invoice-equal-split-variance-amount" value="${Number.isFinite(parseFloat(clientInvoice.equalSplitVarianceAmount)) ? parseFloat(clientInvoice.equalSplitVarianceAmount) : 10}" style="width:120px;" ${isInvoicesIssued ? 'disabled' : ''}>
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
                    <input type="checkbox" class="invoice-issuer-toggle" value="${issuer.id}" ${checked ? 'checked' : ''} style="width:auto;" ${isInvoicesIssued ? 'disabled' : ''}>
                    <span style="font-weight:600; color: var(--text-primary);">${getPersonDisplayName(issuer)}</span>
                  </label>
                  <span class="badge ${issuer.type === 'SEPARATE_COMPANY' ? 'badge-working-partner' : 'badge-partner'}">${issuer.type === 'SEPARATE_COMPANY' ? 'Osobna Firma' : 'Wspólnik'}</span>
                </span>
                <span style="font-size:0.8rem; color: var(--text-secondary);">Aktualnie wyliczone: ${Calculations.formatInvoiceCurrency(currentAllocation?.amount || 0)}</span>
                ${isPercentageMode && checked && !(issuer.type === 'SEPARATE_COMPANY' && companyWithEmployees) ? `<div class="invoice-percentage-controls" style="display:flex; flex-direction:column; gap:0.35rem;"><div style="display:flex; justify-content:space-between; font-size:0.8rem; color: var(--text-secondary);"><span>Udział procentowy</span><strong class="invoice-percentage-value">${percentageValue}%</strong></div><input type="range" min="0" max="100" step="1" class="invoice-percentage-slider" data-issuer-id="${issuer.id}" value="${percentageValue}" style="width:100%; accent-color: var(--accent-primary);" ${isInvoicesIssued ? 'disabled' : ''}><div style="font-size:0.78rem; color: var(--text-secondary);">Kwota z procentu: <strong class="invoice-percentage-amount" style="color: var(--warning);">${Calculations.formatInvoiceCurrency(clientInvoice.netRevenue * (percentageValue / 100))}</strong></div></div>` : ''}
                ${issuer.type === 'SEPARATE_COMPANY' ? `<label style="display:flex; align-items:center; gap:0.5rem; color: var(--text-secondary); font-size:0.8rem; margin:0;"><input type="checkbox" class="invoice-company-with-employees-toggle" data-issuer-id="${issuer.id}" ${companyWithEmployees ? 'checked' : ''} style="width:auto;" ${isInvoicesIssued ? 'disabled' : ''}>Tylko Faktura za firmę z rozliczenia</label>` : ''}
                ${issuer.type === 'SEPARATE_COMPANY' && companyWithEmployees && settlementInfo ? `<div style="font-size:0.78rem; color: var(--text-secondary);">${settlementInfo.invoiceAmountLabel}: <strong style="color: var(--warning);">${Calculations.formatInvoiceCurrency(settlementInfo.invoiceAmount || 0)}</strong></div>` : ''}
                <input type="number" step="0.01" class="invoice-manual-amount" data-issuer-id="${issuer.id}" value="${Number.isFinite(parseFloat(manualAmount)) ? parseFloat(manualAmount) : ''}" placeholder="Ręczna kwota" ${clientInvoice.mode === 'MANUAL' && !(issuer.type === 'SEPARATE_COMPANY' && companyWithEmployees) && !isInvoicesIssued ? '' : 'disabled'} style="display:${clientInvoice.mode === 'MANUAL' && !(issuer.type === 'SEPARATE_COMPANY' && companyWithEmployees) ? 'block' : 'none'};">
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

  if (!isInvoicesIssued) {
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
  }

  lucide.createIcons();
  applyArchivedReadOnlyMode(state.isArchived === true);
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
    // Show concise header in main panel (with trailing dot as requested)
    subtitle.textContent = `Wyrównanie po fakturach dla ${formatMonthLabel(selectedMonth)}.`;

    // Insert or update informational text under the main subtitle so the
    // inner panels are displayed directly in the main view and the short
    // description is visible right below the "Wyrównanie" heading.
    const infoId = 'invoice-equalization-info';
    let infoEl = document.getElementById(infoId);
    const infoText = 'Poniżej widać, kto dostał wpływy z faktur na konto i komu powinien przekazać środki, aby końcowy rozrachunek zgadzał się z rozliczeniem.';
    if (!infoEl) {
      infoEl = document.createElement('div');
      infoEl.id = infoId;
      infoEl.style.marginTop = '0.45rem';
      infoEl.style.marginBottom = '0.85rem';
      infoEl.style.color = 'var(--text-secondary)';
      if (subtitle.parentElement) {
        subtitle.parentElement.insertBefore(infoEl, subtitle.nextSibling);
      }
    }
    infoEl.textContent = infoText;
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

  normalizeCurrencySuffixSpacing(document.getElementById('invoice-equalization-view') || list.closest('.view') || document.getElementById('main-content'));
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
    const dayData = ensureMonthlySheetDayData(s, d);
    
    if (type === 'start') dayData.globalStart = newVal;
    if (type === 'end') dayData.globalEnd = newVal;
    
    // Recalculate
    if (dayData.isWholeTeamChecked && dayData.globalStart && dayData.globalEnd) {
      const calcH = Calculations.calculateHours(dayData.globalStart, dayData.globalEnd);
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

  setupListSortable(document.getElementById('works-sheets-body'), {
    onEnd: () => {
      Store.reorderWorksSheets(getSortableRowIds(document.getElementById('works-sheets-body')));
    }
  });
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

  sheetsForMonth.forEach((s, idx) => {
    const entries = s.entries || [];
    const totalItems = entries.length;
    const totalValue = entries.reduce((sum, e) => sum + (parseFloat(e.quantity) * parseFloat(e.price)), 0);

    const tr = document.createElement('tr');
    tr.dataset.id = s.id;
    tr.innerHTML = `
      <td>${s.client || '-'}</td>
      <td>${s.site || '-'}</td>
      <td>${totalItems} pozycji</td>
      <td style="color: var(--success); font-weight: bold;">${totalValue.toFixed(2)} zł</td>
      <td>
        <div style="display:flex; gap:0.5rem; flex-wrap:wrap;">
          <button class="btn btn-primary btn-sm btn-open-works-sheet" data-id="${s.id}" style="margin-right:0.5rem;">Otwórz</button>
          <button class="btn btn-secondary btn-icon btn-edit-works-sheet" data-id="${s.id}" style="margin-right:0.5rem;">
            <i data-lucide="edit-2" style="width:16px;height:16px"></i>
          </button>
          <button class="btn btn-danger btn-icon btn-delete-works-sheet" data-id="${s.id}">
            <i data-lucide="trash-2" style="width:16px;height:16px"></i>
          </button>
          <button class="btn btn-secondary btn-icon drag-handle btn-reorder-handle" type="button" title="Przytrzymaj, aby zmienić kolejność" aria-label="Przytrzymaj, aby zmienić kolejność">
            <i data-lucide="grip-horizontal" style="width:16px;height:16px"></i>
          </button>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
    appendMobileCardDividerRow(tbody, 5, idx === sheetsForMonth.length - 1);
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
  setupListSortable(document.getElementById('works-sheets-body'), {
    onEnd: () => {
      Store.reorderWorksSheets(getSortableRowIds(document.getElementById('works-sheets-body')));
    }
  });
  normalizeCurrencySuffixSpacing(document.getElementById('works-view'));
  applyArchivedReadOnlyMode(state.isArchived === true);
}

function renderWorksSheetDetail(sheetId) {
  const state = Store.getState();
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
  const worksEntryDateCounts = (sheet.entries || []).reduce((counts, entry) => {
    const key = entry?.date || '';
    counts[key] = (counts[key] || 0) + 1;
    return counts;
  }, {});
  
  const worksEntryOrder = new Map((sheet.entries || []).map((entry, index) => [entry.id, index]));
  let entries = [...(sheet.entries || [])];
  entries.sort((a, b) => {
    const dateDiff = new Date(a.date || '1970-01-01') - new Date(b.date || '1970-01-01');
    if (dateDiff !== 0) return dateDiff;
    return (worksEntryOrder.get(a.id) ?? 0) - (worksEntryOrder.get(b.id) ?? 0);
  });

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
      const canReorderEntry = (worksEntryDateCounts[e.date || ''] || 0) > 1;

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
      tr.className = 'works-entry-row';
      tr.dataset.id = e.id;
      tr.dataset.sortDate = e.date || '';
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
            ${canReorderEntry ? `<button class="btn btn-secondary btn-icon drag-handle btn-reorder-handle" type="button" title="Przytrzymaj, aby zmienić kolejność pozycji z tej samej daty" aria-label="Przytrzymaj, aby zmienić kolejność pozycji z tej samej daty">
              <i data-lucide="grip-horizontal" style="width:16px;height:16px"></i>
            </button>` : ''}
          </div>
        </td>
      `;
      tbody.appendChild(tr);

      if (sheet.activePersons && sheet.activePersons.length > 0) {
        const subTr = document.createElement('tr');
        subTr.className = 'works-hours-row';
        subTr.dataset.parentId = e.id;

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

  setupListSortable(document.getElementById('works-entries-body'), {
    draggable: 'tr.works-entry-row[data-id]',
    onMove: (evt) => allowSortOnlyWithinSameDate(evt),
    onEnd: (evt) => {
      moveWorksEntrySubRowAfterMainRow(evt.item);
      const date = evt.item?.dataset?.sortDate || '';
      if (!date) return;
      const selector = `tr.works-entry-row[data-id][data-sort-date="${date}"]`;
      Store.reorderWorksSheetEntriesForDate(sheetId, date, getSortableRowIds(document.getElementById('works-entries-body'), selector));
    }
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

        updateMonthPickerRestrictions(document.getElementById('works-entry-date'), getSelectedMonthKey());

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

  normalizeCurrencySuffixSpacing(document.getElementById('works-view'));
  lucide.createIcons();
  applyArchivedReadOnlyMode(state.isArchived === true);
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
  
  normalizeCurrencySuffixSpacing(document.getElementById('works-view'));
  lucide.createIcons();
  applyArchivedReadOnlyMode(state.isArchived === true);
}

// ==========================================
// SETTINGS & SCALING
// ==========================================

const DEFAULT_APPEARANCE_SETTINGS = { theme: 'system', scaleLarge: 100, scaleVertical: 100 };
const DEFAULT_MOBILE_APPEARANCE_SETTINGS = { theme: 'system', scaleLarge: 60, scaleVertical: 80 };

const AUTHORIZED_EMAILS = [
  'zete777@gmail.com',
  'piwkolinio@gmail.com',
  'dkrzysztof297@gmail.com',
  'zete@op.pl'
];

function initFirebaseAuth() {
  const auth = firebase.auth();
  
  // Ensure login survives browser restart
  auth.setPersistence(firebase.auth.Auth.Persistence.LOCAL)
    .catch(err => console.error("Persistence error:", err));

  const db = firebase.database();
  let sharedMetaCommonRef = null;
  let currentMonthMetaRef = null;
  let currentMonthMonthlySheetsRef = null;
  let currentMonthWorksSheetsRef = null;
  let currentMonthExpensesRef = null;
  let currentInvoiceTotalsYearRef = null;
  let currentMonthRef = null;
  let sharedDataCommonRef = null;
  let monthChangedHandler = null;
  let isUsingLegacySharedDataSync = false;
  const latestRemoteMonthMetaByKey = {};
  const overlay = document.getElementById('firebase-auth-overlay');
  const appRoot = document.getElementById('app-root');
  
  const emailInput = document.getElementById('firebase-email');
  const passwordInput = document.getElementById('firebase-password');
  const btnLoginEmail = document.getElementById('btn-firebase-email-login');
  const btnLoginGoogle = document.getElementById('btn-firebase-google');
  const errorMsg = document.getElementById('firebase-auth-error');
  const loginContent = document.getElementById('firebase-auth-login-content');
  const loadingMsg = document.getElementById('firebase-auth-loading');

  const setAuthLoadingState = (isLoading = false, message = 'Trwa Logowanie....') => {
    if (loginContent) {
      loginContent.style.display = isLoading ? 'none' : 'block';
    }
    if (loadingMsg) {
      loadingMsg.textContent = message;
      loadingMsg.style.display = isLoading ? 'block' : 'none';
    }
    if (errorMsg && isLoading) {
      errorMsg.style.display = 'none';
      errorMsg.textContent = '';
    }
    if (btnLoginEmail) btnLoginEmail.disabled = isLoading;
    if (btnLoginGoogle) btnLoginGoogle.disabled = isLoading;
    if (emailInput) emailInput.disabled = isLoading;
    if (passwordInput) passwordInput.disabled = isLoading;
  };

  const detachFirebaseSyncListeners = () => {
    if (sharedMetaCommonRef) sharedMetaCommonRef.off('value');
    if (currentMonthMetaRef) currentMonthMetaRef.off('value');
    if (currentMonthMonthlySheetsRef) currentMonthMonthlySheetsRef.off('value');
    if (currentMonthWorksSheetsRef) currentMonthWorksSheetsRef.off('value');
    if (currentMonthExpensesRef) currentMonthExpensesRef.off('value');
    if (currentInvoiceTotalsYearRef) currentInvoiceTotalsYearRef.off('value');
    if (currentMonthRef) currentMonthRef.off('value');
    if (sharedDataCommonRef) sharedDataCommonRef.off('value');
    sharedMetaCommonRef = null;
    currentMonthMetaRef = null;
    currentMonthMonthlySheetsRef = null;
    currentMonthWorksSheetsRef = null;
    currentMonthExpensesRef = null;
    currentInvoiceTotalsYearRef = null;
    currentMonthRef = null;
    sharedDataCommonRef = null;
    Object.keys(latestRemoteMonthMetaByKey).forEach(key => delete latestRemoteMonthMetaByKey[key]);
    if (monthChangedHandler) {
      window.removeEventListener('monthChanged', monthChangedHandler);
      monthChangedHandler = null;
    }
  };

  const syncCommonFromSharedData = () => {
    db.ref('shared_data/common').once('value').then(snapshot => {
      const rawCommon = getTrackedFirebaseSnapshotValue(snapshot);
      if (rawCommon && !window.isOfflineMode) {
        window.isImportingFromFirebase = true;
        Store.applyRemoteCommon(rawCommon, {});
        window.isImportingFromFirebase = false;
      }
    }).catch(error => {
      console.error('firebase: Błąd pobierania common z /shared_data:', error);
    });
  };

  const syncMonthFromSharedData = (monthKey) => {
    if (!monthKey) return;

    db.ref(`shared_data/months/${monthKey}`).once('value').then(snapshot => {
      const rawMonthData = getTrackedFirebaseSnapshotValue(snapshot);
      if (rawMonthData && !window.isOfflineMode) {
        migrateRemoteMonthCollectionsIfNeeded(monthKey, rawMonthData);
        window.isImportingFromFirebase = true;
        Store.applyRemoteMonth(monthKey, rawMonthData, { months: { [monthKey]: {} } });
        window.isImportingFromFirebase = false;

        renderAll();
      }
    }).catch(error => {
      console.error(`firebase: Błąd pobierania miesiąca ${monthKey} z /shared_data:`, error);
    });
  };

  const syncMonthScopeFromSharedData = (monthKey, scopeKey, remoteMeta = {}) => {
    if (!monthKey || !scopeKey || !Store.applyRemoteMonthScope) return Promise.resolve();

    const scopePath = scopeKey === 'monthSettings.commonSnapshot'
      ? `shared_data/months/${monthKey}/monthSettings/commonSnapshot`
      : `shared_data/months/${monthKey}/${scopeKey}`;

    return db.ref(scopePath).once('value').then(snapshot => {
      const rawScopeData = getTrackedFirebaseSnapshotValue(snapshot);
      if (scopeKey !== 'monthSettings.commonSnapshot' && scopeKey !== 'monthSettings' && rawScopeData && typeof rawScopeData === 'object') {
        migrateRemoteMonthCollectionsIfNeeded(monthKey, { [scopeKey]: rawScopeData });
      }

      if (window.isOfflineMode) return;

      window.isImportingFromFirebase = true;
      Store.applyRemoteMonthScope(monthKey, scopeKey, rawScopeData, remoteMeta);
      window.isImportingFromFirebase = false;
      renderAll();
    }).catch(error => {
      console.error(`firebase: Błąd pobierania zakresu ${scopeKey} miesiąca ${monthKey} z /shared_data:`, error);
    });
  };

  const extractYearFromMonthKey = (monthKey = '') => (/^\d{4}-\d{2}$/.test(monthKey) ? monthKey.slice(0, 4) : '');

  const getYearMonthKeys = (year = '') => (/^\d{4}$/.test((year || '').toString())
    ? Array.from({ length: 12 }, (_, index) => `${year}-${String(index + 1).padStart(2, '0')}`)
    : []);

  const buildYearScopedState = (state = {}, year = '') => ({
    version: 'v3',
    common: state?.common || {},
    months: getYearMonthKeys(year).reduce((acc, monthKey) => {
      if (state?.months?.[monthKey]) {
        acc[monthKey] = state.months[monthKey];
      }
      return acc;
    }, {})
  });

  const buildYearScopedMeta = (meta = {}, year = '') => ({
    common: meta?.common || {},
    months: getYearMonthKeys(year).reduce((acc, monthKey) => {
      if (meta?.months?.[monthKey]) {
        acc[monthKey] = meta.months[monthKey];
      }
      return acc;
    }, {})
  });

  const fetchYearMonthMetadata = (year = '') => {
    const monthKeys = getYearMonthKeys(year);
    return Promise.all(monthKeys.map(monthKey =>
      db.ref(`shared_meta/months/${monthKey}`).once('value').then(snapshot => ({
        monthKey,
        meta: getTrackedFirebaseSnapshotValue(snapshot) || {}
      }))
    )).then(entries => ({
      months: Object.fromEntries(entries.map(entry => [entry.monthKey, entry.meta]))
    }));
  };

  const fetchRemoteMonthRecord = (monthKey = '') => {
    if (!monthKey) return Promise.resolve(null);
    return db.ref(`shared_data/months/${monthKey}`).once('value').then(snapshot => {
      const rawMonthData = getTrackedFirebaseSnapshotValue(snapshot);
      if (rawMonthData) {
        migrateRemoteMonthCollectionsIfNeeded(monthKey, rawMonthData);
      }
      return rawMonthData || null;
    });
  };

  const syncYearFromRemoteIfNeeded = (year = '', options = {}) => {
    const { renderAfterSync = true } = options;
    const monthKeys = getYearMonthKeys(year);
    if (monthKeys.length === 0) return Promise.resolve({ months: {} });

    return fetchYearMonthMetadata(year).then(remoteMeta => {
      monthKeys.forEach(monthKey => {
        latestRemoteMonthMetaByKey[monthKey] = remoteMeta.months?.[monthKey] || {};
      });

      const localFullState = Store.getExportData ? Store.getExportData() : { version: 'v3', common: {}, months: {} };
      const monthsToFetch = monthKeys.filter(monthKey => {
        const remoteMonthMeta = remoteMeta.months?.[monthKey] || {};
        return !localFullState.months?.[monthKey]
          || !Object.keys(remoteMonthMeta).length
          || (Store.shouldFetchRemoteMonth && Store.shouldFetchRemoteMonth(monthKey, { months: { [monthKey]: remoteMonthMeta } }));
      });

      if (monthsToFetch.length === 0 || window.isOfflineMode) {
        return remoteMeta;
      }

      return Promise.all(monthsToFetch.map(monthKey =>
        fetchRemoteMonthRecord(monthKey).then(rawMonthData => ({ monthKey, rawMonthData }))
      )).then(results => {
        results.forEach(({ monthKey, rawMonthData }) => {
          if (!rawMonthData) return;
          window.isImportingFromFirebase = true;
          try {
            Store.applyRemoteMonth(monthKey, rawMonthData, { months: { [monthKey]: remoteMeta.months?.[monthKey] || {} } });
          } finally {
            window.isImportingFromFirebase = false;
          }
        });

        if (renderAfterSync) {
          renderAll();
        }

        return remoteMeta;
      });
    });
  };

  const setupLegacySharedDataListeners = () => {
    detachFirebaseSyncListeners();
    isUsingLegacySharedDataSync = true;
    Store.setSharedHistoryCaches({ history: [], dailyBackups: [], monthlyBackups: [] });

    sharedDataCommonRef = db.ref('shared_data/common');
    sharedDataCommonRef.on('value', (snapshot) => {
      const rawCommon = getTrackedFirebaseSnapshotValue(snapshot);
      if (rawCommon && !window.isOfflineMode) {
        window.isImportingFromFirebase = true;
        Store.applyRemoteCommon(rawCommon, {});
        window.isImportingFromFirebase = false;
      }
    });

    const attachLegacyMonthListener = (monthKey) => {
      if (currentMonthRef) currentMonthRef.off('value');
      currentMonthRef = db.ref(`shared_data/months/${monthKey}`);
      currentMonthRef.on('value', (snapshot) => {
        const rawMonthData = getTrackedFirebaseSnapshotValue(snapshot);
        if (rawMonthData && !window.isOfflineMode) {
          window.isImportingFromFirebase = true;
          Store.applyRemoteMonth(monthKey, rawMonthData, { months: { [monthKey]: {} } });
          window.isImportingFromFirebase = false;
        }
      });
    };

    attachLegacyMonthListener(Store.getSelectedMonth());
    monthChangedHandler = (e) => {
      if (!window.isOfflineMode) {
        attachLegacyMonthListener(e.detail.month);
        syncMonthFromSharedData(e.detail.month);
      }
    };
    window.addEventListener('monthChanged', monthChangedHandler);
  };

  const fallbackToLegacySharedDataSync = (error, path = 'shared_meta') => {
    if (!appIsFirebasePermissionDeniedError(error)) return false;
    console.warn(`firebase: Brak uprawnień do ${path}. Przełączam synchronizację na zgodność tylko z /shared_data.`, error);
    if (!isUsingLegacySharedDataSync) {
      setupLegacySharedDataListeners();
      syncCommonFromSharedData();
      syncMonthFromSharedData(Store.getSelectedMonth());
    }
    return true;
  };

  const refreshMonthFromRemoteIfNeeded = (monthKey, remoteMeta) => {
    if (!monthKey || !Store.getRemoteMonthSyncScopesToFetch) return;

    const scopesToFetch = Store.getRemoteMonthSyncScopesToFetch(monthKey, remoteMeta);
    if (!Array.isArray(scopesToFetch) || scopesToFetch.length === 0) return;

    scopesToFetch.forEach(scopeKey => {
      syncMonthScopeFromSharedData(monthKey, scopeKey, remoteMeta);
    });
  };
  
  auth.onAuthStateChanged((user) => {
    if (user) {
      const userEmail = (user.email || '').toLowerCase().trim();
      const isAuthorized = AUTHORIZED_EMAILS.includes(userEmail);
      
      console.log("Użytkownik zalogowany:", userEmail, "Autoryzacja:", isAuthorized);

      if (!isAuthorized) {
        if (overlay) {
          overlay.style.display = 'flex';
          overlay.innerHTML = `
            <div style="text-align: center; background: rgba(0,0,0,0.85); padding: 2rem; border-radius: 12px; border: 1px solid var(--danger); max-width: 400px;">
              <h2 style="color: var(--danger); margin-bottom: 1rem;">Brak uprawnień</h2>
              <p style="color: white; margin-bottom: 1.5rem;">Twoje konto (<strong>${userEmail}</strong>) nie ma uprawnień do przeglądania tej bazy.</p>
              <button onclick="firebase.auth().signOut()" class="btn btn-secondary" style="width: 100%">Wyloguj się</button>
            </div>
          `;
        }
        if (appRoot) appRoot.style.display = 'none';
        detachFirebaseSyncListeners();
        return;
      }

      if (overlay) overlay.style.display = 'none';
      if (appRoot) appRoot.style.display = 'flex';
      setAuthLoadingState(false);
      refreshForceResyncAdminStatus();
      
      const statusText = document.getElementById('sidebar-auth-status-text');
      const btnSidebarLogout = document.getElementById('btn-sidebar-firebase-logout');
      if (statusText) {
         statusText.textContent = user.email || 'Zalogowano';
         statusText.style.color = 'var(--success)';
      }
      if (btnSidebarLogout) {
         btnSidebarLogout.textContent = 'Wyloguj';
         btnSidebarLogout.style.display = 'block';
         btnSidebarLogout.onclick = () => {
           Promise.resolve(Store.flushPendingRemoteSync?.()).finally(() => auth.signOut());
         };
      }
      
      // Sync logic V3 (common + selected month + yearly invoice aggregates)
      const setupV3Listeners = () => {
        detachFirebaseSyncListeners();
        isUsingLegacySharedDataSync = false;
        let activeYear = extractYearFromMonthKey(Store.getSelectedMonth());

        const syncCommonIfNeeded = (remoteCommonMeta = {}) => {
          const wrappedRemoteMeta = { common: remoteCommonMeta };
          if (!Store.shouldFetchRemoteCommon || !Store.shouldFetchRemoteCommon(wrappedRemoteMeta)) return;

          db.ref('shared_data/common').once('value').then(snapshot => {
            const rawCommon = getTrackedFirebaseSnapshotValue(snapshot);
            if (rawCommon && !window.isOfflineMode) {
              window.isImportingFromFirebase = true;
              Store.applyRemoteCommon(rawCommon, wrappedRemoteMeta);
              window.isImportingFromFirebase = false;
            }
          }).catch(error => {
            console.error('firebase: Błąd pobierania common z /shared_data:', error);
          });
        };

        const attachMonthMetaListener = (monthKey) => {
          if (currentMonthMetaRef) currentMonthMetaRef.off('value');
          currentMonthMetaRef = db.ref(`shared_meta/months/${monthKey}`);
          currentMonthMetaRef.on('value', (snapshot) => {
            const normalizedMonthMeta = getTrackedFirebaseSnapshotValue(snapshot) || {};
            latestRemoteMonthMetaByKey[monthKey] = normalizedMonthMeta;
            const remoteMonthMeta = { months: { [monthKey]: normalizedMonthMeta } };
            refreshMonthFromRemoteIfNeeded(monthKey, remoteMonthMeta);
          }, (error) => {
            fallbackToLegacySharedDataSync(error, `shared_meta/months/${monthKey}`);
          });
        };

        const attachInvoiceTotalsYearListener = (yearKey = '') => {
          if (currentInvoiceTotalsYearRef) currentInvoiceTotalsYearRef.off('value');
          if (!/^\d{4}$/.test(yearKey)) {
            currentInvoiceTotalsYearRef = null;
            return;
          }

          currentInvoiceTotalsYearRef = db.ref(`shared_meta/invoiceTotals/years/${yearKey}`);
          currentInvoiceTotalsYearRef.on('value', (snapshot) => {
            Store.setInvoiceYearTotals?.(yearKey, getTrackedFirebaseSnapshotValue(snapshot) || {});
          }, (error) => {
            console.error(`firebase: Błąd nasłuchu agregatów faktur roku ${yearKey}:`, error);
          });
        };

        sharedMetaCommonRef = db.ref('shared_meta/common');
        sharedMetaCommonRef.on('value', (snapshot) => {
          const remoteCommonMeta = getTrackedFirebaseSnapshotValue(snapshot) || {};

          if (Object.keys(remoteCommonMeta).length === 0) {
            db.ref('shared_data/common').once('value').then(commonSnapshot => {
              const remoteCommon = getTrackedFirebaseSnapshotValue(commonSnapshot);
              if (remoteCommon) {
                window.isImportingFromFirebase = true;
                Store.applyRemoteCommon(remoteCommon, {});
                window.isImportingFromFirebase = false;
              }
            });
            return;
          }

          syncCommonIfNeeded(remoteCommonMeta);
        }, (error) => {
          fallbackToLegacySharedDataSync(error, 'shared_meta/common');
        });

        attachMonthMetaListener(Store.getSelectedMonth());
        attachInvoiceTotalsYearListener((Store.getSelectedMonth() || '').slice(0, 4));
        loadSharedBackupCachesForLogin();

        monthChangedHandler = (e) => {
          if (!window.isOfflineMode) {
            const targetMonth = e.detail.month;
            const targetYear = extractYearFromMonthKey(targetMonth);
            const didYearChange = !!targetYear && targetYear !== activeYear;

            attachInvoiceTotalsYearListener(targetYear);

            const attachSelectedMonthMeta = () => attachMonthMetaListener(targetMonth);
            if (didYearChange) {
              activeYear = targetYear;
              syncYearFromRemoteIfNeeded(targetYear).then(attachSelectedMonthMeta).catch(error => {
                console.error(`firebase: Błąd synchronizacji pełnego roku ${targetYear} po zmianie miesiąca:`, error);
                attachSelectedMonthMeta();
              });
              return;
            }

            attachSelectedMonthMeta();
          }
        };
        window.addEventListener('monthChanged', monthChangedHandler);
      };

      const applyRemoteBootstrapState = (remoteState = {}, remoteMeta = {}, options = {}) => {
        const { applyCommon = true, monthKeysToApply = [] } = options;

        if (!window.isOfflineMode && applyCommon && remoteState.common) {
          window.isImportingFromFirebase = true;
          Store.applyRemoteCommon(remoteState.common, { common: remoteMeta.common || {} });
          window.isImportingFromFirebase = false;
        }

        if (window.isOfflineMode) return;

        monthKeysToApply.forEach(monthKey => {
          if (!remoteState.months?.[monthKey]) return;
          window.isImportingFromFirebase = true;
          try {
            Store.applyRemoteMonth(monthKey, remoteState.months[monthKey], { months: { [monthKey]: remoteMeta.months?.[monthKey] || {} } });
          } finally {
            window.isImportingFromFirebase = false;
          }
        });
      };

      const continueWithStandardV3Bootstrap = () => {
        const selectedMonthKey = Store.getSelectedMonth();
        const selectedYear = extractYearFromMonthKey(selectedMonthKey);
        const yearMonthKeys = getYearMonthKeys(selectedYear);

        return Promise.all([
          db.ref('shared_meta/common').once('value'),
          fetchYearMonthMetadata(selectedYear)
        ]).then(([commonMetaSnapshot, yearRemoteMeta]) => {
          let remoteMeta = {
            common: getTrackedFirebaseSnapshotValue(commonMetaSnapshot) || {},
            months: yearRemoteMeta.months || {}
          };

          const localFullState = Store.getExportData ? Store.getExportData() : { version: 'v3', common: {}, months: {} };
          const localFullMeta = Store.getSyncMetadata ? Store.getSyncMetadata() : { common: {}, months: {} };
          const localState = buildYearScopedState(localFullState, selectedYear);
          const localMeta = buildYearScopedMeta(localFullMeta, selectedYear);
          const isPristineLocalState = isEffectivelyEmptyLocalSyncState(localFullState);
          const hasAnyRemoteYearMeta = yearMonthKeys.some(monthKey => Object.keys(remoteMeta.months?.[monthKey] || {}).length > 0);

          const shouldBootstrapRemoteMeta = Object.keys(remoteMeta.common || {}).length === 0
            && !hasAnyRemoteYearMeta;

          const bootstrapPromise = shouldBootstrapRemoteMeta
            ? rebuildEntireRemoteSharedDataStructure().then(result => ({
                meta: buildYearScopedMeta(result?.meta || remoteMeta, selectedYear),
                state: buildYearScopedState(result?.state || { version: 'v3', common: {}, months: {} }, selectedYear),
                preloadedMonthKeys: yearMonthKeys.filter(monthKey => !!result?.state?.months?.[monthKey]),
                preloadedCommon: !!result?.state?.common
              }))
            : Promise.resolve({
                meta: remoteMeta,
                state: null,
                preloadedMonthKeys: [],
                preloadedCommon: false
              });

          return bootstrapPromise.then(({ meta, state: bootstrappedState, preloadedMonthKeys, preloadedCommon }) => {
            remoteMeta = {
              common: meta?.common || remoteMeta.common || {},
              months: meta?.months || remoteMeta.months || {}
            };

            yearMonthKeys.forEach(monthKey => {
              latestRemoteMonthMetaByKey[monthKey] = remoteMeta.months?.[monthKey] || {};
            });

            const shouldFetchCommon = preloadedCommon
              ? false
              : (isPristineLocalState
                  || !Object.keys(remoteMeta.common || {}).length
                  || (Store.shouldFetchRemoteCommon && Store.shouldFetchRemoteCommon({ common: remoteMeta.common || {} })));

            const monthsToFetch = yearMonthKeys.filter(monthKey => {
              if (preloadedMonthKeys.includes(monthKey)) return false;
              const remoteMonthMeta = remoteMeta.months?.[monthKey] || {};
              return isPristineLocalState
                || !localFullState.months?.[monthKey]
                || !Object.keys(remoteMonthMeta).length
                || (Store.shouldFetchRemoteMonth && Store.shouldFetchRemoteMonth(monthKey, { months: { [monthKey]: remoteMonthMeta } }));
            });

            const monthsToApply = [...new Set([...preloadedMonthKeys, ...monthsToFetch])];

            return Promise.all([
              shouldFetchCommon ? db.ref('shared_data/common').once('value') : Promise.resolve(null),
              Promise.all(monthsToFetch.map(monthKey =>
                fetchRemoteMonthRecord(monthKey).then(rawMonthData => ({ monthKey, rawMonthData }))
              ))
            ]).then(([commonSnapshot, fetchedMonths]) => {
              const remoteState = bootstrappedState || {
                version: 'v3',
                common: localState.common || {},
                months: { ...(localState.months || {}) }
              };

              if (shouldFetchCommon && commonSnapshot) {
                remoteState.common = getTrackedFirebaseSnapshotValue(commonSnapshot) || {};
              }

              fetchedMonths.forEach(({ monthKey, rawMonthData }) => {
                if (!rawMonthData) return;
                remoteState.months[monthKey] = rawMonthData;
              });

              if (Store.ensureSyncMetadataBootstrap) {
                Store.ensureSyncMetadataBootstrap();
              }

              const effectiveRemoteState = remoteState;
              const conflictInfo = buildSyncConflictInfoFromMetadata(localMeta, remoteMeta);
              const hasStateDifference = JSON.stringify(localState) !== JSON.stringify(effectiveRemoteState);
              const shouldAskUser = !isPristineLocalState && conflictInfo.hasLocalNewerChanges === true;

              if (!shouldAskUser) {
                if (isPristineLocalState || conflictInfo.hasRemoteNewerChanges || hasStateDifference) {
                  applyRemoteBootstrapState(effectiveRemoteState, remoteMeta, {
                    applyCommon: shouldFetchCommon || preloadedCommon,
                    monthKeysToApply: monthsToApply
                  });
                }

                setupV3Listeners();
                return;
              }

              return askUserHowToResolveSyncConflict(conflictInfo, localState, effectiveRemoteState).then(choice => {
                if (choice === 'local') {
                  const syncOverwriteHistoryResult = Store.recordSharedSyncOverwriteHistory
                    ? Store.recordSharedSyncOverwriteHistory(
                        effectiveRemoteState,
                        localState,
                        remoteMeta,
                        localMeta,
                        'Nadpisano Bazę Główną bazą lokalną po konflikcie synchronizacji'
                      )
                    : { success: false, history: [] };
                  const updates = buildFirebaseStateUpdates(localState, localMeta);
                  if (syncOverwriteHistoryResult?.success && syncOverwriteHistoryResult.entry) {
                      Object.assign(updates, Store.buildSharedHistoryMigrationUpdates?.([
                        syncOverwriteHistoryResult.entry
                      ], syncOverwriteHistoryResult.removedEntries || []) || {});
                  }

                  return appUpdateFirebaseRootWithFallback(updates).then(() => {
                    console.log('firebase: Zachowano nowszą bazę lokalną i wysłano ją do Firebase.');
                    setupV3Listeners();
                  });
                }

                console.log('firebase: Użytkownik wybrał pobranie Bazy Głównej. Wykonuję pełny reset lokalnej bazy i pobieram dane od nowa.');
                return resetLocalDatabaseAndFetchMainDatabase()
                  .then(() => {
                    setupV3Listeners();
                  })
                  .catch(error => {
                    console.error('firebase: Nie udało się wykonać resetu i pobrania Bazy Głównej po wyborze użytkownika. Stosuję standardowe pobranie zakresów.', error);
                    applyRemoteBootstrapState(effectiveRemoteState, remoteMeta, {
                      applyCommon: shouldFetchCommon || preloadedCommon,
                      monthKeysToApply: monthsToApply
                    });
                    setupV3Listeners();
                  });
              });
            });
          });
        }).catch(err => {
          console.error('firebase: Błąd przygotowania porównania sync:', err);
          setupV3Listeners();
        });
      };

      const continueWithForceResyncCheck = (currentUserEmail = '') => {
        return db.ref(FORCE_RESYNC_CONTROL_PATH).once('value').then(snapshot => {
          const forceResyncControl = normalizeForceResyncControl(getTrackedFirebaseSnapshotValue(snapshot) || {});

          if (!shouldForceResyncForEmail(forceResyncControl, currentUserEmail)) {
            return continueWithStandardV3Bootstrap();
          }

          console.warn('firebase: Wykryto aktywny force_resync. Pomijam porównanie shared_meta i wykonuję pełny reset lokalnej bazy.');
          return db.ref('shared_data').once('value').then(sharedDataSnapshot => {
            const sharedData = getTrackedFirebaseSnapshotValue(sharedDataSnapshot) || {};
            return performForceResyncBootstrap(sharedData, currentUserEmail, forceResyncControl, setupV3Listeners);
          }).catch(error => {
            window.isImportingFromFirebase = false;
            console.error('firebase: Błąd podczas force_resync:', error);
            return continueWithStandardV3Bootstrap();
          });
        }).catch(error => {
          console.error('firebase: Błąd sprawdzania force_resync:', error);
          return continueWithStandardV3Bootstrap();
        });
      };

      continueWithForceResyncCheck(userEmail);
      
    } else {
      setAuthLoadingState(false);
      refreshForceResyncAdminStatus();
      if (overlay) {
        overlay.style.display = 'flex';
        if (overlay.innerHTML.includes('Brak uprawnień')) {
           location.reload();
        }
      }
      if (appRoot) appRoot.style.display = 'none';
      detachFirebaseSyncListeners();
    }
  });

  const loginForm = document.getElementById('firebase-login-form');
  if (loginForm) {
    loginForm.onsubmit = (e) => {
      e.preventDefault();
      const email = emailInput?.value?.trim() || '';
      const password = passwordInput?.value || '';
      
      console.log("Attempting login for:", email);
      
      if (!email || !password) {
        const m = 'Wprowadź e-mail i hasło.';
        alert(m);
        if (errorMsg) { errorMsg.textContent = m; errorMsg.style.display = 'block'; }
        return false;
      }

      setAuthLoadingState(true, 'Trwa Logowanie....');
      
      auth.signInWithEmailAndPassword(email, password)
        .then(() => {
          console.log("Login successful!");
        })
        .catch(err => {
          setAuthLoadingState(false);
          console.error("Firebase Login Error:", err);
          let userFriendlyMsg = 'Błąd logowania: ' + err.message;
          if (err.code === 'auth/user-not-found') userFriendlyMsg = 'Nie znaleziono użytkownika o tym adresie e-mail.';
          if (err.code === 'auth/wrong-password') userFriendlyMsg = 'Błędne hasło.';
          
          alert(userFriendlyMsg);
          if (errorMsg) {
            errorMsg.textContent = userFriendlyMsg;
            errorMsg.style.display = 'block';
          }
        });
      return false;
    };
  }

  if (btnLoginGoogle) {
    btnLoginGoogle.addEventListener('click', (e) => {
      e.preventDefault();
      setAuthLoadingState(true, 'Trwa Logowanie....');
      const provider = new firebase.auth.GoogleAuthProvider();
      auth.signInWithPopup(provider).catch(err => {
        setAuthLoadingState(false);
        console.error("Firebase Google Error:", err);
        alert('Błąd Google: ' + err.message);
        if (errorMsg) {
          errorMsg.textContent = 'Błąd Google: ' + err.message;
          errorMsg.style.display = 'block';
        }
      });
    });
  }

  const btnOffline = document.getElementById('btn-offline-mode');
  if (btnOffline) {
    btnOffline.addEventListener('click', () => {
      if (confirm('Czy na pewno chcesz uruchomić aplikację w trybie Offline? Zmiany zostaną zapisane tylko lokalnie na tym urządzeniu.')) {
        startOfflineMode();
      }
    });
  }

  // Check if we were already in offline mode in this session
  if (sessionStorage.getItem('workTrackerOfflineMode') === 'true') {
    startOfflineMode();
  }
}

function startOfflineMode() {
  console.log("URUCHAMIANIE TRYBU OFFLINE (TYLKO DANE LOKALNE)...");
  window.isOfflineMode = true;
  sessionStorage.setItem('workTrackerOfflineMode', 'true');
  setSettingsRemoteDatabaseResetButtonVisibility(false);
  
  const overlay = document.getElementById('firebase-auth-overlay');
  const appRoot = document.getElementById('app-root');
  const statusText = document.getElementById('sidebar-auth-status-text');
  const logoutBtn = document.getElementById('btn-sidebar-firebase-logout');

  // Finalizacja wejścia - używamy tego co już jest w Store (z localStorage)
  try {
    renderAll();
  } catch (error) {
    console.error('BŁĄD STARTU OFFLINE:', error);
    throw error;
  }
  if (overlay) overlay.style.display = 'none';
  if (appRoot) appRoot.style.display = 'flex';
  
  if (statusText) {
    statusText.textContent = 'TRYB OFFLINE';
    statusText.style.color = 'var(--danger)';
    statusText.style.fontWeight = '800';
  }
  if (logoutBtn) {
    logoutBtn.textContent = 'Wyjdź z Offline';
    logoutBtn.onclick = () => {
      sessionStorage.removeItem('workTrackerOfflineMode');
      location.reload();
    };
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

function formatOrphanedPersonCleanupResultMessage(result = {}) {
  const monthLabel = (result?.month || Store.getSelectedMonth() || '').toString();
  if (!result?.changed) {
    return `Nie znaleziono osieroconych wpisów osób w miesiącu ${monthLabel}.`;
  }

  const parts = [];
  if (result.removedMonthlySheetReferences > 0) parts.push(`arkusze godzin: ${result.removedMonthlySheetReferences}`);
  if (result.removedWorksSheetReferences > 0) parts.push(`arkusze prac: ${result.removedWorksSheetReferences}`);
  if (result.removedExpenseEntries > 0) parts.push(`koszty i zaliczki: ${result.removedExpenseEntries}`);
  if (result.removedMonthSettingReferences > 0) parts.push(`ustawienia miesiąca: ${result.removedMonthSettingReferences}`);
  if (result.removedInvoiceReferences > 0) parts.push(`faktury: ${result.removedInvoiceReferences}`);

  const details = parts.length > 0 ? ` (${parts.join(', ')})` : '';
  return `Usunięto ${result.totalRemoved || 0} osieroconych wpisów osób z miesiąca ${monthLabel}.${details}`;
}

function setSettingsRemoteDatabaseResetButtonVisibility(isVisible = !window.isOfflineMode) {
  const button = document.getElementById('btn-reset-download-main-database');
  if (!button) return;
  button.style.display = isVisible ? '' : 'none';
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

const SYNC_CONFLICT_SCOPE_LABELS = {
  'common': 'Dane wspólne',
  'month.monthlySheets': 'Arkusze godzin',
  'month.worksSheets': 'Arkusze wykonanych prac',
  'month.expenses': 'Koszty i zaliczki',
  'month.monthSettings': 'Ustawienia miesiąca',
  'month.monthSettings.commonSnapshot': 'Snapshot miesiąca'
};

function getSyncConflictScopeDisplayLabel(scope = '') {
  return SYNC_CONFLICT_SCOPE_LABELS[scope] || scope || 'Zakres danych';
}

function parseComparableTimestamp(value = '') {
  const timestamp = Date.parse((value || '').toString().trim());
  return Number.isFinite(timestamp) ? timestamp : 0;
}

function compareSyncMetaEntriesByRevisionAndDate(localEntry = {}, remoteEntry = {}) {
  const localRevision = parseInt(localEntry?.revision, 10) || 0;
  const remoteRevision = parseInt(remoteEntry?.revision, 10) || 0;

  if (localRevision > remoteRevision) return 'local';
  if (remoteRevision > localRevision) return 'remote';

  const localTimestamp = parseComparableTimestamp(localEntry?.updatedAt);
  const remoteTimestamp = parseComparableTimestamp(remoteEntry?.updatedAt);

  if (localTimestamp > remoteTimestamp) return 'local';
  if (remoteTimestamp > localTimestamp) return 'remote';

  return 'equal';
}

function buildSyncConflictInfoFromMetadata(localMeta = {}, remoteMeta = {}) {
  if (typeof Store === 'undefined' || !Store.normalizeSyncMetadata) {
    return { hasLocalNewerChanges: false, hasRemoteNewerChanges: false, localNewer: [], remoteNewer: [] };
  }

  const localNormalized = Store.normalizeSyncMetadata(localMeta || {});
  const remoteNormalized = Store.normalizeSyncMetadata(remoteMeta || {});
  const localNewer = [];
  const remoteNewer = [];

  const commonComparison = compareSyncMetaEntriesByRevisionAndDate(localNormalized.common || {}, remoteNormalized.common || {});
  if (commonComparison !== 'equal') {
    const payload = {
      scope: 'common',
      scopeLabel: getSyncConflictScopeDisplayLabel('common'),
      month: '',
      localEntry: localNormalized.common || {},
      remoteEntry: remoteNormalized.common || {}
    };
    if (commonComparison === 'local') localNewer.push(payload);
    if (commonComparison === 'remote') remoteNewer.push(payload);
  }

  const monthKeys = [...new Set([
    ...Object.keys(localNormalized.months || {}),
    ...Object.keys(remoteNormalized.months || {})
  ])].sort((a, b) => b.localeCompare(a, 'pl-PL'));

  monthKeys.forEach(monthKey => {
    ['monthlySheets', 'worksSheets', 'expenses', 'monthSettings', 'monthSettings.commonSnapshot'].forEach(scopeKey => {
      const localEntry = scopeKey === 'monthSettings.commonSnapshot'
        ? (localNormalized.months?.[monthKey]?.monthSettings?.commonSnapshot || {})
        : (scopeKey === 'monthSettings'
            ? (localNormalized.months?.[monthKey]?.monthSettings || {})
            : (localNormalized.months?.[monthKey]?.[scopeKey] || {}));
      const remoteEntry = scopeKey === 'monthSettings.commonSnapshot'
        ? (remoteNormalized.months?.[monthKey]?.monthSettings?.commonSnapshot || {})
        : (scopeKey === 'monthSettings'
            ? (remoteNormalized.months?.[monthKey]?.monthSettings || {})
            : (remoteNormalized.months?.[monthKey]?.[scopeKey] || {}));
      const comparison = compareSyncMetaEntriesByRevisionAndDate(localEntry, remoteEntry);
      if (comparison === 'equal') return;

      const payload = {
        scope: `month.${scopeKey}`,
        scopeLabel: getSyncConflictScopeDisplayLabel(`month.${scopeKey}`),
        month: monthKey,
        localEntry,
        remoteEntry
      };

      if (comparison === 'local') localNewer.push(payload);
      if (comparison === 'remote') remoteNewer.push(payload);
    });
  });

  return {
    hasLocalNewerChanges: localNewer.length > 0,
    hasRemoteNewerChanges: remoteNewer.length > 0,
    localNewer,
    remoteNewer
  };
}

function getLatestSyncMetadataTimestamp(meta = {}) {
  if (typeof Store === 'undefined' || !Store.normalizeSyncMetadata) return '';

  const normalized = Store.normalizeSyncMetadata(meta || {});
  let latestValue = '';
  let latestTimestamp = 0;

  const registerEntry = (entry = {}) => {
    const updatedAt = (entry?.updatedAt || '').toString().trim();
    const timestamp = parseComparableTimestamp(updatedAt);
    if (timestamp <= latestTimestamp) return;
    latestTimestamp = timestamp;
    latestValue = updatedAt;
  };

  registerEntry(normalized.common || {});
  Object.values(normalized.months || {}).forEach(monthMeta => {
    registerEntry(monthMeta?.monthlySheets || {});
    registerEntry(monthMeta?.worksSheets || {});
    registerEntry(monthMeta?.expenses || {});
    registerEntry(monthMeta?.monthSettings || {});
    registerEntry(monthMeta?.monthSettings?.commonSnapshot || {});
  });

  return latestValue;
}

function normalizeImportedDatabasePayload(parsed = {}, fileName = '') {
  return {
    fileName,
    app: (parsed?.app || '').toString().trim(),
    version: Math.max(1, parseInt(parsed?.version, 10) || 1),
    exportedAt: (parsed?.exportedAt || '').toString().trim(),
    stateData: parsed?.data || parsed,
    syncMetadata: parsed?.syncMetadata || parsed?.metadata?.syncMetadata || parsed?.meta?.syncMetadata || {}
  };
}

function buildImportedSyncMetadataForComparison(importPayload = {}, importedState = {}) {
  const explicitMeta = importPayload?.syncMetadata || {};
  if (Object.keys(explicitMeta).length > 0 && typeof Store !== 'undefined' && Store.normalizeSyncMetadata) {
    return Store.normalizeSyncMetadata(explicitMeta);
  }

  if (importPayload?.version !== 1 || !importPayload?.exportedAt || typeof Store === 'undefined' || !Store.buildSyncMetadataFromState || !Store.normalizeSyncMetadata) {
    return {};
  }

  const fallbackMeta = Store.normalizeSyncMetadata(Store.buildSyncMetadataFromState(importedState || {}) || {});
  if (fallbackMeta.common) {
    fallbackMeta.common.updatedAt = importPayload.exportedAt;
  }
  Object.values(fallbackMeta.months || {}).forEach(monthMeta => {
    if (monthMeta?.monthlySheets) monthMeta.monthlySheets.updatedAt = importPayload.exportedAt;
    if (monthMeta?.worksSheets) monthMeta.worksSheets.updatedAt = importPayload.exportedAt;
    if (monthMeta?.expenses) monthMeta.expenses.updatedAt = importPayload.exportedAt;
    if (monthMeta?.monthSettings) monthMeta.monthSettings.updatedAt = importPayload.exportedAt;
    if (monthMeta?.monthSettings?.commonSnapshot) monthMeta.monthSettings.commonSnapshot.updatedAt = importPayload.exportedAt;
  });

  return fallbackMeta;
}

function buildImportSyncWarningContext(importPayload = {}, importedState = {}, remoteBundle = null) {
  const importedMeta = buildImportedSyncMetadataForComparison(importPayload, importedState);
  const remoteMeta = remoteBundle?.meta || {};
  const conflictInfo = buildSyncConflictInfoFromMetadata(importedMeta, remoteMeta);
  const importedLatestTimestamp = getLatestSyncMetadataTimestamp(importedMeta) || (importPayload.version === 1 ? importPayload.exportedAt : '');
  const remoteLatestTimestamp = getLatestSyncMetadataTimestamp(remoteMeta);
  const importedLatestValue = parseComparableTimestamp(importedLatestTimestamp);
  const remoteLatestValue = parseComparableTimestamp(remoteLatestTimestamp);

  return {
    importedMeta,
    conflictInfo,
    importedLatestTimestamp,
    remoteLatestTimestamp,
    hasImportedNewerChanges: conflictInfo.hasLocalNewerChanges || (importedLatestValue > 0 && importedLatestValue > remoteLatestValue),
    hasRemoteNewerChanges: conflictInfo.hasRemoteNewerChanges || (remoteLatestValue > 0 && remoteLatestValue > importedLatestValue),
    shouldWarn: conflictInfo.hasLocalNewerChanges
      || conflictInfo.hasRemoteNewerChanges
      || (importedLatestValue > 0 && remoteLatestValue > 0 && importedLatestValue !== remoteLatestValue)
  };
}

function askUserHowToResolveImportedDatabaseConflict(conflictInfo = {}, importedState = {}, remoteState = {}, options = {}) {
  const {
    fileName = 'plik JSON',
    importedLatestTimestamp = '',
    remoteLatestTimestamp = '',
    hasImportedNewerChanges = false,
    hasRemoteNewerChanges = false
  } = options;

  const importedSummary = summarizeSyncState(importedState);
  const remoteSummary = summarizeSyncState(remoteState);
  const importedLatestLabel = importedLatestTimestamp ? formatHistoryTimestamp(importedLatestTimestamp) : 'brak';
  const remoteLatestLabel = remoteLatestTimestamp ? formatHistoryTimestamp(remoteLatestTimestamp) : 'brak';
  const importActionLabel = hasRemoteNewerChanges && !hasImportedNewerChanges
    ? 'Importuj bazę z pliku (ekstremalnie niezalecane)'
    : 'Importuj bazę z pliku (niezalecane)';

  let title = 'Wykryto konflikt importu bazy danych';
  let bodyLines = [
    'Importowana baza danych różni się od Bazy Danych Głównej na serwerze.',
    'Zalecane jest pozostawienie aktualnej Bazy Głównej i anulowanie importu.'
  ];

  if (hasRemoteNewerChanges && !hasImportedNewerChanges) {
    title = 'Importowana baza jest starsza od Bazy Głównej';
    bodyLines = [
      'Importowana baza danych jest starsza od Bazy Danych Głównej na serwerze.',
      'Ekstremalnie niezalecane jest importowanie tej bazy danych.',
      'Zalecane jest anulowanie importu i pozostanie przy aktualnej Bazie Głównej.'
    ];
  } else if (hasImportedNewerChanges && !hasRemoteNewerChanges) {
    title = 'Wykryto nowszą bazę w pliku importu';
    bodyLines = [
      'Importowana baza danych w pliku jest nowsza od Bazy Danych Głównej na serwerze.',
      'Zalecane jest anulowanie importu i pozostanie przy Bazie Głównej z serwera.',
      'Import z pliku powinien być używany tylko w sytuacjach awaryjnych.'
    ];
  }

  const detailedComparisonText = [
    `Plik importu: ${fileName}`,
    `Importowana baza: ${importedSummary}`,
    `Baza Główna na serwerze: ${remoteSummary}`,
    '',
    `Najnowsza data meta/importu pliku: ${importedLatestLabel}`,
    `Najnowsza data meta serwera: ${remoteLatestLabel}`,
    '',
    formatSyncConflictEntries(conflictInfo.localNewer || [], 'Zakresy nowsze w importowanym pliku'),
    '',
    formatSyncConflictEntries(conflictInfo.remoteNewer || [], 'Zakresy nowsze na serwerze')
  ].join('\n');

  return showSyncConflictDecisionDialog({
    title,
    stepLabel: 'Krok 1 z 2 • Ważny komunikat',
    bodyLines,
    preferredChoice: 'remote',
    remoteLabel: 'Anuluj import (zalecane)',
    localLabel: importActionLabel
  }).then(initialChoice => {
    if (initialChoice === 'remote') {
      return 'remote';
    }

    return showSyncConflictDecisionDialog({
      title: 'Szczegóły różnic przed importem',
      stepLabel: 'Krok 2 z 2 • Szczegóły i ostateczna decyzja',
      bodyLines: [
        'Import z pliku może nadpisać nowsze dane z serwera.',
        '',
        'Poniżej pokazano porównanie importowanej bazy z Bazą Główną.'
      ],
      detailText: detailedComparisonText,
      preferredChoice: initialChoice,
      remoteLabel: 'Anuluj import (zalecane)',
      localLabel: importActionLabel
    });
  });
}

function buildExportPayload() {
  return {
    app: 'work-tracker-html',
    version: 2,
    exportedAt: new Date().toISOString(),
    data: Store.getExportData(),
    syncMetadata: Store.getSyncMetadata ? Store.getSyncMetadata() : { common: {}, months: {} }
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

async function parseImportedDatabaseFile(file) {
  if (!file) return null;

  const content = await readFileAsText(file);
  let parsed;

  try {
    parsed = JSON.parse(content);
  } catch {
    throw new Error('Wybrany plik nie jest poprawnym plikiem JSON.');
  }

  return normalizeImportedDatabasePayload(parsed, file.name || 'plik JSON');
}

async function importDatabaseFromFile(fileOrPayload) {
  if (!fileOrPayload) return;

  const importPayload = fileOrPayload?.stateData
    ? fileOrPayload
    : await parseImportedDatabaseFile(fileOrPayload);
  const stateData = importPayload?.stateData || {};

  window.currentSheetId = null;
  window.currentWorksSheetId = null;
  Store.importState(stateData);
  setSettingsDataStatus(`Zaimportowano bazę danych z pliku: ${importPayload?.fileName || 'plik JSON'}`, 'success');
}

const SCALE_SETTINGS_CONFIG = {
  scaleLarge: {
    inputId: 'scale-large',
    valueLabelId: 'scale-large-text',
    title: 'Skala Poziomo'
  },
  scaleVertical: {
    inputId: 'scale-vertical',
    valueLabelId: 'scale-vertical-text',
    title: 'Skala Pionowo'
  }
};

let activeScaleEditorKey = '';
let activeScaleEditorAnchorRect = null;
let activeScaleMouseSession = null;
let suppressedScaleClickState = { key: '', expiresAt: 0 };
const scaleSliderTouchState = {};
const pendingScalePreviewValues = {};
let pendingScalePreviewUpdate = null;
let pendingScalePreviewFrame = null;
let isScalePinchZoomActive = false;

function getScaleEditorElements() {
  return {
    overlay: document.getElementById('scale-editor-overlay'),
    popover: document.getElementById('scale-editor-popover'),
    title: document.getElementById('scale-editor-title'),
    currentValue: document.getElementById('scale-editor-current-value'),
    range: document.getElementById('scale-editor-range')
  };
}

function getScaleSettingInput(scaleKey = '') {
  const config = SCALE_SETTINGS_CONFIG[scaleKey];
  if (!config) return null;
  return document.getElementById(config.inputId);
}

function getScaleSettingStoredValue(scaleKey = '') {
  const settings = Store.getSettings ? Store.getSettings() : { ...DEFAULT_APPEARANCE_SETTINGS, ...(Store.getState().settings || {}) };
  return clampScaleSettingValue(scaleKey, settings?.[scaleKey]);
}

function getScaleSettingAnchorElement(scaleKey = '', fallbackElement = null) {
  const sourceElement = fallbackElement || getScaleSettingInput(scaleKey);
  return sourceElement?.closest?.('.scale-settings-inline-row') || sourceElement || null;
}

function getScaleSettingAnchorRect(scaleKey = '', fallbackElement = null) {
  const anchorElement = getScaleSettingAnchorElement(scaleKey, fallbackElement);
  if (!anchorElement) return null;

  const rect = anchorElement.getBoundingClientRect();
  return {
    left: rect.left,
    top: rect.top,
    width: rect.width,
    height: rect.height
  };
}

function positionScaleEditorPopover(anchorRect = activeScaleEditorAnchorRect) {
  const { popover } = getScaleEditorElements();
  if (!popover) return;

  if (!anchorRect) {
    popover.style.left = '';
    popover.style.top = '';
    popover.style.width = '';
    popover.style.minHeight = '';
    return;
  }

  const viewportPadding = 8;
  const viewportWidth = Math.max(document.documentElement.clientWidth || 0, window.innerWidth || 0);
  const viewportHeight = Math.max(document.documentElement.clientHeight || 0, window.innerHeight || 0);
  const availableWidth = Math.max(0, viewportWidth - (viewportPadding * 2));
  const availableHeight = Math.max(0, viewportHeight - (viewportPadding * 2));
  const width = Math.min(anchorRect.width, availableWidth);
  const height = Math.min(Math.max(40, anchorRect.height), availableHeight || Math.max(40, anchorRect.height));
  const maxLeft = Math.max(viewportPadding, viewportWidth - width - viewportPadding);
  const maxTop = Math.max(viewportPadding, viewportHeight - height - viewportPadding);
  const clampedLeft = Math.min(Math.max(viewportPadding, anchorRect.left), maxLeft);
  const clampedTop = Math.min(Math.max(viewportPadding, anchorRect.top), maxTop);

  popover.style.left = `${clampedLeft}px`;
  popover.style.top = `${clampedTop}px`;
  popover.style.width = `${width}px`;
  popover.style.minHeight = `${height}px`;
}

function suppressNextScaleClick(scaleKey = '') {
  suppressedScaleClickState = {
    key: scaleKey,
    expiresAt: Date.now() + 300
  };
}

function shouldSuppressScaleClick(scaleKey = '') {
  const shouldSuppress = suppressedScaleClickState.key === scaleKey && Date.now() <= suppressedScaleClickState.expiresAt;
  if (shouldSuppress) {
    suppressedScaleClickState = { key: '', expiresAt: 0 };
  }
  return shouldSuppress;
}

function startScaleMouseSession(scaleKey = '', event = null, input = null) {
  if (!event || !input) return;

  const anchorRect = getScaleSettingAnchorRect(scaleKey, input);
  const inputRect = input.getBoundingClientRect();

  activeScaleMouseSession = {
    scaleKey,
    pointerId: event.pointerId,
    inputRect,
    anchorRect
  };

  openScaleEditor(scaleKey, { anchorRect, anchorElement: input });
  updateScaleSettingValue(scaleKey, getScaleSettingValueFromClientX(scaleKey, event.clientX, inputRect));
}

function finishScaleMouseSession(scaleKey = '', shouldCloseEditor = true) {
  if (!activeScaleMouseSession) return;

  const sessionKey = activeScaleMouseSession.scaleKey;
  activeScaleMouseSession = null;
  suppressNextScaleClick(scaleKey || sessionKey);

  if (shouldCloseEditor) {
    closeScaleEditor();
  }
}

function getScaleSettingOverlayKey() {
  return window.innerHeight > window.innerWidth ? 'scaleVertical' : 'scaleLarge';
}

function shouldScaleSettingUseOverlay(scaleKey = '') {
  return scaleKey === getScaleSettingOverlayKey();
}

function refreshScaleSettingInteractionMode() {
  Object.keys(SCALE_SETTINGS_CONFIG).forEach(scaleKey => {
    const input = getScaleSettingInput(scaleKey);
    if (!input) return;

    const usesOverlay = shouldScaleSettingUseOverlay(scaleKey);
    input.dataset.overlayMode = usesOverlay ? 'true' : 'false';
    input.style.touchAction = 'pan-y';
    input.title = usesOverlay
      ? 'Kliknij lub przesuń poziomo, aby edytować nad aplikacją.'
      : 'Na dotyku edycja przełączy się płynnie na suwak nad aplikacją.';
  });
}

function clampScaleSettingValue(scaleKey = '', value = 100) {
  const input = getScaleSettingInput(scaleKey);
  const min = parseInt(input?.min, 10) || 20;
  const max = parseInt(input?.max, 10) || 200;
  const step = parseInt(input?.step, 10) || 5;
  const parsed = parseFloat(value);
  const safeValue = Number.isFinite(parsed) ? parsed : 100;
  const clampedValue = Math.max(min, Math.min(max, safeValue));
  return min + (Math.round((clampedValue - min) / step) * step);
}

function updateScaleSettingTriggerValue(scaleKey = '', value = 100) {
  const config = SCALE_SETTINGS_CONFIG[scaleKey];
  if (!config) return;

  const normalizedValue = clampScaleSettingValue(scaleKey, value);
  const input = getScaleSettingInput(scaleKey);
  if (input) input.value = normalizedValue;

  const valueLabel = document.getElementById(config.valueLabelId);
  if (valueLabel) valueLabel.textContent = `${normalizedValue}%`;
}

function syncScaleEditorPreviewValue(scaleKey = '', value = 100) {
  if (activeScaleEditorKey !== scaleKey) return;

  const { range, currentValue } = getScaleEditorElements();
  const normalizedValue = clampScaleSettingValue(scaleKey, value);
  if (range) range.value = normalizedValue;
  if (currentValue) currentValue.textContent = `${normalizedValue}%`;
}

function applyScaleSettingPreview(scaleKey = '', value = 100) {
  if (!SCALE_SETTINGS_CONFIG[scaleKey]) return;

  const normalizedValue = clampScaleSettingValue(scaleKey, value);
  pendingScalePreviewValues[scaleKey] = normalizedValue;
  updateScaleSettingTriggerValue(scaleKey, normalizedValue);
  syncScaleEditorPreviewValue(scaleKey, normalizedValue);

  if (scaleKey === (window.innerHeight > window.innerWidth ? 'scaleVertical' : 'scaleLarge')) {
    document.documentElement.style.setProperty('--app-scale', normalizedValue / 100);
  }
}

function flushPendingScaleSettingPreview() {
  if (pendingScalePreviewFrame) {
    window.cancelAnimationFrame(pendingScalePreviewFrame);
    pendingScalePreviewFrame = null;
  }

  if (!pendingScalePreviewUpdate) return;

  const queuedUpdate = pendingScalePreviewUpdate;
  pendingScalePreviewUpdate = null;
  applyScaleSettingPreview(queuedUpdate.scaleKey, queuedUpdate.value);
}

function scheduleScaleSettingPreview(scaleKey = '', value = 100) {
  pendingScalePreviewUpdate = { scaleKey, value };
  if (pendingScalePreviewFrame) return;

  pendingScalePreviewFrame = window.requestAnimationFrame(() => {
    pendingScalePreviewFrame = null;
    const queuedUpdate = pendingScalePreviewUpdate;
    pendingScalePreviewUpdate = null;
    if (!queuedUpdate) return;
    applyScaleSettingPreview(queuedUpdate.scaleKey, queuedUpdate.value);
  });
}

function commitScaleSettingValue(scaleKey = '', fallbackValue = null) {
  if (!SCALE_SETTINGS_CONFIG[scaleKey]) return;

  flushPendingScaleSettingPreview();

  const normalizedValue = clampScaleSettingValue(
    scaleKey,
    Object.prototype.hasOwnProperty.call(pendingScalePreviewValues, scaleKey)
      ? pendingScalePreviewValues[scaleKey]
      : (fallbackValue ?? getScaleSettingStoredValue(scaleKey))
  );

  delete pendingScalePreviewValues[scaleKey];
  Store.updateAppearanceSettings({ [scaleKey]: normalizedValue });
}

function findTouchByIdentifier(touchList, identifier = null) {
  if (!touchList || identifier === null || identifier === undefined) return null;
  return Array.from(touchList).find(touch => touch.identifier === identifier) || null;
}

function updateScaleSliderTouchInteraction(scaleKey = '', touch = null, event = null) {
  const state = scaleSliderTouchState[scaleKey];
  if (!touch || !state) return;

  const deltaX = touch.clientX - state.startX;
  const deltaY = touch.clientY - state.startY;
  const absX = Math.abs(deltaX);
  const absY = Math.abs(deltaY);

  if (absY > 8 && absY > absX && !state.engaged) {
    state.allowOpen = false;
    state.movedHorizontally = false;
    return;
  }

  if (absX > 8 && absX >= absY) {
    state.movedHorizontally = true;
    if (!state.engaged) {
      state.engaged = true;
      openScaleEditor(scaleKey, { anchorRect: state.anchorRect, anchorElement: getScaleSettingInput(scaleKey) });
    }
  }

  if (state.engaged) {
    event?.preventDefault?.();
    scheduleScaleSettingPreview(scaleKey, getScaleSettingValueFromClientX(scaleKey, touch.clientX, state.inputRect));
  }
}

function finishScaleSliderTouchInteraction(scaleKey = '', options = {}) {
  const state = scaleSliderTouchState[scaleKey];
  if (!state) return;

  const { cancelled = false } = options;
  flushPendingScaleSettingPreview();

  if (!cancelled) {
    if (!state.engaged && state.allowOpen) {
      openScaleEditor(scaleKey, { anchorRect: state.anchorRect, anchorElement: getScaleSettingInput(scaleKey) });
      window.setTimeout(closeScaleEditor, 0);
    } else if (state.engaged) {
      closeScaleEditor();
    }
  } else if (state.engaged) {
    closeScaleEditor();
  }

  delete scaleSliderTouchState[scaleKey];
}

function syncScaleEditorUi() {
  const elements = getScaleEditorElements();
  const config = SCALE_SETTINGS_CONFIG[activeScaleEditorKey];
  if (!config || !elements.overlay || !elements.range || !elements.title || !elements.currentValue) return;

  const input = getScaleSettingInput(activeScaleEditorKey);
  const normalizedValue = Object.prototype.hasOwnProperty.call(pendingScalePreviewValues, activeScaleEditorKey)
    ? clampScaleSettingValue(activeScaleEditorKey, pendingScalePreviewValues[activeScaleEditorKey])
    : getScaleSettingStoredValue(activeScaleEditorKey);

  elements.title.textContent = config.title;
  elements.currentValue.textContent = `${normalizedValue}%`;
  elements.range.min = input?.min || '20';
  elements.range.max = input?.max || '300';
  elements.range.step = input?.step || '5';
  elements.range.value = normalizedValue;
}

function setScaleEditorVisibility(isVisible = false) {
  const { overlay, range } = getScaleEditorElements();
  if (!overlay) return;

  overlay.style.display = isVisible ? 'flex' : 'none';
  overlay.setAttribute('aria-hidden', isVisible ? 'false' : 'true');
  document.body.classList.toggle('scale-editor-open', isVisible);

  if (isVisible) {
    positionScaleEditorPopover();
    syncScaleEditorUi();
    window.setTimeout(() => range?.focus?.(), 0);
    return;
  }

  const closingScaleKey = activeScaleEditorKey;
  activeScaleEditorKey = '';
  activeScaleEditorAnchorRect = null;
  activeScaleMouseSession = null;
  positionScaleEditorPopover(null);
  if (closingScaleKey) {
    commitScaleSettingValue(closingScaleKey);
  }
}

function openScaleEditor(scaleKey = '', options = {}) {
  if (!SCALE_SETTINGS_CONFIG[scaleKey]) return;
  activeScaleEditorKey = scaleKey;
  activeScaleEditorAnchorRect = options.anchorRect || getScaleSettingAnchorRect(scaleKey, options.anchorElement || null);
  setScaleEditorVisibility(true);
}

function closeScaleEditor() {
  setScaleEditorVisibility(false);
}

function updateScaleSettingValue(scaleKey = '', value = 100) {
  if (!SCALE_SETTINGS_CONFIG[scaleKey]) return;

  const normalizedValue = clampScaleSettingValue(scaleKey, value);
  applyScaleSettingPreview(scaleKey, normalizedValue);
  commitScaleSettingValue(scaleKey, normalizedValue);
}

function getScaleSettingValueFromClientX(scaleKey = '', clientX = 0, inputRect = null) {
  const input = getScaleSettingInput(scaleKey);
  const rect = inputRect || input?.getBoundingClientRect?.();
  if (!input || !rect || !(rect.width > 0)) return input?.value || 100;

  const min = parseInt(input.min, 10) || 20;
  const max = parseInt(input.max, 10) || 300;
  const ratio = Math.max(0, Math.min(1, (clientX - rect.left) / rect.width));
  return clampScaleSettingValue(scaleKey, min + ((max - min) * ratio));
}

function bindScaleSettingEditor(scaleKey = '') {
  const config = SCALE_SETTINGS_CONFIG[scaleKey];
  if (!config) return;

  const input = getScaleSettingInput(scaleKey);
  if (!input || input.dataset.bound === 'true') return;

  input.dataset.bound = 'true';
  refreshScaleSettingInteractionMode();

  input.addEventListener('input', (event) => {
    if (shouldScaleSettingUseOverlay(scaleKey)) return;
    updateScaleSettingValue(scaleKey, event.target.value);
  });

  input.addEventListener('pointerdown', (event) => {
    if (event.pointerType === 'mouse' && event.button === 0) {
      event.preventDefault();
      startScaleMouseSession(scaleKey, event, input);
    }
  });

  input.addEventListener('click', (event) => {
    if (shouldSuppressScaleClick(scaleKey)) {
      event.preventDefault();
      return;
    }

    if (!shouldScaleSettingUseOverlay(scaleKey)) return;
    event.preventDefault();
    if (!activeScaleEditorKey) {
      openScaleEditor(scaleKey, { anchorElement: input });
    }
  });

  input.addEventListener('keydown', (event) => {
    if (!['Enter', ' ', 'ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) return;
    if (!shouldScaleSettingUseOverlay(scaleKey)) return;
    event.preventDefault();
    openScaleEditor(scaleKey, { anchorElement: input });
  });

  input.addEventListener('touchstart', (event) => {
    if (event.touches.length !== 1 || isScalePinchZoomActive) {
      delete scaleSliderTouchState[scaleKey];
      return;
    }

    const touch = event.touches?.[0];
    if (!touch) return;

    scaleSliderTouchState[scaleKey] = {
      startX: touch.clientX,
      startY: touch.clientY,
      touchId: touch.identifier,
      allowOpen: true,
      movedHorizontally: false,
      engaged: false,
      anchorRect: getScaleSettingAnchorRect(scaleKey, input),
      inputRect: input.getBoundingClientRect()
    };
  }, { passive: true });

  input.addEventListener('touchmove', (event) => {
    if (event.touches.length !== 1 || isScalePinchZoomActive) {
      if (scaleSliderTouchState[scaleKey]?.engaged) {
        closeScaleEditor();
      }
      delete scaleSliderTouchState[scaleKey];
      return;
    }

    const touch = event.touches?.[0];
    if (!touch) return;
    updateScaleSliderTouchInteraction(scaleKey, touch, event);
  }, { passive: false });

  input.addEventListener('touchend', () => {
    finishScaleSliderTouchInteraction(scaleKey);
  });

  input.addEventListener('touchcancel', () => {
    finishScaleSliderTouchInteraction(scaleKey, { cancelled: true });
  });
}

function initSettings() {
  const themeRadios = document.querySelectorAll('input[name="theme-select"]');
  const scaleLargeInput = document.getElementById('scale-large');
  const scaleVerticalInput = document.getElementById('scale-vertical');
  const btnReset = document.getElementById('btn-reset-settings');
  const btnResetMobile = document.getElementById('btn-reset-mobile-settings');
  const btnToggleArchiveManagement = document.getElementById('btn-toggle-archive-management');
  const archiveMonthsList = document.getElementById('settings-archive-months-list');
  const localHistoryList = document.getElementById('history-local-history-list');
  const sharedHistoryList = document.getElementById('history-shared-history-list');
  const dailyBackupsList = document.getElementById('history-daily-backups-list');
  const monthlyBackupsList = document.getElementById('history-monthly-backups-list');
  const btnHistoryForceDailyBackup = document.getElementById('btn-history-force-daily-backup');
  const btnHistoryForceMonthlyBackup = document.getElementById('btn-history-force-monthly-backup');
  const btnExportData = document.getElementById('btn-export-data');
  const btnImportData = document.getElementById('btn-import-data');
  const btnCleanupOrphanedPersonEntries = document.getElementById('btn-cleanup-orphaned-person-entries');
  const btnResetDownloadMainDatabase = document.getElementById('btn-reset-download-main-database');
  const importDataFile = document.getElementById('import-data-file');
  const btnForceResyncOff = document.getElementById('btn-force-resync-off');
  const btnForceResyncRun = document.getElementById('btn-force-resync-run');
  const btnResetFirebaseTransferStats = document.getElementById('btn-reset-firebase-transfer-stats');

  if (!themeRadios.length || !scaleLargeInput || !scaleVerticalInput) return;

  setForceResyncAdminButtonsVisibility(settingsForceResyncControlsUnlocked);
  setSettingsRemoteDatabaseResetButtonVisibility();
  updateSettingsFirebaseTransferStats();
  if (settingsForceResyncControlsUnlocked) {
    refreshForceResyncAdminStatus();
  }

  if (!window.__firebaseTransferStatsUiBound) {
    window.__firebaseTransferStatsUiBound = true;
    window.addEventListener('firebaseTransferStatsChanged', updateSettingsFirebaseTransferStats);
  }

  themeRadios.forEach(radio => {
    radio.addEventListener('change', (e) => {
      Store.updateAppearanceSettings({ theme: e.target.value });
    });
  });

  bindScaleSettingEditor('scaleLarge');
  bindScaleSettingEditor('scaleVertical');
  refreshScaleSettingInteractionMode();

  if (!window.__scaleEditorUiBound) {
    window.__scaleEditorUiBound = true;

    const { overlay, range } = getScaleEditorElements();

    if (range) {
      ['pointerdown', 'mousedown', 'touchstart', 'click'].forEach(eventName => {
        range.addEventListener(eventName, (event) => event.stopPropagation());
      });

      range.addEventListener('input', (e) => {
        if (!activeScaleEditorKey) return;
        scheduleScaleSettingPreview(activeScaleEditorKey, e.target.value);
      });

      range.addEventListener('change', (e) => {
        if (!activeScaleEditorKey) return;
        applyScaleSettingPreview(activeScaleEditorKey, e.target.value);
        commitScaleSettingValue(activeScaleEditorKey, e.target.value);
      });
    }

    if (overlay) {
      overlay.addEventListener('click', (e) => {
        if (e.target === overlay) {
          closeScaleEditor();
        }
      });
    }

    window.addEventListener('pointermove', (e) => {
      if (e.pointerType !== 'mouse' || !activeScaleMouseSession || e.pointerId !== activeScaleMouseSession.pointerId) return;
      updateScaleSettingValue(
        activeScaleMouseSession.scaleKey,
        getScaleSettingValueFromClientX(activeScaleMouseSession.scaleKey, e.clientX, activeScaleMouseSession.inputRect)
      );
    });

    ['pointerup', 'pointercancel'].forEach(eventName => {
      window.addEventListener(eventName, (e) => {
        if (e.pointerType !== 'mouse' || !activeScaleMouseSession || e.pointerId !== activeScaleMouseSession.pointerId) return;
        finishScaleMouseSession(activeScaleMouseSession.scaleKey, true);
      });
    });

    window.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && activeScaleEditorKey) {
        closeScaleEditor();
      }
    });

    window.addEventListener('touchmove', (event) => {
      Object.entries(scaleSliderTouchState).forEach(([scaleKey, state]) => {
        if (!state?.engaged) return;
        const trackedTouch = findTouchByIdentifier(event.touches, state.touchId);
        if (!trackedTouch) return;
        updateScaleSliderTouchInteraction(scaleKey, trackedTouch, event);
      });
    }, { passive: false });

    window.addEventListener('touchend', (event) => {
      Object.keys(scaleSliderTouchState).forEach(scaleKey => {
        const state = scaleSliderTouchState[scaleKey];
        if (!state) return;
        if (!findTouchByIdentifier(event.changedTouches, state.touchId)) return;
        finishScaleSliderTouchInteraction(scaleKey);
      });
    });

    window.addEventListener('touchcancel', (event) => {
      Object.keys(scaleSliderTouchState).forEach(scaleKey => {
        const state = scaleSliderTouchState[scaleKey];
        if (!state) return;
        if (!findTouchByIdentifier(event.changedTouches, state.touchId)) return;
        finishScaleSliderTouchInteraction(scaleKey, { cancelled: true });
      });
    });
  }

  if (btnReset) {
    btnReset.addEventListener('click', () => {
      if (confirm('Czy na pewno chcesz przywrócić domyślne ustawienia wyglądu?')) {
        resetAppearanceSettings('default');
      }
    });
  }

  if (btnResetFirebaseTransferStats) {
    btnResetFirebaseTransferStats.addEventListener('click', () => {
      Store.resetFirebaseTransferStats?.();
      updateSettingsFirebaseTransferStats();
    });
  }

  if (btnResetMobile) {
    btnResetMobile.addEventListener('click', () => {
      if (confirm('Czy na pewno chcesz przywrócić mobilne ustawienia wyglądu?')) {
        resetAppearanceSettings('mobile');
      }
    });
  }

  if (btnToggleArchiveManagement) {
    setArchiveManagementExpanded(false);
    btnToggleArchiveManagement.addEventListener('click', () => {
      const isExpanded = btnToggleArchiveManagement.dataset.expanded === 'true';
      setArchiveManagementExpanded(!isExpanded);
    });
  }

  if (archiveMonthsList) {
    archiveMonthsList.addEventListener('click', (e) => {
      const button = e.target.closest('.btn-manage-archive-month');
      if (!button) return;

      const month = button.getAttribute('data-month') || getSelectedMonthKey();
      const action = button.getAttribute('data-action');

      if (action === 'toggle-archive') {
        const status = Store.getMonthArchiveStatus(month);
        manageMonthArchive(status.isArchived ? 'unarchive' : 'archive', month);
        return;
      }

      if (action === 'delete-snapshot') {
        manageMonthArchive('delete-snapshot', month);
      }
    });
  }

  const bindRestoreButtons = (container, restoreType) => {
    if (!container) return;
    container.addEventListener('click', async (e) => {
      const button = e.target.closest('.btn-restore-history-entry');
      if (!button || button.getAttribute('data-restore-type') !== restoreType) return;

      const entryId = button.getAttribute('data-entry-id') || '';
      let result = { success: false, message: 'Przywracanie jest niedostępne.' };

      if (restoreType === 'local-history-undo') {
        const historyState = Store.getHistoryState ? Store.getHistoryState() : null;
        const latestEntryId = historyState?.localUndoEntries?.[0]?.id || '';
        if (!latestEntryId || latestEntryId !== entryId) return;
        result = Store.undo ? Store.undo() : result;
      } else if (restoreType === 'shared-history-before') {
        if (!confirm('Przywrócić stan sprzed tej zmiany? Zostanie zapisany nowy wpis historii.')) return;
        const entry = await ensureSharedHistoryEntryDetailLoaded(entryId);
        if (!entry?.detailsLoaded) {
          alert('Nie udało się wczytać szczegółów wpisu historii.');
          return;
        }
        result = Store.restoreSharedHistoryEntry ? Store.restoreSharedHistoryEntry(entryId) : result;
      } else if (restoreType === 'shared-history-after') {
        if (!confirm('Przywrócić stan po tej zmianie? Zostanie zapisany nowy wpis historii.')) return;
        const entry = await ensureSharedHistoryEntryDetailLoaded(entryId);
        if (!entry?.detailsLoaded) {
          alert('Nie udało się wczytać szczegółów wpisu historii.');
          return;
        }
        result = Store.restoreSharedHistoryEntryAfter ? Store.restoreSharedHistoryEntryAfter(entryId) : result;
      } else if (restoreType === 'shared-history-merge-before') {
        if (!confirm('Przywrócić usunięte wpisy z tego wpisu historii i połączyć je z aktualnym stanem tabeli?')) return;
        const entry = await ensureSharedHistoryEntryDetailLoaded(entryId);
        if (!entry?.detailsLoaded) {
          alert('Nie udało się wczytać szczegółów wpisu historii.');
          return;
        }
        result = Store.restoreSharedHistoryEntryMergeBefore ? Store.restoreSharedHistoryEntryMergeBefore(entryId) : result;
      } else if (restoreType === 'shared-history-remove-added') {
        if (!confirm('Usunąć z aktualnego stanu wpisy, które zostały dodane przez tę zmianę historii?')) return;
        const entry = await ensureSharedHistoryEntryDetailLoaded(entryId);
        if (!entry?.detailsLoaded) {
          alert('Nie udało się wczytać szczegółów wpisu historii.');
          return;
        }
        result = Store.restoreSharedHistoryEntryRemoveAdded ? Store.restoreSharedHistoryEntryRemoveAdded(entryId) : result;
      } else if (restoreType === 'daily-backup') {
        if (!confirm('Przywrócić snapshot dzienny? To przywróci cały miesiąc z daty wykonania tego snapshotu i nadpisze bieżące dane tego miesiąca.')) return;
        const entry = await ensureBackupEntryDetailLoaded('daily', entryId);
        if (!entry?.detailsLoaded) {
          alert('Nie udało się wczytać pełnych danych snapshotu dziennego.');
          return;
        }
        result = Store.restoreBackupEntry ? Store.restoreBackupEntry('daily', entryId) : result;
      } else if (restoreType === 'monthly-backup') {
        if (!confirm('Przywrócić snapshot miesięczny? To przywróci całą bazę danych z chwili wykonania tego snapshotu i nadpisze bieżące dane aplikacji.')) return;
        const entry = await ensureBackupEntryDetailLoaded('monthly', entryId);
        if (!entry?.detailsLoaded) {
          alert('Nie udało się wczytać pełnych danych snapshotu miesięcznego.');
          return;
        }
        result = Store.restoreBackupEntry ? Store.restoreBackupEntry('monthly', entryId) : result;
      }

      if (!result?.success && result?.message) {
        alert(result.message);
      }
    });
  };

  bindRestoreButtons(localHistoryList, 'local-history-undo');
  bindRestoreButtons(sharedHistoryList, 'shared-history-before');
  bindRestoreButtons(sharedHistoryList, 'shared-history-after');
  bindRestoreButtons(sharedHistoryList, 'shared-history-merge-before');
  bindRestoreButtons(sharedHistoryList, 'shared-history-remove-added');
  bindRestoreButtons(dailyBackupsList, 'daily-backup');
  bindRestoreButtons(monthlyBackupsList, 'monthly-backup');

  if (sharedHistoryList) {
    sharedHistoryList.addEventListener('click', async (e) => {
      const button = e.target.closest('.btn-toggle-history-entry-details');
      if (!button) return;

      const entryId = button.getAttribute('data-entry-id') || '';
      if (!entryId) return;

      const detailsState = getSharedHistoryEntryDetailsState(entryId);
      detailsState.expanded = !detailsState.expanded;
      detailsState.error = '';

      if (!detailsState.expanded) {
        renderHistoryManagement();
        return;
      }

      const existingEntry = Store.getSharedHistoryEntry?.(entryId);
      if (existingEntry?.detailsLoaded === true) {
        renderHistoryManagement();
        return;
      }

      detailsState.loading = true;
      renderHistoryManagement();

      try {
        const entry = await ensureSharedHistoryEntryDetailLoaded(entryId);
        if (!entry?.detailsLoaded) {
          detailsState.error = 'Nie udało się wczytać szczegółów tego wpisu.';
        }
      } catch (error) {
        detailsState.error = error instanceof Error ? error.message : 'Nie udało się wczytać szczegółów tego wpisu.';
      } finally {
        detailsState.loading = false;
        renderHistoryManagement();
      }
    });
  }

  if (btnHistoryForceDailyBackup) {
    btnHistoryForceDailyBackup.addEventListener('click', () => {
      if (!confirm('Wykonać Snapshot Dzienny i nadpisać istniejący snapshot z dzisiaj?')) return;
      const result = Store.createOrReplaceDailyBackup
        ? Store.createOrReplaceDailyBackup()
        : { success: false, message: 'Ręczne tworzenie snapshotu dziennego jest niedostępne.' };
      if (!result?.success && result?.message) {
        alert(result.message);
      }
    });
  }

  if (btnHistoryForceMonthlyBackup) {
    btnHistoryForceMonthlyBackup.addEventListener('click', () => {
      if (!confirm('Wykonać Snapshot Miesięczny i nadpisać istniejący snapshot z bieżącego miesiąca?')) return;
      const result = Store.createOrReplaceMonthlyBackup
        ? Store.createOrReplaceMonthlyBackup()
        : { success: false, message: 'Ręczne tworzenie snapshotu miesięcznego jest niedostępne.' };
      if (!result?.success && result?.message) {
        alert(result.message);
      }
    });
  }

  updateHistorySecretSnapshotActionVisibility();

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

      try {
        setSettingsDataStatus('Trwa import bazy danych...', '');
        const importPayload = await parseImportedDatabaseFile(file);
        const importedStateForComparison = Store.importState
          ? Store.importState(importPayload.stateData, true)
          : (importPayload.stateData || {});
        const shouldSyncToFirebase = typeof firebase !== 'undefined'
          && firebase.auth
          && firebase.auth().currentUser
          && !window.isImportingFromFirebase
          && !window.isOfflineMode;
        const remoteBundleBeforeImport = shouldSyncToFirebase
          ? await fetchRemoteStateBundleForSharedHistory()
          : null;

        let isImportApproved = false;
        if (shouldSyncToFirebase && remoteBundleBeforeImport) {
          const importWarningContext = buildImportSyncWarningContext(importPayload, importedStateForComparison, remoteBundleBeforeImport);
          if (importWarningContext.shouldWarn) {
            const choice = await askUserHowToResolveImportedDatabaseConflict(
              importWarningContext.conflictInfo,
              importedStateForComparison,
              remoteBundleBeforeImport.state,
              {
                fileName: file.name,
                importedLatestTimestamp: importWarningContext.importedLatestTimestamp,
                remoteLatestTimestamp: importWarningContext.remoteLatestTimestamp,
                hasImportedNewerChanges: importWarningContext.hasImportedNewerChanges,
                hasRemoteNewerChanges: importWarningContext.hasRemoteNewerChanges
              }
            );

            if (choice !== 'local') {
              setSettingsDataStatus('Import został anulowany. Pozostawiono aktualną Bazę Główną.', '');
              return;
            }

            isImportApproved = true;
          }
        }

        if (!isImportApproved) {
          const importConfirmMessage = shouldSyncToFirebase
            ? `Zaimportować bazę danych z pliku "${file.name}"? Obecne dane aplikacji zostaną nadpisane i wysłane do chmury Firebase.`
            : `Zaimportować bazę danych z pliku "${file.name}"? Obecne dane aplikacji zostaną nadpisane lokalnie.`;
          if (!confirm(importConfirmMessage)) {
            return;
          }
        }

        await importDatabaseFromFile(importPayload);
        
        if (shouldSyncToFirebase) {
          const exportData = Store.getExportData();
          const syncMetadata = Store.getSyncMetadata ? Store.getSyncMetadata() : { common: {}, months: {} };
          const updates = buildFirebaseStateUpdates(exportData, syncMetadata);
          const syncOverwriteHistoryResult = (remoteBundleBeforeImport && Store.recordSharedSyncOverwriteHistory)
            ? Store.recordSharedSyncOverwriteHistory(
                remoteBundleBeforeImport.state,
                exportData,
                remoteBundleBeforeImport.meta,
                syncMetadata,
                `Zaimportowano bazę danych z pliku: ${file.name}`
              )
            : { success: false, history: [] };

          if (syncOverwriteHistoryResult?.success && syncOverwriteHistoryResult.entry) {
            Object.assign(updates, Store.buildSharedHistoryMigrationUpdates?.([syncOverwriteHistoryResult.entry]) || {});
          }

          await appUpdateFirebaseRootWithFallback(updates);
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

  if (btnCleanupOrphanedPersonEntries) {
    btnCleanupOrphanedPersonEntries.addEventListener('click', () => {
      const month = Store.getSelectedMonth();
      const monthLabel = (month || '').toString();
      if (!confirm(`Usunąć osierocone wpisy osób z aktywnego miesiąca ${monthLabel}?`)) return;

      try {
        setSettingsDataStatus('Trwa czyszczenie osieroconych wpisów osób...', '');
        const result = Store.cleanupOrphanedPersonEntries
          ? Store.cleanupOrphanedPersonEntries(month)
          : { success: false, message: 'Mechanizm czyszczenia osieroconych wpisów nie jest dostępny.' };

        if (result?.success === false) {
          setSettingsDataStatus(result.message || 'Nie udało się wyczyścić osieroconych wpisów osób.', 'error');
          return;
        }

        setSettingsDataStatus(formatOrphanedPersonCleanupResultMessage(result), result?.changed ? 'success' : '');
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Nie udało się wyczyścić osieroconych wpisów osób.';
        setSettingsDataStatus(message, 'error');
        alert(message);
      }
    });
  }

  if (btnResetDownloadMainDatabase) {
    btnResetDownloadMainDatabase.addEventListener('click', async () => {
      if (!confirm('Czy na pewno chcesz wyczyścić lokalną bazę danych i pobrać od nowa pełną Bazę Główną z Firebase?')) return;

      try {
        btnResetDownloadMainDatabase.disabled = true;
        setSettingsDataStatus('Trwa reset lokalnej bazy i pobieranie Bazy Głównej z Firebase...', '');
        await resetLocalDatabaseAndFetchMainDatabase();
        setSettingsDataStatus('Lokalna baza została zresetowana i odbudowana z aktualnej Bazy Głównej.', 'success');
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Nie udało się pobrać Bazy Głównej z Firebase.';
        setSettingsDataStatus(message, 'error');
        alert(message);
      } finally {
        btnResetDownloadMainDatabase.disabled = false;
      }
    });
  }

  if (btnForceResyncOff) {
    btnForceResyncOff.addEventListener('click', async () => {
      if (!confirm('Ustawić force_resync na OFF? Nowe logowania nie będą już wymuszały resetu lokalnej bazy.')) return;

      try {
        setForceResyncAdminButtonsDisabled(true);
        setForceResyncAdminStatus('Wyłączam force_resync...');
        await setForceResyncControlEnabled(false);
        setForceResyncAdminStatus('force_resync został ustawiony na OFF.', 'success');
        await refreshForceResyncAdminStatus();
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Nie udało się wyłączyć force_resync.';
        setForceResyncAdminStatus(message, 'error');
        alert(message);
      } finally {
        setForceResyncAdminButtonsDisabled(false);
      }
    });
  }

  if (btnForceResyncRun) {
    btnForceResyncRun.addEventListener('click', async () => {
      if (!confirm('Wykonać force_resync dla wszystkich użytkowników? Lista e-maili, które już wykonały reset, zostanie wyczyszczona.')) return;

      try {
        setForceResyncAdminButtonsDisabled(true);
        setForceResyncAdminStatus('Włączam force_resync i czyszczę listę wykonanych resetów...');
        await setForceResyncControlEnabled(true);
        setForceResyncAdminStatus('force_resync został włączony. Przy następnym logowaniu użytkownicy pobiorą bazę od nowa.', 'success');
        await refreshForceResyncAdminStatus();
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Nie udało się włączyć force_resync.';
        setForceResyncAdminStatus(message, 'error');
        alert(message);
      } finally {
        setForceResyncAdminButtonsDisabled(false);
      }
    });
  }

  renderArchiveMonthManagement();
  renderHistoryManagement();
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

  updateScaleSettingTriggerValue('scaleLarge', settings.scaleLarge);
  updateScaleSettingTriggerValue('scaleVertical', settings.scaleVertical);
  refreshScaleSettingInteractionMode();
  syncScaleEditorUi();

  // Apply Scale
  const isVertical = window.innerHeight > window.innerWidth;
  const currentScale = isVertical ? settings.scaleVertical : settings.scaleLarge;
  root.style.setProperty('--app-scale', currentScale / 100);
}

let scaleOverlayTimeout = null;
function showScaleOverlay(scalePercent, options = {}) {
  const overlay = document.getElementById('scale-info-overlay');
  if (!overlay) return;

  overlay.textContent = `${Math.round(scalePercent)}%`;
  overlay.classList.add('visible');

  if (options.persistent === true) {
    clearTimeout(scaleOverlayTimeout);
    return;
  }

  clearTimeout(scaleOverlayTimeout);
  scaleOverlayTimeout = setTimeout(() => {
    overlay.classList.remove('visible');
  }, 1000);
}

function initInteractiveScaling() {
  const clampAppScalePercent = (value) => Math.max(20, Math.min(200, Math.round(value)));
  const getCurrentScaleKey = () => (window.innerHeight > window.innerWidth ? 'scaleVertical' : 'scaleLarge');
  const applyLiveScalePreview = (scalePercent) => {
    document.documentElement.style.setProperty('--app-scale', scalePercent / 100);
  };

  // 1. Ctrl + Mouse Wheel
  window.addEventListener('wheel', (e) => {
    if (e.ctrlKey) {
      e.preventDefault();
      const settings = Store.getSettings ? Store.getSettings() : Store.getState().settings;
      const isVertical = window.innerHeight > window.innerWidth;
      const currentScale = isVertical ? settings.scaleVertical : settings.scaleLarge;
      
      let nextScale = currentScale - (e.deltaY > 0 ? 5 : -5);
      nextScale = Math.max(20, Math.min(200, nextScale));

      if (nextScale !== currentScale) {
        const update = isVertical ? { scaleVertical: nextScale } : { scaleLarge: nextScale };

        const mainContent = document.querySelector('.main-content');
        let oldLogicalX = 0, oldLogicalY = 0;
        let rect = null;
        if (mainContent) {
          rect = mainContent.getBoundingClientRect();
          const oldScaleFactor = currentScale / 100;
          oldLogicalX = mainContent.scrollLeft + (e.clientX - rect.left) / oldScaleFactor;
          oldLogicalY = mainContent.scrollTop + (e.clientY - rect.top) / oldScaleFactor;
        }

        Store.updateAppearanceSettings(update);
        showScaleOverlay(nextScale);

        if (mainContent && rect) {
          requestAnimationFrame(() => {
            const newScaleFactor = nextScale / 100;
            const updatedRect = mainContent.getBoundingClientRect();
            mainContent.scrollLeft = oldLogicalX - (e.clientX - updatedRect.left) / newScaleFactor;
            mainContent.scrollTop = oldLogicalY - (e.clientY - updatedRect.top) / newScaleFactor;
          });
        }
      }
    }
  }, { passive: false });

  // 2. Pinch to Zoom
  let initialPinchDistance = null;
  let initialScale = null;
  let pinchScaleKey = 'scaleLarge';
  let pinchPreviewScale = null;
  let pinchFrame = null;
  let initialPinchLogicalX = null;
  let initialPinchLogicalY = null;
  let currentPinchCenterX = null;
  let currentPinchCenterY = null;

  const flushPinchPreview = () => {
    pinchFrame = null;
    if (!Number.isFinite(pinchPreviewScale)) return;

    applyLiveScalePreview(pinchPreviewScale);
    showScaleOverlay(pinchPreviewScale, { persistent: true });

    const mainContent = document.querySelector('.main-content');
    if (mainContent && initialPinchLogicalX !== null && currentPinchCenterX !== null) {
      const newScaleFactor = pinchPreviewScale / 100;
      const updatedRect = mainContent.getBoundingClientRect();
      const targetScrollLeft = initialPinchLogicalX - (currentPinchCenterX - updatedRect.left) / newScaleFactor;
      const targetScrollTop = initialPinchLogicalY - (currentPinchCenterY - updatedRect.top) / newScaleFactor;
      mainContent.scrollLeft = targetScrollLeft;
      mainContent.scrollTop = targetScrollTop;
    }
  };

  const schedulePinchPreview = () => {
    if (pinchFrame) return;
    pinchFrame = window.requestAnimationFrame(flushPinchPreview);
  };

  const finishPinchScaling = () => {
    if (initialPinchDistance === null) return;

    if (pinchFrame) {
      window.cancelAnimationFrame(pinchFrame);
      flushPinchPreview();
    }

    const finalScale = clampAppScalePercent(pinchPreviewScale ?? initialScale ?? 100);
    if (Math.abs(finalScale - (initialScale ?? finalScale)) >= 1) {
      Store.updateAppearanceSettings({ [pinchScaleKey]: finalScale });
      showScaleOverlay(finalScale);
    } else {
      applySettings();
    }

    initialPinchDistance = null;
    initialScale = null;
    pinchPreviewScale = null;
    pinchScaleKey = 'scaleLarge';
    isScalePinchZoomActive = false;
    initialPinchLogicalX = null;
    initialPinchLogicalY = null;
    currentPinchCenterX = null;
    currentPinchCenterY = null;
  };

  window.addEventListener('touchstart', (e) => {
    if (e.touches.length === 2) {
      isScalePinchZoomActive = true;
      if (activeScaleEditorKey) {
        closeScaleEditor();
      }
      const settings = Store.getSettings ? Store.getSettings() : Store.getState().settings;
      pinchScaleKey = getCurrentScaleKey();
      initialScale = settings[pinchScaleKey];
      pinchPreviewScale = initialScale;
      initialPinchDistance = Math.hypot(
        e.touches[0].pageX - e.touches[1].pageX,
        e.touches[0].pageY - e.touches[1].pageY
      );

      currentPinchCenterX = (e.touches[0].clientX + e.touches[1].clientX) / 2;
      currentPinchCenterY = (e.touches[0].clientY + e.touches[1].clientY) / 2;

      const mainContent = document.querySelector('.main-content');
      if (mainContent) {
        const rect = mainContent.getBoundingClientRect();
        const scaleFactor = initialScale / 100;
        initialPinchLogicalX = mainContent.scrollLeft + (currentPinchCenterX - rect.left) / scaleFactor;
        initialPinchLogicalY = mainContent.scrollTop + (currentPinchCenterY - rect.top) / scaleFactor;
      }
    }
  }, { passive: true });

  window.addEventListener('touchmove', (e) => {
    if (e.touches.length === 2 && initialPinchDistance !== null) {
      e.preventDefault();
      const currentDistance = Math.hypot(
        e.touches[0].pageX - e.touches[1].pageX,
        e.touches[0].pageY - e.touches[1].pageY
      );

      currentPinchCenterX = (e.touches[0].clientX + e.touches[1].clientX) / 2;
      currentPinchCenterY = (e.touches[0].clientY + e.touches[1].clientY) / 2;

      const factor = currentDistance / initialPinchDistance;
      const nextScale = clampAppScalePercent(initialScale * factor);

      if (Math.abs(nextScale - (pinchPreviewScale ?? initialScale)) >= 1) {
        pinchPreviewScale = nextScale;
        schedulePinchPreview();
      }
    }
  }, { passive: false });

  window.addEventListener('touchend', (e) => {
    if (e.touches.length < 2) {
      finishPinchScaling();
    }
  });

  window.addEventListener('touchcancel', finishPinchScaling);

  // Handle Resize/Orientation change
  window.addEventListener('resize', () => {
    refreshScaleSettingInteractionMode();
    if (activeScaleEditorKey && !shouldScaleSettingUseOverlay(activeScaleEditorKey)) {
      closeScaleEditor();
    }
    applySettings();
  });
}
