const Persistence = { LOCAL: 'LOCAL', SESSION: 'SESSION', NONE: 'NONE' };

function buildAuthClass() {
  const AuthClass = Object.assign(() => {}, {
    GoogleAuthProvider: class { constructor() {} },
    EmailAuthProvider: class { constructor() {} },
    Auth: { Persistence },
    Persistence,
  });
  return AuthClass;
}

/**
 * Injects a Firebase mock into the page before any scripts run.
 * Simulates a logged-in user so the app renders past the auth overlay.
 */
async function mockFirebaseAuth(page, { email = 'zete777@gmail.com', uid = 'test-uid-001' } = {}) {
  await page.addInitScript(({ email, uid }) => {
    const Persistence = { LOCAL: 'LOCAL', SESSION: 'SESSION', NONE: 'NONE' };

    const fakeUser = {
      uid,
      email,
      displayName: 'Test User',
      getIdToken: () => Promise.resolve('fake-id-token'),
    };

    const authListeners = [];
    let currentUser = null;

    const fakeAuth = {
      currentUser: null,
      onAuthStateChanged(cb) {
        authListeners.push(cb);
        setTimeout(() => cb(currentUser), 0);
        return () => {};
      },
      signInWithEmailAndPassword() {
        return Promise.resolve({ user: fakeUser });
      },
      signInWithPopup() {
        return Promise.resolve({ user: fakeUser });
      },
      signOut() {
        currentUser = null;
        fakeAuth.currentUser = null;
        authListeners.forEach(cb => cb(null));
        return Promise.resolve();
      },
      setPersistence() { return Promise.resolve(); },
    };

    const fakeDatabase = () => ({
      ref(path) {
        const ref = {
          once() { return Promise.resolve({ val: () => null, exists: () => false }); },
          on(event, cb) { setTimeout(() => cb({ val: () => null, exists: () => false }), 50); return () => {}; },
          off() {},
          set() { return Promise.resolve(); },
          update() { return Promise.resolve(); },
          remove() { return Promise.resolve(); },
          child(p) { return fakeDatabase().ref((path || '') + '/' + p); },
        };
        return ref;
      },
    });

    const AuthFn = Object.assign(() => fakeAuth, {
      GoogleAuthProvider: class { constructor() {} },
      EmailAuthProvider: class { constructor() {} },
      Auth: { Persistence },
      Persistence,
    });

    window.__mockFirebase = {
      apps: [],
      auth: AuthFn,
      database: fakeDatabase,
      initializeApp(config) {
        window.__mockFirebase.apps.push({});
        setTimeout(() => {
          currentUser = fakeUser;
          fakeAuth.currentUser = fakeUser;
          authListeners.forEach(cb => cb(fakeUser));
        }, 50);
        return {};
      },
    };

    Object.defineProperty(window, 'firebase', {
      get() { return window.__mockFirebase; },
      configurable: true,
    });
  }, { email, uid });
}

/**
 * Mocks Firebase with no logged-in user (shows login overlay).
 */
async function mockFirebaseUnauthenticated(page) {
  await page.addInitScript(() => {
    const Persistence = { LOCAL: 'LOCAL', SESSION: 'SESSION', NONE: 'NONE' };
    const authListeners = [];

    const fakeAuth = {
      currentUser: null,
      onAuthStateChanged(cb) {
        authListeners.push(cb);
        setTimeout(() => cb(null), 0);
        return () => {};
      },
      signInWithEmailAndPassword() {
        return Promise.reject(Object.assign(new Error('WRONG_PASSWORD'), { code: 'auth/wrong-password' }));
      },
      signInWithPopup() {
        return Promise.reject(Object.assign(new Error('CANCELLED'), { code: 'auth/popup-closed-by-user' }));
      },
      signOut() { return Promise.resolve(); },
      setPersistence() { return Promise.resolve(); },
    };

    const fakeDatabase = () => ({
      ref() {
        return {
          once() { return Promise.resolve({ val: () => null, exists: () => false }); },
          on() { return () => {}; },
          off() {},
          set() { return Promise.resolve(); },
          update() { return Promise.resolve(); },
        };
      },
    });

    const AuthFn = Object.assign(() => fakeAuth, {
      GoogleAuthProvider: class {},
      EmailAuthProvider: class {},
      Auth: { Persistence },
      Persistence,
    });

    window.__mockFirebase = {
      apps: [],
      auth: AuthFn,
      database: fakeDatabase,
      initializeApp() {
        window.__mockFirebase.apps.push({});
        return {};
      },
    };

    Object.defineProperty(window, 'firebase', {
      get() { return window.__mockFirebase; },
      configurable: true,
    });
  });
}

module.exports = { mockFirebaseAuth, mockFirebaseUnauthenticated };
