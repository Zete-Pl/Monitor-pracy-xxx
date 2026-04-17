const { test, expect } = require('@playwright/test');
const { mockFirebaseUnauthenticated } = require('../helpers/firebase-mock');

test.describe('Bezpieczeństwo — OWASP', () => {
  test.beforeEach(async ({ page }) => {
    await mockFirebaseUnauthenticated(page);
    await page.goto('/');
  });

  test('A05: Content-Security-Policy jest zdefiniowana', async ({ page }) => {
    const csp = await page.$eval(
      'meta[http-equiv="Content-Security-Policy"]',
      el => el.getAttribute('content')
    );
    expect(csp).toBeTruthy();
    expect(csp).toContain('default-src');
    expect(csp).toContain('frame-ancestors');
  });

  test('A05: X-Frame-Options jest ustawione na DENY', async ({ page }) => {
    const val = await page.$eval(
      'meta[http-equiv="X-Frame-Options"]',
      el => el.getAttribute('content')
    );
    expect(val).toBe('DENY');
  });

  test('A05: X-Content-Type-Options jest ustawione na nosniff', async ({ page }) => {
    const val = await page.$eval(
      'meta[http-equiv="X-Content-Type-Options"]',
      el => el.getAttribute('content')
    );
    expect(val).toBe('nosniff');
  });

  test('A05: Referrer-Policy jest skonfigurowana', async ({ page }) => {
    const val = await page.$eval(
      'meta[http-equiv="Referrer-Policy"]',
      el => el.getAttribute('content')
    );
    expect(val).toBeTruthy();
  });

  test('A06: Firebase config ładuje się z zewnętrznego pliku (window.firebaseConfig)', async ({ page }) => {
    const config = await page.evaluate(() => window.firebaseConfig);
    expect(config).toBeTruthy();
    expect(config.apiKey).toBeTruthy();
    expect(config.projectId).toBeTruthy();
  });

  test('A06: Lucide nie używa @latest', async ({ page }) => {
    const src = await page.$eval(
      'script[src*="lucide"]',
      el => el.getAttribute('src')
    );
    expect(src).not.toContain('@latest');
    expect(src).toMatch(/@\d+\.\d+\.\d+/);
  });

  test('A06: Skrypty CDN mają atrybut crossorigin', async ({ page }) => {
    const scripts = await page.$$eval(
      'script[src*="gstatic.com"], script[src*="unpkg.com"], script[src*="jsdelivr.net"]',
      els => els.map(el => ({ src: el.src, crossorigin: el.crossOrigin }))
    );
    expect(scripts.length).toBeGreaterThan(0);
    for (const s of scripts) {
      expect(s.crossorigin, `Brak crossorigin na: ${s.src}`).toBeTruthy();
    }
  });

  test('A03: Strona ładuje się bez błędów JS w konsoli', async ({ page }) => {
    const errors = [];
    page.on('pageerror', err => errors.push(err.message));
    await page.reload();
    await page.waitForSelector('#firebase-auth-overlay', { state: 'visible' });
    const critical = errors.filter(e => !e.includes('firebase') && !e.includes('lucide'));
    expect(critical).toHaveLength(0);
  });

  test('A03: Klucze API nie są zakodowane w js/app.js (sprawdzenie przez response)', async ({ page }) => {
    const response = await page.goto('/js/app.js');
    const body = await response.text();
    expect(body).not.toContain('AIzaSy');
    expect(body).not.toContain('firebaseConfig = {');
  });
});
