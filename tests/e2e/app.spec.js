const { test, expect } = require('@playwright/test');
const { mockFirebaseAuth } = require('../helpers/firebase-mock');

test.describe('Aplikacja — po zalogowaniu (mock)', () => {
  test.beforeEach(async ({ page }) => {
    await mockFirebaseAuth(page);
    await page.goto('/');
    // Czekamy aż app-root stanie się widoczny (auth callback uruchomiony)
    await page.waitForSelector('#app-root', { state: 'visible', timeout: 10_000 });
  });

  test('ukrywa overlay logowania po autentykacji', async ({ page }) => {
    const overlay = page.locator('#firebase-auth-overlay');
    await expect(overlay).toBeHidden();
  });

  test('wyświetla główny kontener aplikacji', async ({ page }) => {
    await expect(page.locator('#app-root')).toBeVisible();
  });

  test('widok Podsumowanie jest domyślnie aktywny', async ({ page }) => {
    const summaryView = page.locator('#summary-view');
    await expect(summaryView).toBeVisible();
  });

  test('pasek boczny zawiera linki nawigacyjne', async ({ page }) => {
    const navItems = page.locator('.nav-item');
    const count = await navItems.count();
    expect(count).toBeGreaterThanOrEqual(8);
  });

  test('nawigacja do widoku Osoby działa', async ({ page }) => {
    await page.click('[data-target="people-view"]');
    await expect(page.locator('#people-view')).toBeVisible();
    await expect(page.locator('#summary-view')).toBeHidden();
  });

  test('nawigacja do widoku Klienci działa', async ({ page }) => {
    await page.click('[data-target="clients-view"]');
    await expect(page.locator('#clients-view')).toBeVisible();
  });

  test('nawigacja do widoku Godziny działa', async ({ page }) => {
    await page.click('[data-target="hours-view"]');
    await expect(page.locator('#hours-view')).toBeVisible();
  });

  test('nawigacja do widoku Rozliczenie działa', async ({ page }) => {
    await page.click('[data-target="settlement-view"]');
    await expect(page.locator('#settlement-view')).toBeVisible();
  });

  test('nawigacja do widoku Wypłaty działa', async ({ page }) => {
    await page.click('[data-target="payouts-view"]');
    await expect(page.locator('#payouts-view')).toBeVisible();
  });

  test('nawigacja do widoku Ustawienia działa', async ({ page }) => {
    await page.click('[data-target="settings-view"]');
    await expect(page.locator('#settings-view')).toBeVisible();
  });

  test('selektor miesiąca jest widoczny', async ({ page }) => {
    const monthSelect = page.locator('#global-month-select');
    await expect(monthSelect).toBeVisible();
    const value = await monthSelect.inputValue();
    expect(value).toMatch(/^\d{4}-\d{2}$/);
  });

  test('zmiana miesiąca aktualizuje wartość selektora', async ({ page }) => {
    const monthSelect = page.locator('#global-month-select');
    await monthSelect.fill('2024-06');
    await monthSelect.dispatchEvent('change');
    const value = await monthSelect.inputValue();
    expect(value).toBe('2024-06');
  });

  test('tytuł strony to Monitor pracy', async ({ page }) => {
    await expect(page).toHaveTitle(/Monitor pracy/i);
  });
});

test.describe('Aplikacja — stan niezalogowany', () => {
  test('pokazuje overlay logowania gdy brak sesji', async ({ page }) => {
    const { mockFirebaseUnauthenticated } = require('../helpers/firebase-mock');
    await mockFirebaseUnauthenticated(page);
    await page.goto('/');
    await expect(page.locator('#firebase-auth-overlay')).toBeVisible();
    await expect(page.locator('#app-root')).toBeHidden();
  });
});
