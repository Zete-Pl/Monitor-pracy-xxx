const { test, expect } = require('@playwright/test');
const { mockFirebaseUnauthenticated } = require('../helpers/firebase-mock');

test.describe('Strona logowania', () => {
  test.beforeEach(async ({ page }) => {
    await mockFirebaseUnauthenticated(page);
    await page.goto('/');
    await page.waitForSelector('#firebase-auth-overlay', { state: 'visible' });
  });

  test('wyświetla overlay logowania dla niezalogowanego użytkownika', async ({ page }) => {
    const overlay = page.locator('#firebase-auth-overlay');
    await expect(overlay).toBeVisible();
  });

  test('zawiera formularz z polem email i hasło', async ({ page }) => {
    await expect(page.locator('#firebase-email')).toBeVisible();
    await expect(page.locator('#firebase-password')).toBeVisible();
    await expect(page.locator('#btn-firebase-email-login')).toBeVisible();
  });

  test('zawiera przycisk logowania przez Google', async ({ page }) => {
    await expect(page.locator('#btn-firebase-google')).toBeVisible();
  });

  test('pokazuje błąd gdy formularz jest pusty', async ({ page }) => {
    await page.locator('#btn-firebase-email-login').click();
    const errorMsg = page.locator('#firebase-auth-error');
    await expect(errorMsg).toBeVisible();
    await expect(errorMsg).toContainText('Wprowadź e-mail i hasło');
  });

  test('pokazuje błąd przy nieprawidłowych danych logowania', async ({ page }) => {
    await page.fill('#firebase-email', 'test@example.com');
    await page.fill('#firebase-password', 'wrongpassword');
    await page.locator('#btn-firebase-email-login').click();
    const errorMsg = page.locator('#firebase-auth-error');
    await expect(errorMsg).toBeVisible({ timeout: 5000 });
    await expect(errorMsg).toContainText('Błędne hasło');
  });

  test('pole email ma typ email (HTML5 validation)', async ({ page }) => {
    const type = await page.locator('#firebase-email').getAttribute('type');
    expect(type).toBe('email');
  });

  test('pole hasła ma typ password', async ({ page }) => {
    const type = await page.locator('#firebase-password').getAttribute('type');
    expect(type).toBe('password');
  });
});

test.describe('Rate limiting logowania', () => {
  test('blokuje logowanie po 5 nieudanych próbach', async ({ page }) => {
    await mockFirebaseUnauthenticated(page);
    await page.goto('/');
    await page.waitForSelector('#firebase-auth-overlay', { state: 'visible' });

    for (let i = 0; i < 5; i++) {
      await page.fill('#firebase-email', 'test@example.com');
      await page.fill('#firebase-password', `wrongpassword${i}`);
      await page.locator('#btn-firebase-email-login').click();
      await page.waitForSelector('#firebase-auth-error', { state: 'visible' });
      await page.waitForTimeout(100);
    }

    // 6. próba — powinna pokazać komunikat o blokadzie
    await page.fill('#firebase-email', 'test@example.com');
    await page.fill('#firebase-password', 'anotherpassword');
    await page.locator('#btn-firebase-email-login').click();

    const errorMsg = page.locator('#firebase-auth-error');
    await expect(errorMsg).toContainText('Zbyt wiele prób logowania');
  });
});
