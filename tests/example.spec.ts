import { test, expect } from '@playwright/test';
import { authenticator } from 'otplib';

test('has title', async ({ page }) => {
  const username = process.env.TEST_USERNAME;
  const password = process.env.TEST_PASSWORD;
  const secret =  process.env.TEST_TOKEN;
  const file =  process.env.TEST_URL;
  await page.goto(file);
  await page.getByRole('textbox').fill(username);
  await page.getByRole('button', { name: 'Next' }).click();
  await page.locator('input[type="password"]').fill(password);
  await page.getByRole('button', { name: 'Sign in' }).click();
  const token = authenticator.generate(secret);
  await page.getByLabel('code').fill(token);
  await page.getByRole('button', { name: 'Verify' }).click();
  await page.getByRole('button', { name: 'Yes' }).click();
  await page.waitForLoadState('domcontentloaded');
  await page.waitForTimeout(15000);
  await page.locator('iframe[name="WacFrame_Excel_0"]').contentFrame().getByRole('button', { name: 'File', exact: true }).click();
  await page.locator('iframe[name="WacFrame_Excel_0"]').contentFrame().getByText('Create a Copy').click();
  const downloadPromise = page.waitForEvent('download');
  await page.locator('iframe[name="WacFrame_Excel_0"]').contentFrame().getByText('Download a Copy').click();
  const download = await downloadPromise;
  await download.saveAs('data/' + download.suggestedFilename());
});