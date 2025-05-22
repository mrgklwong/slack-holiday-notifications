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
  await page.getByRole('textbox', { name: 'Enter the password for gary.' }).fill(password);
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

//
//https://clearroute24-my.sharepoint.com/:x:/g/personal/gary_wong_clearroute_io/EX3_Vs5eL-lOtlXgJfpZcooB1GVQeUhwhF3pxp0uVSZPKg?e=NYjOc3
//https://clearroute24-my.sharepoint.com/personal/gary_wong_clearroute_io/_layouts/15/Doc.aspx?sourcedoc=%7BC1F4785E-D56D-4F04-9C83-01067D50AA6E%7D&file=Report%20Leave%20requests.xlsx&action=default&mobileredirect=true