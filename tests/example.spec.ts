import { test, expect } from '@playwright/test';
import { authenticator } from 'otplib';

test('has title', async ({ page }) => {
  const password = process.env.TEST_PASSWORD;
  const secret =  process.env.TEST_TOKEN;
  await page.goto('https://clearroute24.sharepoint.com/_layouts/15/sharepoint.aspx?login_hint=gary.wong%40clearroute.io');
  await page.getByRole('textbox', { name: 'Enter the password for gary.' }).fill(password);
  await page.getByRole('button', { name: 'Sign in' }).click();
  const token = authenticator.generate(secret);
  await page.getByLabel('code').fill(token);
  await page.getByRole('button', { name: 'Verify' }).click();
  await page.getByRole('button', { name: 'Yes' }).click();
  await expect(page).toHaveTitle('SharePoint');
  await page.goto('https://clearroute24-my.sharepoint.com/:x:/r/personal/gary_wong_clearroute_io/_layouts/15/Doc.aspx?sourcedoc=%7BC1F4785E-D56D-4F04-9C83-01067D50AA6E%7D&file=Report%20Leave%20requests.xlsx&action=default&mobileredirect=true&DefaultItemOpen=1');
  await page.locator('iframe[name="WacFrame_Excel_0"]').contentFrame().getByRole('button', { name: 'File' }).click();
  await page.locator('iframe[name="WacFrame_Excel_0"]').contentFrame().getByText('Create a Copy').click();
  const downloadPromise = page.waitForEvent('download');
  await page.locator('iframe[name="WacFrame_Excel_0"]').contentFrame().getByText('Download a Copy').click();
  const download = await downloadPromise;
  await download.saveAs('data/' + download.suggestedFilename());
});