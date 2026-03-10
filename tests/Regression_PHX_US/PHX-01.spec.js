import { test, expect } from '@playwright/test';
test.afterEach(async ({ page }, testInfo) => {
  await testInfo.attach('Full Page Screenshot', { body: await page.screenshot({ fullPage: true }),contentType: 'image/png',});
});

test('PHX-01 OrderCheckoutFlow_ShipForMe(AddedNote)_MasterCreditCard_U3', async ({ page }) => {
  await page.goto('https://stage10.phenomenex.com/');
  await page.getByRole('button', { name: 'Accept All Cookies' }).click();
  await page.getByRole('link', { name: 'Sign In' }).click();

  await page.getByRole('textbox', { name: 'Email Address' }).fill('U3_stage_sit@yopmail.com');

  await page.getByRole('textbox', { name: 'Password' }).fill('Welcome@123');
  await page.getByRole('button', { name: 'Sign in' }).click();
  
  await expect(page).toHaveURL('https://stage10.phenomenex.com/', { waitUntil: 'load', timeout: 200_000});
  await page.getByRole('textbox', { name: 'Search by Part No., Product,' }).click();
  await page.locator('textarea').fill('00F-4496-A0');
  await page.keyboard.press('Enter');
  await page.getByRole('button', { name: 'Add To Cart' }).nth(0).click();
  await page.getByRole('button', { name: 'Continue Shopping' }).click();
  await page.getByRole('textbox', { name: 'Search by Part No., Product,' }).click();

  //await page.locator('#holder').getByText('close').nth(0).click();
  await page.locator('//*[@id="holder"]//app-header-search-modal//span[2]/i').click();
  await page.getByRole('textbox', { name: 'Search by Part No., Product,' }).nth(0).click();

  await page.locator('textarea').nth(0).fill('00B-4446-B0');
  await page.keyboard.press('Enter');
  await page.getByRole('button', { name: 'Add To Cart' }).nth(0).click();
  await page.getByRole('button', { name: 'Continue Shopping' }).click();
  await page.getByRole('link', { name: 'Cart shopping_cart' }).click();
  await page.waitForURL(/cart\.html/, { waitUntil: 'domcontentloaded' });

  await page.getByRole('button', { name: 'Checkout' }).click();
  await page.waitForURL(/addresses\.html/, { waitUntil: 'domcontentloaded' });

  await page.getByRole('button', { name: 'Proceed to Shipping Method' }).click();
  await page.waitForURL(/shipping\.html/, { waitUntil: 'domcontentloaded' });

  //await page.goto('https://stage-shop.phenomenex.com/au/en/shipping.html');
  await page.getByRole('button', { name: 'Proceed to Payment' }).click();
  await page.waitForURL(/payment\.html/, { waitUntil: 'domcontentloaded' });

  await page.evaluate(() => { window.scrollBy(0, 500);});
  await page.getByText('Use Card').nth(0).click();
  //await page.getByRole('button', { name: 'Use Card' }).nth(0).click();
  await page.evaluate(() => { window.scrollBy(0, 700);});
  await page.getByRole('checkbox').scrollIntoViewIfNeeded();

  await page.locator('(//input[@id="accept-term"])[2]').check();
  await page.getByRole('button', { name: 'Place your order' }).click();
  await page.waitForURL(/receipt\.html/, { waitUntil: 'domcontentloaded' });
  await expect(page).toHaveURL(/^https:\/\/stage-shop\.phenomenex\.com\/us\/en\/receipt\.html/);
  await expect(page.locator('text=/Order Confirmed/i')).toHaveText(/Order Confirmed/i);
  
});