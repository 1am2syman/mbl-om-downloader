const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

(async () => {
  // Use the temporary Edge profile we created to leverage your session
  const userDataDir = path.join(process.env.LOCALAPPDATA, 'Temp', 'EdgeProxy');
  
  console.log('Launching Edge with profile:', userDataDir);
  
  const context = await chromium.launchPersistentContext(userDataDir, {
    channel: 'msedge',
    headless: true, // Set to true if you want to run in background later
    viewport: { width: 1280, height: 720 },
    slowMo: 100, // Slow down actions to make it visible
  });

  const page = await context.newPage();

  try {
    console.log('Navigating to Order Management page...');
    await page.goto('https://mblnet.maricoapps.biz/components/ordermanagement/order-page-gt-report', {
      waitUntil: 'networkidle',
      timeout: 60000
    });

    // Handle potential login popup
    const popupPromise = page.waitForEvent('popup', { timeout: 10000 }).catch(() => null);
    
    // If a popup appears, handle login
    const loginPopup = await popupPromise;
    if (loginPopup) {
      console.log('Login popup detected, attempting login...');
      await loginPopup.waitForLoadState();
      
      // Use your recorded credentials
      await loginPopup.getByRole('textbox', { name: 'Enter your email or phone' }).fill('syedsazzad.ali@marico.com');
      await loginPopup.getByRole('button', { name: 'Next' }).click();
      
      await loginPopup.getByRole('textbox', { name: 'Enter the password for' }).fill('1amuSyman');
      await loginPopup.getByRole('textbox', { name: 'Enter the password for' }).press('Enter');
      
      // Wait for the popup to close after successful login
      await loginPopup.waitForEvent('close', { timeout: 30000 }).catch(() => console.log('Popup close timeout'));
    }

    console.log('Selecting report parameters...');
    
    // 1st Select
    await page.locator('.mat-select-placeholder').first().click();
    await page.locator('.mat-checkbox-inner-container').first().click(); // Select options
    await page.locator('.cdk-overlay-backdrop').click(); // Close dropdown

    // 2nd Select
    await page.locator('.mat-select-placeholder').click();
    await page.locator('.mat-checkbox-inner-container').first().click(); // Select options
    await page.locator('.cdk-overlay-backdrop').click(); // Close dropdown

    console.log('Starting Export to Excel...');
    
    // Setup download listener before clicking
    const downloadPromise = page.waitForEvent('download');
    
    await page.getByRole('button', { name: 'Export To Excel' }).click();
    
    const download = await downloadPromise;
    const downloadPath = path.join(__dirname, download.suggestedFilename());
    
    await download.saveAs(downloadPath);
    console.log(`✅ Download completed: ${downloadPath}`);

    // Click OK on any post-export dialogs if they exist
    await page.getByRole('button', { name: 'OK' }).click().catch(() => {});

    console.log('Automation finished successfully.');

  } catch (error) {
    console.error('❌ Automation failed:', error);
    // Take a screenshot on failure for debugging
    await page.screenshot({ path: 'error_screenshot.png' });
  } finally {
    // Keep browser open for 5 seconds to see result before closing
    await page.waitForTimeout(5000);
    await context.close();
  }
})();
