const { chromium } = require('playwright');
const path = require('path');
const { execSync } = require('child_process');
const fs = require('fs');
const notifier = require('node-notifier');

// Load configuration
const configPath = path.join(__dirname, 'config.json');
const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));

// Determine output directory
const outputDir = config.outputFolder && config.outputFolder.trim() !== "" 
    ? path.resolve(config.outputFolder) 
    : process.cwd();

// Ensure output directory exists
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
}

function notifyFailure(errorMessage) {
    console.log('\n🚨 TRIPLE ALERT: NOTIFYING USER OF FAILURE 🚨');
    
    // 1. Sound Alert (System Beep)
    process.stdout.write('\x07'); 

    // 2. System Notification (Toast)
    notifier.notify({
        title: '❌ OM Automation FAILED',
        message: `Failed after 5 attempts. Last error: ${errorMessage}`,
        sound: true,
        wait: true
    });

    // 3. Desktop Popup (Stays until closed)
    try {
        const script = `Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('The OM Automation script has failed after 5 attempts.\\n\\nLast error: ${errorMessage.replace(/'/g, "''")}', 'Automation Error', 'OK', 'Error')`;
        execSync(`powershell -Command "${script}"`);
    } catch (e) {
        console.log('Could not show popup, check console for errors.');
    }
}

async function runAutomation() {
    // Only kill processes if NOT in headless mode (to avoid killing other Edge instances)
    if (!config.headless) {
        try {
            execSync('taskkill /F /IM msedge.exe /T', { stdio: 'ignore' });
        } catch (e) {}
    }

    const browser = await chromium.launch({
        channel: 'msedge',
        headless: config.headless,
        slowMo: config.headless ? 0 : 100 // No need to slow down in headless mode
    });

    const context = await browser.newContext({
        viewport: { width: 1280, height: 720 },
        // Use a real User Agent to avoid detection in headless mode
        userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'
    });

    const page = await context.newPage();

    try {
        console.log(`Navigating to the report page (Headless: ${config.headless})...`);
        await page.goto(config.reportUrl);

        console.log('Waiting for login popup...');
        const loginPopup = await page.waitForEvent('popup', { timeout: 45000 });
        await loginPopup.waitForLoadState();

        console.log('Performing login...');
        await loginPopup.getByRole('textbox', { name: 'Enter your email or phone' }).fill(config.email);
        await loginPopup.getByRole('button', { name: 'Next' }).click();
        
        await loginPopup.getByRole('textbox', { name: 'Enter the password for' }).waitFor({ state: 'visible', timeout: 30000 });
        await loginPopup.getByRole('textbox', { name: 'Enter the password for' }).fill(config.password);
        await loginPopup.getByRole('textbox', { name: 'Enter the password for' }).press('Enter');

        try {
            await loginPopup.getByRole('button', { name: 'Yes' }).click({ timeout: 5000 });
        } catch (e) {}

        await loginPopup.waitForEvent('close', { timeout: 15000 }).catch(() => {});

        console.log('Refreshing page once...');
        await page.reload({ waitUntil: 'load' });
        await page.waitForTimeout(2000);

        console.log('Re-inserting report link into URL bar...');
        await page.goto(config.reportUrl, { waitUntil: 'load', timeout: 60000 });

        console.log('Waiting for page elements...');
        await page.waitForSelector('.mat-select-placeholder', { timeout: 60000 });

        console.log('Selecting report parameters...');
        await page.locator('.mat-select-placeholder').first().click();
        await page.locator('.mat-checkbox-inner-container').first().click();
        
        console.log('Pressing Escape...');
        await page.keyboard.press('Escape');
        await page.waitForTimeout(1000);

        console.log('Selecting second set of parameters...');
        await page.locator('.mat-select-placeholder').click();
        await page.locator('.mat-checkbox-inner-container').first().click();
        await page.locator('.cdk-overlay-backdrop').click();

        console.log('Exporting to Excel...');
        const downloadPromise = page.waitForEvent('download');
        await page.getByRole('button', { name: 'Export To Excel' }).click();
        
        const download = await downloadPromise;
        const downloadPath = path.join(outputDir, download.suggestedFilename());
        await download.saveAs(downloadPath);
        
        console.log(`✅ Success! File saved to: ${downloadPath}`);
        await page.getByRole('button', { name: 'OK' }).click({ timeout: 2000 }).catch(() => {});
        
        return true; 

    } catch (error) {
        console.error('❌ Attempt failed:', error.message);
        if (!config.headless) {
            await page.screenshot({ path: path.join(outputDir, 'debug_error.png') });
        }
        throw error;
    } finally {
        await context.close();
        await browser.close();
    }
}

(async () => {
    const MAX_RETRIES = 5;
    let success = false;
    let lastError = "";

    for (let i = 1; i <= MAX_RETRIES; i++) {
        console.log(`\n--- Starting Attempt ${i} of ${MAX_RETRIES} ---`);
        try {
            success = await runAutomation();
            if (success) break;
        } catch (e) {
            lastError = e.message;
            if (i < MAX_RETRIES) {
                console.log(`Waiting 10 seconds before retry...`);
                await new Promise(r => setTimeout(r, 10000));
            }
        }
    }

    if (!success) {
        notifyFailure(lastError);
        process.exit(1);
    } else {
        console.log('\n✨ Automation completed successfully.');
        process.exit(0);
    }
})();
