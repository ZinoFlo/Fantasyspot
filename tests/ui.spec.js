const { test, expect } = require('@playwright/test');

test.beforeEach(async ({ page }) => {
  // Mock Office.js
  await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', route => {
    route.fulfill({
      status: 200,
      contentType: 'application/javascript',
      body: '// Mock Office.js',
    });
  });

  // Inject Office mock before scripts run
  await page.addInitScript(() => {
    window.Office = {
      HostType: { PowerPoint: 'PowerPoint' },
      FileType: { Compressed: 'Compressed' },
      AsyncResultStatus: { Succeeded: 'Succeeded', Failed: 'Failed' },
      onReady: (callback) => {
        // Simulate async initialization
        setTimeout(() => {
          callback({ host: 'PowerPoint' });
        }, 100);
      },
      context: {
        document: {
          getFileAsync: (fileType, options, callback) => {
            setTimeout(() => {
              callback({
                status: 'Succeeded',
                value: {
                  size: 100,
                  sliceCount: 2,
                  getSliceAsync: (index, cb) => {
                    setTimeout(() => {
                      cb({
                        status: 'Succeeded',
                        value: { data: new Uint8Array(50) }
                      });
                    }, 50);
                  },
                  closeAsync: (cb) => {
                    setTimeout(() => { if (cb) cb(); }, 10);
                  }
                }
              });
            }, 100);
          }
        }
      }
    };
  });
});

test('should initialize and show initials', async ({ page }) => {
  await page.goto('/index.html');
  const initials = page.locator('#initials-display');
  await expect(initials).toHaveText('JV');
});

test('should have pluralized button text', async ({ page }) => {
  await page.goto('/index.html');
  const button = page.locator('#read-files-btn');
  await expect(button).toHaveText('Read Active File(s)');
});

test('should read file(s) and show success status', async ({ page }) => {
  await page.goto('/index.html');

  // Wait for initialization
  await expect(page.locator('#initials-display')).toHaveText('JV');

  const button = page.locator('#read-files-btn');
  await button.click();

  const status = page.locator('#status');

  // Check intermediate state
  await expect(status).toHaveText(/Reading active file\(s\)\.\.\.|File size: 100 bytes/);

  // Check final success state
  await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
});
