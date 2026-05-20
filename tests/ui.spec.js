const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Add-in UI', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept Office.js script request and mock it
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async route => {
      await route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: 'window.Office = window.Office || {};'
      });
    });

    // Inject Office mock BEFORE the application scripts load
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          setTimeout(() => {
            callback({ host: null }); // Simulate browser environment
          }, 100);
        },
        HostType: {
          PowerPoint: 'PowerPoint'
        },
        FileType: {
          Compressed: 'compressed'
        },
        AsyncResultStatus: {
          Succeeded: 'succeeded',
          Failed: 'failed'
        },
        context: {
          document: {
            getFileAsync: (fileType, options, callback) => {
              console.log('getFileAsync called');
              setTimeout(() => {
                callback({
                  status: 'succeeded',
                  value: {
                    size: 100,
                    sliceCount: 1,
                    getSliceAsync: (index, sliceCallback) => {
                      setTimeout(() => {
                        sliceCallback({
                          status: 'succeeded',
                          value: { data: new Uint8Array(100) }
                        });
                      }, 50);
                    },
                    closeAsync: (closeCallback) => {
                      if (closeCallback) closeCallback();
                    }
                  }
                });
              }, 50);
            }
          }
        }
      };
    });

    await page.goto('http://localhost:3000/index.html');
  });

  test('should display correctly initialized initials', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV', { timeout: 10000 });
  });

  test('should have pluralized button text', async ({ page }) => {
    const button = page.locator('#read-files-btn');
    await expect(button).toHaveText('Read Active Files');
  });

  test('should show success message when reading active files', async ({ page }) => {
    const button = page.locator('#read-files-btn');
    const status = page.locator('#status');

    // Wait for Office.onReady to be called before clicking
    await page.waitForFunction(() => document.getElementById('initials-display').textContent === 'JV');

    await button.click();
    await expect(status).toHaveText(/Reading active file\(s\)\.\.\./, { timeout: 10000 });
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./, { timeout: 10000 });
  });
});
