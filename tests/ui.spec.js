const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept Office.js script to prevent external network requests
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
        onReady: (callback) => {
          setTimeout(() => {
            callback({ host: null }); // Simulate browser testing environment
          }, 100);
        },
        HostType: { PowerPoint: 'PowerPoint' },
        FileType: { Compressed: 'Compressed' },
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        context: {
          document: {
            getFileAsync: (fileType, options, callback) => {
              setTimeout(() => {
                callback({
                  status: 'succeeded',
                  value: {
                    size: 1024,
                    sliceCount: 2,
                    getSliceAsync: (index, sliceCallback) => {
                      setTimeout(() => {
                        sliceCallback({
                          status: 'succeeded',
                          value: { data: new Uint8Array(512) }
                        });
                      }, 50);
                    },
                    closeAsync: (closeCallback) => {
                      setTimeout(() => {
                        closeCallback({ status: 'succeeded' });
                      }, 10);
                    }
                  }
                });
              }, 100);
            }
          }
        }
      };
    });

    await page.goto('/index.html');
  });

  test('should display correct initials', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should have pluralized button text', async ({ page }) => {
    const button = page.locator('#read-files-btn');
    await expect(button).toHaveText('Read Active Files');
  });

  test('should update status when reading files', async ({ page }) => {
    // Wait for Office.onReady to complete (initials will be set to 'JV')
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const button = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await button.click();
    await expect(status).toHaveText(/Reading active file\(s\).../);

    // Wait for the success message (handling progress updates)
    await expect(status).toHaveText(/Successfully read active file\(s\): 1024 bytes\./, { timeout: 5000 });
  });

  test('button-container should have correct styling', async ({ page }) => {
    const container = page.locator('.button-container');
    await expect(container).toHaveCSS('margin-top', '30px');
  });
});
