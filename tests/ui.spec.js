const { test, expect } = require('@playwright/test');

test.describe('Office Add-in UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept office.js script to prevent external network calls
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', (route) => {
      route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js',
      });
    });

    // Mock the Office environment before the page loads
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          // Trigger the callback with a small delay to simulate real loading
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
          }, 100);
        },
        HostType: { PowerPoint: 'PowerPoint' },
        FileType: { Compressed: 'Compressed' },
        AsyncResultStatus: { Succeeded: 'Succeeded', Failed: 'Failed' },
        context: {
          document: {
            getFileAsync: (fileType, options, callback) => {
              // Simulate getFileAsync success
              setTimeout(() => {
                callback({
                  status: 'Succeeded',
                  value: {
                    size: 1000,
                    sliceCount: 2,
                    getSliceAsync: (index, sliceCallback) => {
                      setTimeout(() => {
                        sliceCallback({
                          status: 'Succeeded',
                          value: {
                            data: new Array(500).fill(0),
                          },
                        });
                      }, 100);
                    },
                    closeAsync: (closeCallback) => {
                      setTimeout(() => {
                        closeCallback();
                      }, 50);
                    },
                  },
                });
              }, 100);
            },
          },
        },
      };
    });

    await page.goto('http://localhost:3000/index.html');
  });

  test('should display initials after initialization', async ({ page }) => {
    const initialsDisplay = page.locator('#initials-display');
    // Wait for the mock onReady to trigger and update the UI
    await expect(initialsDisplay).toHaveText('JV', { timeout: 5000 });
  });

  test('should read active files and update status', async ({ page }) => {
    // Ensure initialized
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await readBtn.click();

    // Check intermediate "Reading" status
    await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);

    // Check for progress updates (asserting on 50% as a reliable intermediate state)
    await expect(status).toHaveText(/Reading progress: 50%/, { timeout: 5000 });

    // Check final success status
    await expect(status).toHaveText(/Successfully read active file\(s\): 1000 bytes\./, { timeout: 5000 });
  });
});
