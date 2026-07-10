const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Taskpane', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept Office.js script to prevent external network requests
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', (route) => {
      route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js'
      });
    });

    // Inject Office mock before the page scripts run
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
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
              setTimeout(() => {
                callback({
                  status: 'Succeeded',
                  value: {
                    size: 131072, // 128 KB
                    sliceCount: 2,
                    getSliceAsync: (sliceIndex, sliceCallback) => {
                      setTimeout(() => {
                        sliceCallback({
                          status: 'Succeeded',
                          value: {
                            data: new Uint8Array(65536) // 64 KB
                          }
                        });
                      }, 500); // 500ms delay to ensure progress is visible
                    },
                    closeAsync: (closeCallback) => {
                      setTimeout(() => {
                        closeCallback();
                      }, 100);
                    }
                  }
                });
              }, 100);
            }
          }
        }
      };
    });

    await page.goto('/');
  });

  test('should initialize with initials JV', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should read active files and show progress', async ({ page }) => {
    const status = page.locator('#status');
    const readBtn = page.locator('#read-files-btn');

    // Wait for initialization
    await expect(page.locator('#initials-display')).toHaveText('JV');

    // Start reading
    await readBtn.click();

    // Verify initial status
    await expect(status).toHaveText('Reading active file(s)...');

    // Verify progress - first slice
    await expect(status).toHaveText('Reading progress: 50%');

    // Verify final success message
    // Note: The "100%" progress state may be transitioned too quickly to be reliably captured.
    await expect(status).toHaveText(/Successfully read active file\(s\): 131072 bytes\./);
  });
});
