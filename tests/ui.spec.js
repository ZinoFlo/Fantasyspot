const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Add-in UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept and mock office.js file retrieval to avoid loading external JS and throwing errors
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', (route) => {
      route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js library script loading',
      });
    });

    // Injected Office environment before scripts run
    await page.addInitScript(() => {
      // Create global Office structure
      window.Office = {
        HostType: {
          PowerPoint: 'PowerPoint',
        },
        FileType: {
          Compressed: 'Compressed',
        },
        AsyncResultStatus: {
          Succeeded: 'Succeeded',
          Failed: 'Failed',
        },
        onReady: (callback) => {
          // Add small delay to mimic asynchronous Office loading
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
          }, 100);
        },
        context: {
          document: {
            getFileAsync: (fileType, options, callback) => {
              // Simulate file handling
              setTimeout(() => {
                callback({
                  status: 'Succeeded',
                  value: {
                    size: 131072, // 128 KB
                    sliceCount: 2,
                    closeAsync: (closeCallback) => {
                      setTimeout(closeCallback, 50);
                    },
                    getSliceAsync: (sliceIndex, sliceCallback) => {
                      setTimeout(() => {
                        sliceCallback({
                          status: 'Succeeded',
                          value: {
                            data: new Uint8Array(65536), // 64 KB per slice
                          },
                        });
                      }, 500); // 500ms delay to make sure UI intermediate states are captured
                    },
                  },
                });
              }, 100);
            },
          },
        },
      };
    });

    await page.goto('/');
  });

  test('should display initials JV after initialization', async ({ page }) => {
    // Wait for the initialization to transition the display text to 'JV'
    await expect(page.locator('#initials-display')).toHaveText('JV');
  });

  test('should successfully read active file(s) with progress updates', async ({ page }) => {
    // Ensure initialized
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    // Click the Read Active File(s) button
    await readBtn.click();

    // Check initial reading status
    await expect(status).toHaveText('Reading active file(s)...');

    // Check progress and success sequences
    await expect(status).toHaveText(/Active file size: 131072 bytes\. Reading 2 slices\.\.\./);
    await expect(status).toHaveText(/(Reading progress: (50|100)%|Successfully read active file\(s\): 131072 bytes\.)/);
    await expect(status).toHaveText(/Successfully read active file\(s\): 131072 bytes\./);
  });
});
