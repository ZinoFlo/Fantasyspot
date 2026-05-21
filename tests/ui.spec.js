const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery UI', () => {
  test.beforeEach(async ({ page }) => {
    // Mock Office.js
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
      await route.fulfill({ body: '// Mock Office.js' });
    });

    // Inject Office mock before scripts run
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          setTimeout(() => {
            callback({ host: null });
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
                    size: 100,
                    sliceCount: 2,
                    getSliceAsync: (index, sliceCallback) => {
                      setTimeout(() => {
                        sliceCallback({
                          status: 'Succeeded',
                          value: { data: new Uint8Array(50) }
                        });
                      }, 1000);
                    },
                    closeAsync: (closeCallback) => {
                      closeCallback();
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

  test('should display initial state correctly', async ({ page }) => {
    await expect(page.locator('h1')).toHaveText('Eco-growth Discovery');
    // Wait for initials to be updated by onReady mock
    await expect(page.locator('#initials-display')).toHaveText('JV');
    await expect(page.locator('#read-files-btn')).toHaveText('Read Active Files');
  });

  test('should handle file reading process', async ({ page }) => {
    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    // Wait for the button to be ready (initials are displayed)
    await expect(page.locator('#initials-display')).toHaveText('JV');

    await readBtn.click();

    // Check progress (mock has 2 slices)
    // We check for the first progress update
    await expect(status).toHaveText(/Reading progress: 50%/, { timeout: 5000 });

    // We skip the 100% check as it may transition to the final success message
    // too quickly for the Playwright poller to capture.

    // Final success message
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
  });
});
