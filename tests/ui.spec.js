const { test, expect } = require('@playwright/test');

test.describe('Office Add-in UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Mock Office.js
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
      await route.fulfill({ body: '// Mock Office.js' });
    });

    // Inject Office mock object
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          setTimeout(() => {
            callback({ host: null }); // Simulate browser environment
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
                    size: 100,
                    sliceCount: 2,
                    getSliceAsync: (index, sliceCallback) => {
                      setTimeout(() => {
                        sliceCallback({
                          status: 'succeeded',
                          value: { data: new Uint8Array(50) }
                        });
                      }, 500); // Increased delay to ensure Playwright can catch intermediate states
                    },
                    closeAsync: (closeCallback) => {
                      setTimeout(() => closeCallback(), 50);
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

  test('should initialize with correct initials', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should have the correctly pluralized button text', async ({ page }) => {
    const button = page.locator('#read-files-btn');
    await expect(button).toHaveText('Read Active File(s)');
  });

  test('should read active files and update status', async ({ page }) => {
    // Wait for initialization to ensure the button click handler is attached
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const button = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await button.click();

    // Check intermediate status
    await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);

    // Check progress (at least the first one to ensure it's working)
    await expect(status).toHaveText(/Reading progress: 50%/, { timeout: 5000 });

    // Check final status
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
  });
});
