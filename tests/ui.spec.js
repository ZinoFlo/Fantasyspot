const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept Office.js script request and fulfill with a mock
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async route => {
      await route.fulfill({ body: '// Mock Office.js' });
    });

    // Mock Office object before the application scripts run
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
                    size: 100,
                    sliceCount: 2,
                    getSliceAsync: (index, sliceCallback) => {
                      setTimeout(() => {
                        sliceCallback({
                          status: 'Succeeded',
                          value: { data: new Array(50).fill(0) }
                        });
                      }, 500); // 500ms delay to capture intermediate status
                    },
                    closeAsync: (closeCallback) => {
                      setTimeout(() => {
                        closeCallback({ status: 'Succeeded' });
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

  test('should initialize with correct initials', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should read active files and show progress', async ({ page }) => {
    // Wait for initialization
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    await readBtn.click();

    const status = page.locator('#status');

    // Check for intermediate state
    await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);

    // Check for progress (mocked with 2 slices, 50% is a reliable intermediate state)
    await expect(status).toHaveText(/Reading progress: 50%/, { timeout: 5000 });

    // The "100%" state might be too transient to reliably capture as it is immediately
    // followed by the success message. We verify the final state.

    // Check for success message
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
  });
});
