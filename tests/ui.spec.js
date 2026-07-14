const { test, expect } = require('@playwright/test');

test.describe('Office Add-in UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept Office.js script to prevent external network dependency
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', (route) => {
      route.fulfill({ body: '// Mock Office.js' });
    });

    // Inject Office mock before the application loads
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
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
                      }, 100);
                    },
                    closeAsync: (closeCallback) => {
                      if (closeCallback) closeCallback();
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

  test('should display "JV" after initialization', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should show correct status messages when reading file(s)', async ({ page }) => {
    // Wait for initialization
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await readBtn.click();

    // Verify sequence of status messages
    await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);
    await expect(status).toHaveText(/Total size: 100 bytes\. Reading 2 slices\.\.\./);
    await expect(status).toHaveText(/Reading progress: (50|100)%/);
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
  });
});
