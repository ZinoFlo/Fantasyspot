const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Mock Office.js before the page loads
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
                    size: 200,
                    sliceCount: 2,
                    getSliceAsync: (index, sliceCallback) => {
                      setTimeout(() => {
                        sliceCallback({
                          status: 'succeeded',
                          value: { data: new Array(100).fill(0) }
                        });
                      }, 500);
                    },
                    closeAsync: (closeCallback) => {
                      closeCallback();
                    }
                  }
                });
              }, 200);
            }
          }
        }
      };
    });

    // Intercept the real office.js script to prevent network requests
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', route => {
      route.fulfill({ body: '// Mocked Office.js' });
    });

    await page.goto('/');
  });

  test('should display initials after Office is ready', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should read active files and update status', async ({ page }) => {
    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await expect(page.locator('#initials-display')).toHaveText('JV');

    await readBtn.click();

    // Check intermediate status
    await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);

    // Check progress (mocked with 2 slices)
    await expect(status).toHaveText(/Reading progress: 50%/, { timeout: 5000 });

    // Final status
    await expect(status).toHaveText(/Successfully read active file\(s\): 200 bytes\./);
  });
});
