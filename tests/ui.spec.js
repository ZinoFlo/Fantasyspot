const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Taskpane', () => {
  test.beforeEach(async ({ page }) => {
    // Mock Office.js before navigating
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          // Delaying slightly to simulate initialization
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
                      }, 500);
                    },
                    closeAsync: (closeCallback) => {
                      setTimeout(() => {
                        closeCallback();
                      }, 50);
                    }
                  }
                });
              }, 100);
            }
          }
        }
      };
    });

    // Intercept the external office.js script to prevent network requests
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', route => {
      route.fulfill({ body: '// Mock Office.js' });
    });

    await page.goto('/index.html');
  });

  test('should initialize and display JV initials', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should update status when reading files', async ({ page }) => {
    // Ensure initialized
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await readBtn.click();

    // Verify transition to "Reading file(s)..."
    await expect(status).toHaveText('Reading file(s)...');

    // Verify intermediate progress (mocked slices)
    await expect(status).toHaveText(/Reading progress: 50%/, { timeout: 2000 });

    // Verify final success message
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
  });
});
