const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Add-in', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept and mock office.js
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
      await route.fulfill({ body: '// Mock Office.js' });
    });

    // Inject Office mock
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          setTimeout(() => {
            callback({ host: null }); // Simulate browser environment
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
                      }, 100);
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

    await page.goto('/index.html');
  });

  test('should initialize with correct initials', async ({ page }) => {
    const initialsDisplay = page.locator('#initials-display');
    await expect(initialsDisplay).toHaveText('JV');
  });

  test('should read active file(s) and update status', async ({ page }) => {
    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await expect(readBtn).toHaveText('Read Active File(s)');

    // Ensure the app is initialized before clicking
    await expect(page.locator('#initials-display')).toHaveText('JV');

    await readBtn.click();

    // Verify progress sequence
    await expect(status).toHaveText('Reading active file(s)...');
    await expect(status).toHaveText(/File size: 100 bytes. Reading 2 slices.../);
    // Assert on a reliable intermediate state and the final terminal state
    await expect(status).toHaveText('Reading progress: 50%');
    await expect(status).toHaveText('Successfully read active file(s): 100 bytes.');
  });
});
