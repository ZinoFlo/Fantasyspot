const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Add-in', () => {
  test.beforeEach(async ({ page }) => {
    // Mock Office.js
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
                      // Increase delay to ensure Playwright can capture intermediate states
                      setTimeout(() => {
                        sliceCallback({
                          status: 'Succeeded',
                          value: { data: new Array(50).fill(0) }
                        });
                      }, 200);
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

    await page.goto('/index.html');
  });

  test('should initialize with correct initials', async ({ page }) => {
    const initialsDisplay = page.locator('#initials-display');
    // Wait for the initials to be updated from '--' to 'JV'
    await expect(initialsDisplay).toHaveText('JV');
  });

  test('should read active file(s) successfully', async ({ page }) => {
    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');
    const initialsDisplay = page.locator('#initials-display');

    // Wait for initialization to complete before interacting
    await expect(initialsDisplay).toHaveText('JV');

    await expect(readBtn).toHaveText('Read Active File(s)');

    // Initial status should be empty
    await expect(status).toBeEmpty();

    await readBtn.click();

    // Verify sequence of status messages.
    // Note: We use regex or partial matches where appropriate to handle fast transitions.
    await expect(status).toHaveText('Reading active file(s)...');
    await expect(status).toHaveText(/File size: 100 bytes/);
    await expect(status).toHaveText('Reading progress: 50%');

    // The "100%" message might be very brief before the success message.
    // We check for the final success message as the terminal state.
    await expect(status).toHaveText('Successfully read active file(s): 100 bytes.');
  });
});
