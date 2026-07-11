const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Taskpane', () => {
  test.beforeEach(async ({ page }) => {
    // Mock Office.js library
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', route => {
      route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js'
      });
    });

    // Inject Office mock before scripts load
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

    await page.goto('/');
  });

  test('should initialize with correct initials', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should read active file(s) and update status', async ({ page }) => {
    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await expect(readBtn).toBeVisible();
    await readBtn.click();

    await expect(status).toHaveText('Reading active file(s)...');

    // Check for intermediate progress message
    // Note: Due to fast processing, we allow skipping directly to the success message
    // but we expect at least one of these states to be reachable or the terminal state.
    await expect(status).toHaveText(/(Reading progress: (50|100)%|Successfully read active file\(s\): 100 bytes total.)/);

    // Check for final success message
    await expect(status).toHaveText('Successfully read active file(s): 100 bytes total.');
  });
});
