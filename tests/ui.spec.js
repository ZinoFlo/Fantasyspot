const { test, expect } = require('@playwright/test');

test.describe('Office Add-in UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Mock Office.js
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async route => {
      await route.fulfill({ body: '// Mock Office.js' });
    });

    // Inject Office mock before scripts run
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
                      }, 500); // Delay to capture intermediate status
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

    await page.goto('http://localhost:3000/index.html');
  });

  test('should initialize with initials JV', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should read active files and show progress', async ({ page }) => {
    // Ensure initialized
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    await readBtn.click();

    const status = page.locator('#status');

    // Check initial status
    await expect(status).toHaveText('Reading active file(s)...');

    // Check progress
    await expect(status).toHaveText('Reading progress: 50%');

    // Check final status
    await expect(status).toHaveText('Successfully read active file(s): 100 bytes.');
  });
});
