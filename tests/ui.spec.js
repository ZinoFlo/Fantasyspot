const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Mock Office.js
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async route => {
      await route.fulfill({ body: '// Mock Office.js' });
    });

    // Inject Office mock
    await page.addInitScript(() => {
      window.Office = {
        onReady: (cb) => {
          setTimeout(() => {
            cb({ host: 'PowerPoint' });
          }, 100);
        },
        HostType: { PowerPoint: 'PowerPoint' },
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        FileType: { Compressed: 'compressed' },
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
                          value: { data: new Array(50).fill(0) }
                        });
                      }, 50);
                    },
                    closeAsync: (closeCallback) => {
                      closeCallback();
                    }
                  }
                });
              }, 50);
            }
          }
        }
      };
    });

    await page.goto('/index.html');
  });

  test('should display initial state correctly', async ({ page }) => {
    await expect(page.locator('h1')).toHaveText('Eco-growth Discovery');
    await expect(page.locator('#initials-display')).toHaveText('JV');
    await expect(page.locator('#read-files-btn')).toHaveText('Read Active Files');
  });

  test('should read active files and show progress', async ({ page }) => {
    const status = page.locator('#status');
    const readBtn = page.locator('#read-files-btn');

    // Wait for initialization
    await expect(page.locator('#initials-display')).toHaveText('JV');

    await readBtn.click();

    // Verify transition of status messages
    await expect(status).toHaveText('Reading active file(s)...');
    await expect(status).toHaveText('Successfully read active file(s): 100 bytes.');
  });
});
