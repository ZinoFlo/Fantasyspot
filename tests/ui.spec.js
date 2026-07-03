const { test, expect } = require('@playwright/test');

test.describe('Office Add-in UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Mock Office.js
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', route => {
      route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js'
      });
    });

    // Inject Office mock object
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
                    getSliceAsync: (sliceIndex, sliceCallback) => {
                      setTimeout(() => {
                        sliceCallback({
                          status: 'succeeded',
                          value: { data: new Array(50).fill(0) }
                        });
                      }, 500);
                    },
                    closeAsync: (closeCallback) => {
                      if (closeCallback) closeCallback();
                    }
                  }
                });
              }, 50);
            }
          }
        }
      };
    });

    await page.goto('http://localhost:3000/index.html');
  });

  test('should display initials after initialization', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should update status message during file reading', async ({ page }) => {
    // Wait for initialization
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await readBtn.click();

    // Verify progress (using 50% as a reliable intermediate state)
    await expect(status).toHaveText(/Reading progress: 50%/);

    // Verify success message
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
  });
});
