const { test, expect } = require('@playwright/test');

test.describe('Office Add-in UI', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept Office.js script to prevent network requests and use our mock
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
      await route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js',
      });
    });

    // Mock Office object and constants before the page scripts run
    await page.addInitScript(() => {
      window.Office = {
        HostType: { PowerPoint: 'PowerPoint' },
        FileType: { Compressed: 'Compressed' },
        AsyncResultStatus: { Succeeded: 'Succeeded', Failed: 'Failed' },
        onReady: (callback) => {
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
          }, 100);
        },
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
                          value: { data: new Uint8Array(50) },
                        });
                      }, 50);
                    },
                    closeAsync: (closeCallback) => {
                      if (closeCallback) closeCallback();
                    },
                  },
                });
              }, 100);
            },
          },
        },
      };
    });

    await page.goto('/index.html');
  });

  test('should initialize with correct initials', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should read active files successfully', async ({ page }) => {
    // Ensure initials are loaded, indicating Office.onReady has fired
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    await readBtn.click();

    const status = page.locator('#status');
    await expect(status).toHaveText('Reading active file(s)...');

    // Check for intermediate progress state
    await expect(status).toHaveText(/Reading progress: (50|100)%/);

    // Final success message
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
  });
});
