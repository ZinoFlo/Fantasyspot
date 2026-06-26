const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Taskpane', () => {
  test.beforeEach(async ({ page }) => {
    // Mock Office.js
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
      await route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js',
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
                          value: { data: new Uint8Array(50) },
                        });
                      }, 500); // 500ms delay to capture intermediate state
                    },
                    closeAsync: (closeCallback) => {
                      closeCallback();
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

  test('should read active files and update status', async ({ page }) => {
    // Wait for initialization
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await readBtn.click();

    // Verify intermediate "Reading" state
    await expect(status).toHaveText(/Reading file\(s\).../);

    // Verify progress update (from getSliceAsync delay)
    await expect(status).toHaveText(/Reading progress: 50%/);

    // Verify final success state
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes./);
  });
});
