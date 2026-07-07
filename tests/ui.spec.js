const { test, expect } = require('@playwright/test');

test.describe('Office Add-in UI', () => {
  test.beforeEach(async ({ page }) => {
    // Mock Office.js
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
      await route.fulfill({
        body: '// Mock Office.js',
        contentType: 'application/javascript',
      });
    });

    // Injected Office mock
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          setTimeout(() => {
            callback({ host: null }); // Mocking browser environment
          }, 100);
        },
        HostType: { PowerPoint: 'PowerPoint' },
        FileType: { Compressed: 'compressed' },
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
                          value: {
                            data: new Uint8Array(50),
                          },
                        });
                      }, 500); // Increased delay to ensure UI updates are caught
                    },
                    closeAsync: (closeCallback) => {
                      setTimeout(() => {
                        closeCallback();
                      }, 10);
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

  test('should display initial state and then "JV" after initialization', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should read file(s) and update status', async ({ page }) => {
    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await expect(page.locator('#initials-display')).toHaveText('JV');

    await readBtn.click();

    await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);
    await expect(status).toHaveText(/File size: 100 bytes\. Reading 2 slices\.\.\./);
    await expect(status).toHaveText(/Reading progress: 50%/);
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
  });
});
