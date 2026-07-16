const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Add-in', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept Office.js script and mock it
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', (route) => {
      route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js',
      });
    });

    // Inject Office mock before page loads
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
                          value: { data: new Array(50).fill(0) },
                        });
                      }, 500);
                    },
                    closeAsync: (closeCallback) => {
                      setTimeout(() => {
                        closeCallback();
                      }, 100);
                    },
                  },
                });
              }, 100);
            },
          },
        },
      };
    });

    await page.goto('/');
  });

  test('initializes correctly with JV initials', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('reads active files and updates status', async ({ page }) => {
    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await expect(readBtn).toHaveText('Read Active File(s)');

    // Ensure initialized before clicking
    await expect(page.locator('#initials-display')).toHaveText('JV');

    await readBtn.click();

    // Check intermediate status
    await expect(status).toHaveText(/Reading active file\(s\)\.\.\.|Total byte size: 100\. Reading 2 slice\(s\)\.\.\./);

    // Check progress
    await expect(status).toHaveText(/Reading progress: (50|100)%/, { timeout: 5000 });

    // Check final status
    await expect(status).toHaveText('Successfully read active file(s): 100 bytes.', { timeout: 10000 });
  });
});
