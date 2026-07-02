const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Add-in', () => {
  test.beforeEach(async ({ page }) => {
    // Mock Office.js environment
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          setTimeout(() => {
            callback({ host: null }); // Simulate browser/standalone host
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
                          value: { data: new Array(50).fill(0) }
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

    // Mock the external Office.js script to prevent network requests
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', route => {
      route.fulfill({ body: '// Mock Office.js' });
    });

    await page.goto('/');
  });

  test('should initialize with initials JV', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should read active file(s) successfully', async ({ page }) => {
    // Ensure initialized
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    await readBtn.click();

    const status = page.locator('#status');

    // Check intermediate progress
    await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);
    await expect(status).toHaveText(/File size: 100 bytes\. Reading 2 slice\(s\)\.\.\./);

    // Check terminal success message
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
  });
});
