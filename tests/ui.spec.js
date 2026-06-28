const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Add-in UI', () => {
  test.beforeEach(async ({ page }) => {
    // Mock Office.js before the page loads
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          // Simulate Office.js being ready with PowerPoint host
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
              // Simulate successful file retrieval
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

    // Intercept the external Office.js script to prevent network requests and errors
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', (route) => {
      route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js'
      });
    });

    await page.goto('/index.html');
  });

  test('should initialize and display initials', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should read active file(s) and update status', async ({ page }) => {
    // Wait for initialization
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    await expect(readBtn).toHaveText('Read Active File(s)');

    await readBtn.click();

    const status = page.locator('#status');

    // Check for intermediate status messages
    await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);
    await expect(status).toHaveText(/File size: 100 bytes\. Reading 2 slices\.\.\./);

    // Check for progress updates
    await expect(status).toHaveText(/Reading progress: 50%/);

    // Check for final success message
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
  });
});
