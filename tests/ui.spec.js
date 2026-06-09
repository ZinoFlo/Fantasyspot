const { test, expect } = require('@playwright/test');

test.describe('Office Add-in UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Mock Office.js before the page loads
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          // Simulate Office.js being ready with a slight delay
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
          }, 100);
        },
        context: {
          document: {
            getFileAsync: (fileType, options, callback) => {
              // Mock getFileAsync response
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
                            data: new Array(50).fill(0)
                          }
                        });
                      }, 500); // 500ms delay to capture intermediate status
                    },
                    closeAsync: (closeCallback) => {
                      setTimeout(closeCallback, 100);
                    }
                  }
                });
              }, 200);
            }
          }
        },
        FileType: { Compressed: 'compressed' },
        AsyncResultStatus: { Succeeded: 'succeeded' },
        HostType: { PowerPoint: 'PowerPoint' }
      };
    });

    // Intercept the external Office.js script request and fulfill it with a mock
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', route => {
      route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js'
      });
    });

    await page.goto('http://localhost:3000/index.html');
  });

  test('should initialize and display initials "JV"', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should read active file and show progress', async ({ page }) => {
    // Wait for initialization
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await readBtn.click();

    // Check for intermediate status
    await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);

    // Check for progress update (due to 500ms delay in mock)
    await expect(status).toHaveText(/Reading progress: 50%/);

    // Check for success message
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
  });
});
