const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Add-in', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept Office.js script to prevent external network calls
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', route => {
      route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js'
      });
    });

    // Inject Mock Office Object
    await page.addInitScript(() => {
      window.Office = {
        HostType: { PowerPoint: 'PowerPoint' },
        FileType: { Compressed: 'Compressed' },
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        onReady: (callback) => {
          // Trigger callback with a slight delay to simulate async initialization
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
          }, 100);
        },
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
                      // Use a longer delay to ensure UI states are capturable
                      setTimeout(() => {
                        sliceCallback({
                          status: 'succeeded',
                          value: { data: new Uint8Array(50) }
                        });
                      }, 500);
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

    await page.goto('http://localhost:3000/index.html');
  });

  test('should initialize with correct initials', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should read active files successfully', async ({ page }) => {
    // Ensure initialized
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    await readBtn.click();

    const status = page.locator('#status');

    // Check for intermediate status
    await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);

    // Check for file size and slice count message
    await expect(status).toHaveText(/File size: 100 bytes\. Reading 2 slice\(s\)\.\.\./, { timeout: 10000 });

    // Check for progress (50% is a reliable intermediate state)
    await expect(status).toHaveText(/Reading progress: 50%/, { timeout: 10000 });

    // The transition from 100% to Success might be very fast.
    // We primarily care that it reaches the final success state.
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./, { timeout: 10000 });
  });
});
