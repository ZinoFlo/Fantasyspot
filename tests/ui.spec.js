const { test, expect } = require('@playwright/test');

test.describe('Office Add-in UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // 1. Intercept external Office.js request to prevent network dependencies
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
      // Fulfill with a comment-only string so the global Office object
      // injected via addInitScript is used as the primary implementation.
      await route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js script loaded'
      });
    });

    // 2. Inject mock Office object before page loads
    await page.addInitScript(() => {
      window.Office = {
        HostType: {
          PowerPoint: 'PowerPoint'
        },
        FileType: {
          Compressed: 'Compressed'
        },
        AsyncResultStatus: {
          Succeeded: 'Succeeded',
          Failed: 'Failed'
        },
        onReady: (callback) => {
          // Add slight delay to simulate async Office initialization
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
          }, 100);
        },
        context: {
          document: {
            getFileAsync: (fileType, options, callback) => {
              // Simulate asynchronous behavior with 500ms delay to capture intermediate progress states
              setTimeout(() => {
                callback({
                  status: 'Succeeded', // Succeeded status matches AsyncResultStatus.Succeeded
                  value: {
                    size: 131072, // 128 KB
                    sliceCount: 2,
                    getSliceAsync: (sliceIndex, sliceCallback) => {
                      setTimeout(() => {
                        sliceCallback({
                          status: 'Succeeded',
                          value: {
                            data: new Uint8Array(65536) // 64 KB per slice
                          }
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
              }, 500);
            }
          }
        }
      };
    });

    // 3. Navigate to index.html
    await page.goto('/');
  });

  test('should display initials, transition to "JV", and read active file(s) with progress updates', async ({ page }) => {
    // Check initial state
    const initialsLocator = page.locator('#initials-display');
    const statusLocator = page.locator('#status');
    const readBtn = page.locator('#read-files-btn');

    // Initially initials are '--' or 'JV' depending on speed, but they must eventually transition to 'JV'
    await expect(initialsLocator).toHaveText('JV');

    // Click on Read Active File(s) button
    await readBtn.click();

    // Verify initial "Reading active file(s)..." message is displayed
    await expect(statusLocator).toHaveText('Reading active file(s)...');

    // Verify intermediate progress or final success message using a regex assertion
    // to handle rapid state transitions robustly.
    await expect(statusLocator).toHaveText(/(File size: 131072 bytes\. Reading 2 slices\.\.\.|Reading progress: (50|100)%|Successfully read active file\(s\): 131072 bytes\.)/);

    // Eventually, it should display the final success message
    await expect(statusLocator).toHaveText('Successfully read active file(s): 131072 bytes.', { timeout: 5000 });
  });
});
