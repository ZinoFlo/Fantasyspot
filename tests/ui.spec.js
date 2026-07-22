const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Taskpane UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept external Office.js script to prevent network calls and let our mock take effect.
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', (route) => {
      route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mocked Office.js loaded',
      });
    });

    // Add init script to mock the Office namespace and its API before index.js runs.
    await page.addInitScript(() => {
      window.Office = {
        HostType: { PowerPoint: 'PowerPoint' },
        FileType: { Compressed: 'Compressed' },
        AsyncResultStatus: {
          Succeeded: 'Succeeded',
          Failed: 'Failed',
        },
        context: {
          document: {
            getFileAsync(fileType, options, callback) {
              // Simulate async nature of PowerPoint Office.js
              setTimeout(() => {
                callback({
                  status: window.Office.AsyncResultStatus.Succeeded,
                  value: {
                    sliceCount: 2,
                    size: 100,
                    getSliceAsync(sliceIndex, sliceCallback) {
                      setTimeout(() => {
                        sliceCallback({
                          status: window.Office.AsyncResultStatus.Succeeded,
                          value: {
                            // Slice size/data mock
                            data: new Uint8Array(50),
                          },
                        });
                      }, 500); // Wait 500ms to allow UI progress assertions to catch the transition states
                    },
                    closeAsync(closeCallback) {
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
        onReady(callback) {
          // Trigger callbacks after DOM scripts execute, simulating native behavior
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
          }, 100);
        },
      };
    });

    await page.goto('/');
  });

  test('should display co-op initials as "JV" upon Office onReady', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should read presentation files and display reading progress and success messages', async ({ page }) => {
    // Wait for initials-display to update to "JV" (signals Office is ready and handlers attached)
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await expect(status).toBeEmpty();

    // Trigger read action
    await readBtn.click();

    // Check intermediate reading states (due to the 500ms delay per slice, we can catch progress)
    await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);

    // Check intermediate size message
    await expect(status).toHaveText(/Active file\(s\) size: 100 bytes\. Reading 2 slices\.\.\./);

    // Wait and verify first slice progress (50%)
    await expect(status).toHaveText(/Reading progress: 50%/);

    // Wait and verify complete progress / success message
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
  });
});
