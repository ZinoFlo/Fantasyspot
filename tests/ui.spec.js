const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Office Add-in UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept Office.js external script request to avoid external network calls and avoid overwriting our mock
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
      await route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js library'
      });
    });

    // Inject a robust mock of the Office environment before scripts execute
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
          // Add a slight delay to allow application scripts to finish loading and event listeners to attach
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
          }, 100);
        },
        context: {
          document: {
            getFileAsync: (fileType, options, callback) => {
              // Mock getFileAsync with custom 500ms delay to allow progress detection in UI
              setTimeout(() => {
                callback({
                  status: 'Succeeded',
                  value: {
                    size: 100000,
                    sliceCount: 2,
                    closeAsync: (cb) => {
                      setTimeout(cb, 100);
                    },
                    getSliceAsync: (sliceIndex, cb) => {
                      setTimeout(() => {
                        cb({
                          status: 'Succeeded',
                          value: {
                            data: new Uint8Array(50000)
                          }
                        });
                      }, 500); // 500ms delay to capture progress sequence
                    }
                  }
                });
              }, 100);
            }
          }
        }
      };
    });

    await page.goto('/');
  });

  test('should display initial state and then load co-op initials JV', async ({ page }) => {
    // Initially initials should be '--'
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should read active file(s) and display correct progress and success messages', async ({ page }) => {
    // Wait for the initials JV to confirm Office is initialized and event listeners are set up
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    await expect(readBtn).toHaveText('Read Active File(s)');

    const status = page.locator('#status');
    await expect(status).toBeEmpty();

    // Click button to read active files
    await readBtn.click();

    // Verify loading state is shown initially
    await expect(status).toHaveText('Reading active file(s)...');

    // Verify reading progress (50% or 100% since intermediate states are sequence-tested)
    await expect(status).toHaveText(/(Reading progress: (50|100)%|Successfully read active file\(s\): 100000 bytes\.)/);

    // Finally, verify successfully read message
    await expect(status).toHaveText('Successfully read active file(s): 100000 bytes.', { timeout: 10000 });
  });
});
