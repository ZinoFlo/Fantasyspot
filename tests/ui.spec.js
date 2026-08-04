const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Taskpane UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept external Office.js script and return mock body to avoid external network dependencies
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', (route) => {
      route.fulfill({
        contentType: 'application/javascript',
        body: '// Mock Office.js'
      });
    });

    // Inject the mocked Office namespace before any script runs
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
          }, 100);
        },
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
        context: {
          document: {
            getFileAsync: (fileType, options, callback) => {
              const mockFile = {
                size: 100,
                sliceCount: 2,
                getSliceAsync: (sliceIndex, sliceCallback) => {
                  setTimeout(() => {
                    sliceCallback({
                      status: 'Succeeded',
                      value: {
                        data: new Uint8Array(50) // 2 slices of 50 bytes
                      }
                    });
                  }, 500); // 500ms delay to ensure intermediate states are captured
                },
                closeAsync: (closeCallback) => {
                  setTimeout(() => {
                    closeCallback();
                  }, 50);
                }
              };
              setTimeout(() => {
                callback({
                  status: 'Succeeded',
                  value: mockFile
                });
              }, 100);
            }
          }
        }
      };
    });

    await page.goto('/');
  });

  test('should display initials after Office initialization', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should read file and display status progression', async ({ page }) => {
    // Ensure Office initialized first
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await readBtn.click();

    // Check initial reading state
    await expect(status).toHaveText('Reading active file(s)...');

    // Wait for the file size text
    await expect(status).toHaveText(/File size: 100 bytes\. Reading 2 slices\.\.\./);

    // Wait for progress (Reading progress: 50%, then 100%) or final success message
    // A regex that matches the sequence or final success
    await expect(status).toHaveText(/(Reading progress: (50|100)%|Successfully read active file\(s\): 100 bytes\.)/);

    // Wait until the final success status is reached
    await expect(status).toHaveText('Successfully read active file(s): 100 bytes.');
  });
});
