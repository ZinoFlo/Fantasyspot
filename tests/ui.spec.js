const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Add-in', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept Office.js script to prevent external network requests
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', route => {
      route.fulfill({ body: '// Mock Office.js' });
    });

    // Inject Office mock before any script runs
    await page.addInitScript(() => {
      window.Office = {
        onReady: (cb) => {
          setTimeout(() => {
            cb({ host: null }); // Simulate browser environment
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
                    getSliceAsync: (index, cb) => {
                      setTimeout(() => {
                        cb({
                          status: 'Succeeded',
                          value: { data: new Uint8Array(50) }
                        });
                      }, 500); // 500ms delay to capture intermediate states
                    },
                    closeAsync: (cb) => {
                      if (cb) cb();
                    }
                  }
                });
              }, 100);
            }
          }
        }
      };
    });

    await page.goto('/index.html');
  });

  test('should initialize with initials JV', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should read active file(s) successfully', async ({ page }) => {
    // Wait for initialization
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await readBtn.click();

    // Verify initial "Reading" message
    await expect(status).toHaveText('Reading active file(s)...');

    // Verify progress message
    await expect(status).toHaveText(/Reading progress: 50%/);

    // Verify final success message
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes/);
  });
});
