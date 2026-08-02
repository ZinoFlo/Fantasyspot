const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Office Add-in UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept Office.js external script and block/fulfill it to avoid network requests
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
      await route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mocked Office.js library'
      });
    });

    // Inject the mock Office environment before scripts execute
    await page.addInitScript(() => {
      window.Office = {
        HostType: { PowerPoint: 'PowerPoint' },
        FileType: { Compressed: 'Compressed' },
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        context: {
          document: {
            getFileAsync: (fileType, options, callback) => {
              // Simulate async delay
              setTimeout(() => {
                callback({
                  status: 'succeeded',
                  value: {
                    size: 150000,
                    sliceCount: 3,
                    getSliceAsync: (sliceIndex, sliceCallback) => {
                      setTimeout(() => {
                        sliceCallback({
                          status: 'succeeded',
                          value: {
                            // slice.data must have a .length property matching the slice size
                            data: new Uint8Array(50000)
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
              }, 100);
            }
          }
        },
        onReady: (callback) => {
          // Allow application scripts to load and add event listeners first, then invoke onReady
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
          }, 50);
        }
      };
    });

    await page.goto('/');
  });

  test('should display proper initials (JV) after Office.js initialized', async ({ page }) => {
    // The initials-display defaults to "--" and should become "JV"
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should successfully read active file(s) and show progress', async ({ page }) => {
    // Verify initial layout
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');

    const status = page.locator('#status');
    await expect(status).toHaveText('');

    const readBtn = page.locator('#read-files-btn');
    await readBtn.click();

    // Verify progress steps
    // 1. Loading active file(s)
    await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);

    // 2. File size message
    await expect(status).toHaveText(/Total size of active presentation\(s\): 150000 bytes. Reading 3 slices\.\.\./);

    // 3. Reading progress updates
    await expect(status).toHaveText(/(Reading progress: (33|67|100)%|Successfully read active file\(s\)\. Total byte size of active presentation\(s\): 150000 bytes\.)/);

    // 4. Final success status
    await expect(status).toHaveText(/Successfully read active file\(s\)\. Total byte size of active presentation\(s\): 150000 bytes\./);
  });
});
