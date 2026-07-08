const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept Office.js script and mock it
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', route => {
      route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js',
      });
    });

    // Mock Office object and its constants
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
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
                          value: { data: new Uint8Array(50) },
                        });
                      }, 500); // Add delay to capture intermediate status
                    },
                    closeAsync: (closeCallback) => {
                      closeCallback();
                    },
                  },
                });
              }, 100);
            },
          },
        },
      };
    });

    await page.goto('/index.html');
  });

  test('should display initials "JV" after initialization', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should update status message during file reading process', async ({ page }) => {
    // Wait for initialization
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    await readBtn.click();

    const status = page.locator('#status');

    // Check initial reading message
    await expect(status).toHaveText(/Reading active file\(s\).../);

    // Check file size message
    await expect(status).toHaveText(/File size: 100 bytes. Reading 2 slices from the presentation\(s\).../);

    // Check intermediate progress message
    await expect(status).toHaveText(/Reading progress: 50%/);

    // Check terminal success message
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes./);
  });
});
