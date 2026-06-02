const { test, expect } = require('@playwright/test');

test.describe('Office Add-in UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept Office.js request
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', route => {
      route.fulfill({ body: '// Mock Office.js' });
    });

    // Mock Office environment
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
                    size: 200,
                    sliceCount: 2,
                    getSliceAsync: (index, sliceCallback) => {
                      setTimeout(() => {
                        sliceCallback({
                          status: 'Succeeded',
                          value: { data: new Uint8Array(100) }
                        });
                      }, 500); // Delay to capture progress updates
                    },
                    closeAsync: (closeCallback) => {
                      setTimeout(() => closeCallback(), 100);
                    }
                  }
                });
              }, 500);
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

  test('should read active files and update status', async ({ page }) => {
    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await expect(page.locator('#initials-display')).toHaveText('JV');
    await readBtn.click();

    // Verify sequence of status messages
    await expect(status).toHaveText('Reading active file(s)...');
    await expect(status).toHaveText(/File size: 200 bytes. Reading 2 slices.../);
    await expect(status).toHaveText('Reading progress: 50%');
    // Note: '100%' might be skipped by Playwright's polling if the success message follows too quickly
    await expect(status).toHaveText(/Successfully read active file\(s\): 200 bytes./);
  });
});
