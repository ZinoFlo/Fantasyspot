const { test, expect } = require('@playwright/test');

test.describe('Office Add-in UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Mock Office.js library
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
      await route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js'
      });
    });

    // Inject Office mock object before the page loads
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
                    getSliceAsync: (sliceIndex, sliceCallback) => {
                      setTimeout(() => {
                        sliceCallback({
                          status: 'Succeeded',
                          value: { data: new Uint8Array(50) }
                        });
                      }, 500); // Delay to capture intermediate progress
                    },
                    closeAsync: (closeCallback) => {
                      setTimeout(() => {
                        closeCallback({ status: 'Succeeded' });
                      }, 100);
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

  test('should read active files and update status messages', async ({ page }) => {
    // Wait for initialization
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await readBtn.click();

    // Verify sequence of status messages
    await expect(status).toHaveText('Reading active file(s)...');
    await expect(status).toHaveText(/File size: 100 bytes\. Reading 2 slice\(s\)\.\.\./);
    await expect(status).toHaveText('Reading progress: 50%');
    await expect(status).toHaveText('Successfully read active file(s): 100 bytes.');
  });
});
