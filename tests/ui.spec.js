const { test, expect } = require('@playwright/test');

test.describe('Office Add-in UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Mock the Office.js library
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', (route) => {
      route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js content'
      });
    });

    // Inject the Office mock before the page loads
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
          }, 100);
        },
        HostType: { PowerPoint: 'PowerPoint' },
        FileType: { Compressed: 'Compressed' },
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
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
                      setTimeout(() => {
                        sliceCallback({
                          status: 'succeeded',
                          value: { data: new Uint8Array(50) }
                        });
                      }, 500); // Delay to capture intermediate status
                    },
                    closeAsync: (closeCallback) => {
                      closeCallback();
                    }
                  }
                });
              }, 500); // Delay to capture initial status
            }
          }
        }
      };
    });

    await page.goto('http://localhost:3000/index.html');
  });

  test('should initialize and display initials', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should read active files and update status', async ({ page }) => {
    // Wait for initialization
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    await readBtn.click();

    const status = page.locator('#status');

    // 1. Initial status
    await expect(status).toHaveText('Reading active file(s)...');

    // 2. File size and slice count info
    await expect(status).toHaveText(/File size: 100 bytes\. Reading 2 slices\.\.\./);

    // 3. Progress updates (capturing at least one)
    await expect(status).toHaveText(/Reading progress: (50|100)%/);

    // 4. Success message
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
  });
});
