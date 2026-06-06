// @ts-check
const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Taskpane', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept Office.js to prevent external network requests and errors
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
      await route.fulfill({ body: '// Mock Office.js' });
    });

    // Mock the Office environment
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          // Simulate Office environment being ready with a slight delay
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
          }, 100);
        },
        HostType: { PowerPoint: 'PowerPoint' },
        FileType: { Compressed: 'compressed' },
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        context: {
          document: {
            getFileAsync: (fileType, options, callback) => {
              // Mock successful file retrieval
              setTimeout(() => {
                callback({
                  status: 'succeeded',
                  value: {
                    size: 100,
                    sliceCount: 2,
                    getSliceAsync: (index, sliceCallback) => {
                      // Mock slice retrieval with a delay to test intermediate status
                      setTimeout(() => {
                        sliceCallback({
                          status: 'succeeded',
                          value: { data: new Array(50).fill(0) }
                        });
                      }, 500);
                    },
                    closeAsync: (closeCallback) => {
                      setTimeout(() => closeCallback(), 10);
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

  test('should initialize and display initials', async ({ page }) => {
    // Verify initials transition from '--' to 'JV'
    await expect(page.locator('#initials-display')).toHaveText('JV');
  });

  test('should read active files and display progress', async ({ page }) => {
    // Wait for initialization
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await readBtn.click();

    // Verify initial "Reading" status
    await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);

    // Verify file size and slice count message
    await expect(status).toHaveText(/File size: 100 bytes\. Reading 2 slices\.\.\./);

    // Verify progress update (50% for first slice)
    await expect(status).toHaveText(/Reading progress: 50%/);

    // Verify final success message
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
  });
});
