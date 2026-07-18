const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Add-in UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept office.js script loading and fulfill with empty mock script
    // to allow our custom window.Office mock injected via addInitScript to be used.
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
      await route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js',
      });
    });

    // Mock the Office namespace before the page loads
    await page.addInitScript(() => {
      window.Office = {
        HostType: {
          PowerPoint: 'PowerPoint',
        },
        FileType: {
          Compressed: 'Compressed',
        },
        AsyncResultStatus: {
          Succeeded: 'Succeeded',
          Failed: 'Failed',
        },
        onReady: (callback) => {
          // Initialize with a slight delay to allow scripts to load and attach event listeners
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
          }, 100);
        },
        context: {
          document: {
            getFileAsync: (fileType, options, callback) => {
              // Simulate async file retrieval with a mock file object
              setTimeout(() => {
                const mockFile = {
                  size: 150000, // 150 KB
                  sliceCount: 3,
                  getSliceAsync: (sliceIndex, sliceCallback) => {
                    setTimeout(() => {
                      // Generate a block of data for each slice
                      const sliceSize = 65536;
                      const dataSize = (sliceIndex === 2) ? (150000 - 2 * sliceSize) : sliceSize;
                      const sliceData = new Uint8Array(dataSize);
                      sliceCallback({
                        status: 'Succeeded',
                        value: {
                          data: Array.from(sliceData),
                          index: sliceIndex,
                        },
                      });
                    }, 500); // 500ms delay to capture intermediate progress states
                  },
                  closeAsync: (closeCallback) => {
                    setTimeout(() => {
                      closeCallback({ status: 'Succeeded' });
                    }, 50);
                  },
                };
                callback({
                  status: 'Succeeded',
                  value: mockFile,
                });
              }, 100);
            },
          },
        },
      };
    });

    await page.goto('/');
  });

  test('should display initials "JV" after initialization', async ({ page }) => {
    const initials = page.locator('#initials-display');
    // Verify initials transition to 'JV' after Office.onReady callback fires
    await expect(initials).toHaveText('JV');
  });

  test('should read file slices and update progress and success message', async ({ page }) => {
    // Wait for add-in to initialize
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const status = page.locator('#status');
    const readBtn = page.locator('#read-files-btn');

    // Click the Read button
    await readBtn.click();

    // Verify progress text and final success message using regex
    await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);

    // Verify intermediate progress states and success message
    await expect(status).toHaveText(/(Reading progress: (33|67|100)%|Successfully read active file\(s\): 150000 bytes\.)/);

    // Verify final success message
    await expect(status).toHaveText('Successfully read active file(s): 150000 bytes.', { timeout: 5000 });
  });
});
