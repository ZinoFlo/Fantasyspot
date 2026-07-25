const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Office Add-in UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // 1. Mock the external Office.js script request to prevent network dependencies
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
      await route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mocked Office.js script',
      });
    });

    // 2. Inject the global Office object mock before the page scripts run
    await page.addInitScript(() => {
      const OfficeMock = {
        HostType: { PowerPoint: 'PowerPoint' },
        FileType: { Compressed: 'Compressed' },
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        onReady: (callback) => {
          // Trigger the ready callback with a slight delay to allow scripts to load
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
          }, 100);
        },
        context: {
          document: {
            getFileAsync: (fileType, options, callback) => {
              // Simulated file with size 131072 bytes (128 KB) and 2 slices of 64 KB
              const mockFile = {
                size: 131072,
                sliceCount: 2,
                getSliceAsync: (sliceIndex, sliceCallback) => {
                  // Use a 500ms delay in getSliceAsync to ensure the UI updates are visible
                  setTimeout(() => {
                    sliceCallback({
                      status: 'succeeded',
                      value: {
                        data: new Array(65536).fill(0), // 64 KB slice
                      },
                    });
                  }, 500);
                },
                closeAsync: (closeCallback) => {
                  setTimeout(() => {
                    closeCallback({ status: 'succeeded' });
                  }, 50);
                }
              };

              setTimeout(() => {
                callback({
                  status: 'succeeded',
                  value: mockFile,
                });
              }, 100);
            },
          },
        },
      };

      window.Office = OfficeMock;
    });
  });

  test('should initialize and display initials "JV"', async ({ page }) => {
    await page.goto('/');

    // Verify initials Display is updated to 'JV' after Office.onReady runs
    const initialsDisplay = page.locator('#initials-display');
    await expect(initialsDisplay).toHaveText('JV', { timeout: 5000 });
  });

  test('should read active file(s) and display progression and success', async ({ page }) => {
    await page.goto('/');

    // Wait for the mock to initialize and display initials 'JV'
    const initialsDisplay = page.locator('#initials-display');
    await expect(initialsDisplay).toHaveText('JV');

    const statusLocator = page.locator('#status');
    await expect(statusLocator).toBeEmpty();

    // Click "Read Active File(s)" button
    const readBtn = page.locator('#read-files-btn');
    await readBtn.click();

    // Verify initial "Reading active file(s)..." message
    await expect(statusLocator).toHaveText('Reading active file(s)...');

    // Verify file details and loading progression. Since there are 2 slices, progress will show 50% and 100%.
    // We expect transition state checks:
    await expect(statusLocator).toHaveText(/(File size: 131072 bytes. Reading 2 slices...|Reading progress: (50|100)%|Successfully read active file\(s\): 131072 bytes.)/, { timeout: 5000 });

    // Wait for the success message to be finally displayed
    await expect(statusLocator).toHaveText('Successfully read active file(s): 131072 bytes.', { timeout: 5000 });
  });
});
