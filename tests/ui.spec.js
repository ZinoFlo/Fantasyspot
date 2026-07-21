const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Taskpane UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept external Office.js load to prevent fetching external scripts
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
      await route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js library loader',
      });
    });

    // Inject mock Office object BEFORE document starts loading/parsing scripts
    await page.addInitScript(() => {
      window.Office = {
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
        onReady: (callback) => {
          // Trigger the callback asynchronously to simulate Office initialization flow
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
          }, 100);
        },
        context: {
          document: {
            getFileAsync: (fileType, options, callback) => {
              // Simulate small delay for reading active file(s)
              setTimeout(() => {
                callback({
                  status: 'Succeeded',
                  value: {
                    size: 131072, // 128 KB
                    sliceCount: 2,
                    getSliceAsync: (sliceIndex, sliceCallback) => {
                      setTimeout(() => {
                        sliceCallback({
                          status: 'Succeeded',
                          value: {
                            data: new Uint8Array(65536) // 64 KB slice data
                          }
                        });
                      }, 100);
                    },
                    closeAsync: (closeCallback) => {
                      setTimeout(() => {
                        closeCallback({ status: 'Succeeded' });
                      }, 50);
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

  test('should display initials "JV" after Office initialization', async ({ page }) => {
    // Wait for the initials to transition from '--' to 'JV'
    const initialsDisplay = page.locator('#initials-display');
    await expect(initialsDisplay).toHaveText('JV');
  });

  test('should perform full file reading flow and show progress updates', async ({ page }) => {
    // Wait for Office to be ready
    const initialsDisplay = page.locator('#initials-display');
    await expect(initialsDisplay).toHaveText('JV');

    const status = page.locator('#status');
    const readBtn = page.locator('#read-files-btn');

    // Initially status should be empty
    await expect(status).toBeEmpty();

    // Click on Read Active File(s)
    await readBtn.click();

    // Verify transition of states using regex matchers or sequential asserts
    // Stage 1: Initiating file reading
    await expect(status).toHaveText('Reading active file(s)...');

    // Stage 2: Retreived file size info
    await expect(status).toHaveText(/Active file\(s\) size: 131072 bytes\. Reading 2 slices\.\.\./);

    // Stage 3 & 4: Progress stages (e.g. 50% -> 100%) and then success message
    await expect(status).toHaveText(/(Reading progress: (50|100)%|Successfully read active file\(s\): 131072 bytes\.)/);

    // Ensure final state is reached
    await expect(status).toHaveText('Successfully read active file(s): 131072 bytes.');
  });
});
