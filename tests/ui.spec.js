const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery UI', () => {
  test.beforeEach(async ({ page }) => {
    // Mock Office.js
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
      await route.fulfill({ body: '// Mock Office.js' });
    });

    // Inject Office mock
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          setTimeout(() => {
            callback({ host: null }); // Simulate browser environment
          }, 100);
        },
        HostType: { PowerPoint: 'PowerPoint' },
        context: {
          document: {
            getFileAsync: (fileType, options, callback) => {
              setTimeout(() => {
                callback({
                  status: 'succeeded',
                  value: {
                    size: 1000,
                    sliceCount: 2,
                    getSliceAsync: (index, sliceCallback) => {
                      setTimeout(() => {
                        sliceCallback({
                          status: 'succeeded',
                          value: { data: new Uint8Array(500) }
                        });
                      }, 200);
                    },
                    closeAsync: (closeCallback) => {
                      setTimeout(() => {
                        closeCallback();
                      }, 10);
                    }
                  }
                });
              }, 100);
            }
          }
        },
        FileType: { Compressed: 'compressed' },
        AsyncResultStatus: { Succeeded: 'succeeded' }
      };
    });

    await page.goto('/');
  });

  test('should display initials JV after initialization', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should have the correct button label', async ({ page }) => {
    const button = page.locator('#read-files-btn');
    await expect(button).toHaveText('Read Active Files');
  });

  test('should update status message when reading files', async ({ page }) => {
    const button = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await button.click();

    // Check intermediate state
    await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);

    // Check final state
    await expect(status).toHaveText(/Successfully read active file\(s\): 1000 bytes across 2 slice\(s\)\./);
  });

  test('button should be centered and have container with margin', async ({ page }) => {
    const container = page.locator('.button-container');
    const button = page.locator('#read-files-btn');

    const containerMarginTop = await container.evaluate(el => window.getComputedStyle(el).marginTop);
    expect(containerMarginTop).toBe('30px');

    const buttonMarginTop = await button.evaluate(el => window.getComputedStyle(el).marginTop);
    expect(buttonMarginTop).toBe('0px');

    const buttonMarginLeft = await button.evaluate(el => parseInt(window.getComputedStyle(el).marginLeft));
    const buttonMarginRight = await button.evaluate(el => parseInt(window.getComputedStyle(el).marginRight));

    expect(buttonMarginLeft).toBeGreaterThan(0);
    expect(Math.abs(buttonMarginLeft - buttonMarginRight)).toBeLessThanOrEqual(1);
  });
});
