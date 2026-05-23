const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery UI', () => {
  test.beforeEach(async ({ page }) => {
    // Mock Office.js
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
      await route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js',
      });
    });

    // Inject Office mock
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
          }, 100);
        },
        HostType: {
          PowerPoint: 'PowerPoint',
        },
        FileType: {
          Compressed: 'compressed',
        },
        AsyncResultStatus: {
          Succeeded: 'succeeded',
          Failed: 'failed',
        },
        context: {
          document: {
            getFileAsync: (fileType, options, callback) => {
              setTimeout(() => {
                callback({
                  status: 'succeeded',
                  value: {
                    size: 100,
                    sliceCount: 2,
                    getSliceAsync: (sliceIndex, sliceCallback) => {
                      setTimeout(() => {
                        sliceCallback({
                          status: 'succeeded',
                          value: {
                            data: new Uint8Array(50),
                          },
                        });
                      }, 50);
                    },
                    closeAsync: (closeCallback) => {
                      setTimeout(closeCallback, 50);
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

  test('should display correctly initialized initials', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should have the "Read Active Files" button', async ({ page }) => {
    const button = page.locator('#read-files-btn');
    await expect(button).toHaveText('Read Active Files');
  });

  test('should update status when reading files', async ({ page }) => {
    const button = page.locator('#read-files-btn');
    const status = page.locator('#status');

    // Wait for initialization
    await expect(page.locator('#initials-display')).toHaveText('JV');

    await button.click();

    // Asserting on the final state is most reliable in this environment.
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes./);
  });

  test('button should be wrapped in button-container and centered', async ({ page }) => {
    const container = page.locator('.button-container');
    await expect(container).toBeVisible();

    const button = page.locator('#read-files-btn');
    const box = await button.boundingBox();
    const viewport = page.viewportSize();

    // Check if horizontally centered (approx)
    const buttonCenter = box.x + box.width / 2;
    const viewportCenter = viewport.width / 2;
    expect(Math.abs(buttonCenter - viewportCenter)).toBeLessThan(5);

    // Check margin-top on container
    const marginTop = await container.evaluate((el) => window.getComputedStyle(el).marginTop);
    expect(marginTop).toBe('30px');
  });
});
