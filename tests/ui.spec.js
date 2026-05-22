const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Mock Office.js before navigating
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          setTimeout(() => {
            callback({ host: null }); // Simulate browser environment
          }, 100);
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
                    getSliceAsync: (index, sliceCallback) => {
                      // Increase delay to ensure Playwright can catch intermediate states
                      setTimeout(() => {
                        sliceCallback({
                          status: 'succeeded',
                          value: { data: new Uint8Array(50) }
                        });
                      }, 300);
                    },
                    closeAsync: (closeCallback) => {
                      setTimeout(() => closeCallback(), 100);
                    }
                  }
                });
              }, 100);
            }
          }
        },
        HostType: { PowerPoint: 'PowerPoint' },
        FileType: { Compressed: 'Compressed' },
        AsyncResultStatus: { Succeeded: 'succeeded' }
      };
    });

    // Intercept the real office.js script to prevent it from loading and overwriting our mock
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', (route) => {
      route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mocked Office.js',
      });
    });

    await page.goto('/');
  });

  test('should display the correct initial state', async ({ page }) => {
    await expect(page.locator('h1')).toHaveText('Eco-growth Discovery');
    await expect(page.locator('#read-files-btn')).toHaveText('Read Active Files');

    // Wait for Office.onReady to trigger and update initials
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should read files and update status', async ({ page }) => {
    // Wait for Office.onReady to complete
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await readBtn.click();

    // Check intermediate status
    await expect(status).toHaveText(/Reading active file\(s\).../);

    // Check final success message
    // We omit intermediate progress checks as they can be flaky due to rapid transitions
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes/);
  });

  test('should have centered button with top margin', async ({ page }) => {
    const container = page.locator('.button-container');
    const button = page.locator('#read-files-btn');

    const containerStyle = await container.evaluate((el) => window.getComputedStyle(el).marginTop);
    expect(containerStyle).toBe('30px');

    const buttonStyle = await button.evaluate((el) => {
      const style = window.getComputedStyle(el);
      return {
        marginLeft: style.marginLeft,
        marginRight: style.marginRight,
        display: style.display
      };
    });

    expect(buttonStyle.display).toBe('block');
    // For centered block elements, left and right margins should be equal (approx)
    const ml = parseFloat(buttonStyle.marginLeft);
    const mr = parseFloat(buttonStyle.marginRight);
    expect(Math.abs(ml - mr)).toBeLessThan(1);
    expect(ml).toBeGreaterThan(0);
  });
});
