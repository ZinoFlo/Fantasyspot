const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery UI', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept Office.js script and return a mock
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async route => {
      await route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js',
      });
    });

    // Inject mock Office object before scripts run
    await page.addInitScript(() => {
      window.Office = {
        onReady: (cb) => {
          // Trigger onReady with a slight delay
          setTimeout(() => {
            cb({ host: null });
          }, 100);
        },
        HostType: { PowerPoint: 'PowerPoint' },
        AsyncResultStatus: { Succeeded: 'Succeeded', Failed: 'Failed' },
        FileType: { Compressed: 'Compressed' },
        context: {
          document: {
            getFileAsync: (fileType, options, callback) => {
              callback({
                status: 'Succeeded',
                value: {
                  size: 100,
                  sliceCount: 1,
                  getSliceAsync: (index, cb) => {
                    cb({
                      status: 'Succeeded',
                      value: { data: new Uint8Array(100) },
                    });
                  },
                  closeAsync: (cb) => cb(),
                },
              });
            },
          },
        },
      };
    });

    await page.goto('/index.html');
  });

  test('should display correct title and button text', async ({ page }) => {
    await expect(page.locator('h1')).toHaveText('Eco-growth Discovery');
    await expect(page.locator('#read-files-btn')).toHaveText('Read Active Files');
  });

  test('should initialize initials and display "JV"', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should show success message when button is clicked', async ({ page }) => {
    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await readBtn.click();
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes/);
  });

  test('button container should have correct margin-top', async ({ page }) => {
    const container = page.locator('.button-container');
    const marginTop = await container.evaluate((el) => window.getComputedStyle(el).marginTop);
    expect(marginTop).toBe('30px');
  });

  test('button should be centered', async ({ page }) => {
    const button = page.locator('#read-files-btn');
    const styles = await button.evaluate((el) => {
      const s = window.getComputedStyle(el);
      return {
        marginLeft: parseFloat(s.marginLeft),
        marginRight: parseFloat(s.marginRight),
        display: s.display
      };
    });

    expect(styles.display).toBe('block');
    // For auto margin, we expect both to be equal and non-zero (if viewport > button width)
    expect(styles.marginLeft).toBeGreaterThan(0);
    expect(Math.abs(styles.marginLeft - styles.marginRight)).toBeLessThan(1);
  });
});
