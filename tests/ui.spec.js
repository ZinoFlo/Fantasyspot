const { test, expect } = require('@playwright/test');

test.beforeEach(async ({ page }) => {
  // Mock Office.js
  await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', route => {
    route.fulfill({
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
                  getSliceAsync: (index, sliceCallback) => {
                    setTimeout(() => {
                      sliceCallback({
                        status: 'Succeeded',
                        value: { data: new Uint8Array(50) }
                      });
                    }, 500);
                  },
                  closeAsync: (closeCallback) => {
                    closeCallback();
                  }
                }
              });
            }, 200);
          }
        }
      }
    };
  });

  await page.goto('/');
});

test('should initialize with correct initials', async ({ page }) => {
  const initials = page.locator('#initials-display');
  await expect(initials).toHaveText('JV');
});

test('should read active files successfully', async ({ page }) => {
  const readBtn = page.locator('#read-files-btn');
  const status = page.locator('#status');

  await expect(page.locator('#initials-display')).toHaveText('JV');

  await readBtn.click();

  // Check intermediate status
  await expect(status).toHaveText('Reading active file(s)...');

  // Check progress
  await expect(status).toHaveText(/Reading progress: 50%/, { timeout: 10000 });

  // Final success message
  await expect(status).toHaveText('Successfully read active file(s): 100 bytes.');
});

test('should have correct button container styling', async ({ page }) => {
  const container = page.locator('.button-container');
  await expect(container).toHaveCSS('margin-top', '30px');

  const button = page.locator('#read-files-btn');
  const marginLeft = await button.evaluate(el => window.getComputedStyle(el).marginLeft);
  const marginRight = await button.evaluate(el => window.getComputedStyle(el).marginRight);

  expect(parseFloat(marginLeft)).toBeGreaterThan(0);
  expect(Math.abs(parseFloat(marginLeft) - parseFloat(marginRight))).toBeLessThan(1);
});
