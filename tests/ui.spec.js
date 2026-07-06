const { test, expect } = require('@playwright/test');

test.beforeEach(async ({ page }) => {
  // Mock Office.js before the page loads
  await page.addInitScript(() => {
    window.Office = {
      onReady: (callback) => {
        // Simulate Office initialization with a slight delay
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
                    }, 50);
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
      }
    };
  });

  // Intercept the real office.js script to avoid network dependencies
  await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', (route) => {
    route.fulfill({ body: '// Mock Office.js' });
  });

  await page.goto('/');
});

test('should initialize and show initials JV', async ({ page }) => {
  const initials = page.locator('#initials-display');
  await expect(initials).toHaveText('JV');
});

test('should read file successfully', async ({ page }) => {
  // Wait for initialization
  await expect(page.locator('#initials-display')).toHaveText('JV');

  const readBtn = page.locator('#read-files-btn');
  await readBtn.click();

  const status = page.locator('#status');
  await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);

  // Verify progress updates and final success message
  await expect(status).toHaveText(/Reading progress: 50%/, { timeout: 5000 });
  await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./, { timeout: 5000 });
});
