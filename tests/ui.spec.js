const { test, expect } = require('@playwright/test');

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
                        value: { data: new Array(50).fill(0) }
                      });
                    }, 50);
                  },
                  closeAsync: (closeCallback) => {
                    closeCallback();
                  }
                }
              });
            }, 100);
          }
        }
      }
    };
  });

  await page.goto('/index.html');
});

test('initializes and displays initials', async ({ page }) => {
  const initials = page.locator('#initials-display');
  await expect(initials).toHaveText('JV');
});

test('reads active file(s) successfully', async ({ page }) => {
  // Ensure initialized
  await expect(page.locator('#initials-display')).toHaveText('JV');

  const readBtn = page.locator('#read-files-btn');
  const status = page.locator('#status');

  await readBtn.click();

  // Check intermediate state
  await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);

  // Wait for success message
  await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
});
