const { test, expect } = require('@playwright/test');

test.beforeEach(async ({ page }) => {
  // Mock Office.js script
  await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', route => {
    route.fulfill({
      status: 200,
      contentType: 'application/javascript',
      body: '// Mock Office.js',
    });
  });

  // Inject Office mock before scripts run
  await page.addInitScript(() => {
    window.Office = {
      onReady: (callback) => {
        // Simulate slightly delayed initialization
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
                    }, 500);
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

  await page.goto('/index.html');
});

test('initials display updates to JV', async ({ page }) => {
  const initials = page.locator('#initials-display');
  await expect(initials).toHaveText('JV', { timeout: 5000 });
});

test('read active files workflow', async ({ page }) => {
  // Wait for initialization
  await expect(page.locator('#initials-display')).toHaveText('JV');

  const readBtn = page.locator('#read-files-btn');
  const status = page.locator('#status');

  await readBtn.click();

  // Check sequence of status messages
  await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);
  // Using a less strict sequence check because transitions can be very fast
  await expect(status).toHaveText(/Reading progress: 50%/);
  await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
});
