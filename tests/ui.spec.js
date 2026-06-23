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
          callback({ host: null }); // Simulate browser testing environment
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
                    setTimeout(() => {
                      closeCallback();
                    }, 50);
                  }
                }
              });
            }, 200);
          }
        }
      }
    };
  });

  await page.goto('http://localhost:3000/index.html');
});

test('initialization updates initials', async ({ page }) => {
  const initials = page.locator('#initials-display');
  await expect(initials).toHaveText('JV');
});

test('read active file(s) process', async ({ page }) => {
  // Wait for initialization
  await expect(page.locator('#initials-display')).toHaveText('JV');

  const readBtn = page.locator('#read-files-btn');
  const status = page.locator('#status');

  await readBtn.click();

  // Check intermediate states
  await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);
  await expect(status).toHaveText(/Reading progress: 50%/);

  // Check final state
  await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
});
