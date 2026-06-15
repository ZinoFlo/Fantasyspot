const { test, expect } = require('@playwright/test');

test.beforeEach(async ({ page }) => {
  // Mock Office.js
  await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
    await route.fulfill({ body: '// Mock Office.js' });
  });

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
                    }, 500); // 500ms delay to capture intermediate status
                  },
                  closeAsync: (closeCallback) => {
                    setTimeout(() => {
                      closeCallback();
                    }, 100);
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

test('should initialize with JV initials', async ({ page }) => {
  const initials = page.locator('#initials-display');
  await expect(initials).toHaveText('JV');
});

test('should read active files successfully', async ({ page }) => {
  // Wait for initialization
  await expect(page.locator('#initials-display')).toHaveText('JV');

  const readBtn = page.locator('#read-files-btn');
  const status = page.locator('#status');

  await readBtn.click();

  // Check initial status
  await expect(status).toHaveText('Reading active file(s)...');

  // Check file size info
  await expect(status).toHaveText(/File size: 100 bytes\. Reading 2 slice\(s\)\.\.\./);

  // Check progress
  await expect(status).toHaveText('Reading progress: 50%');

  // Check success message
  await expect(status).toHaveText('Successfully read active file(s): 100 bytes.');
});
