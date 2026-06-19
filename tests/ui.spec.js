const { test, expect } = require('@playwright/test');

test('should initialize and read file correctly', async ({ page }) => {
  // Mock Office.js library
  await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
    await route.fulfill({ body: '// Mock Office.js' });
  });

  // Inject Office mock
  await page.addInitScript(() => {
    window.Office = {
      HostType: { PowerPoint: 'PowerPoint' },
      FileType: { Compressed: 'Compressed' },
      AsyncResultStatus: { Succeeded: 'Succeeded', Failed: 'Failed' },
      onReady: (callback) => {
        // Delay slightly to ensure app scripts are loaded
        setTimeout(() => {
          callback({ host: 'PowerPoint' });
        }, 100);
      },
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
                    }, 100);
                  },
                  closeAsync: (closeCallback) => {
                    setTimeout(() => {
                      closeCallback({ status: 'Succeeded' });
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

  await page.goto('http://localhost:3000/index.html');

  // Verify initialization
  const initialsDisplay = page.locator('#initials-display');
  await expect(initialsDisplay).toHaveText('JV');

  // Trigger file reading
  const readBtn = page.locator('#read-files-btn');
  await readBtn.click();

  // Verify status messages
  const status = page.locator('#status');
  await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);
  await expect(status).toHaveText(/File size: 100 bytes\. Reading 2 slices\.\.\./);

  // Wait for progress updates
  // Note: '100%' progress might be skipped by Playwright due to rapid transition to success message.
  await expect(status).toHaveText(/Reading progress: 50%/);

  // Verify final success message
  await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
});
