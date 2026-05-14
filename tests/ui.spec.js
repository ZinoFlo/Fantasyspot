const { test, expect } = require('@playwright/test');

test('Verify pluralization and UI updates', async ({ page }) => {
  // Mock Office.js and the Office object
  await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
    await route.fulfill({ body: '// Mock Office.js' });
  });

  await page.addInitScript(() => {
    window.Office = {
      onReady: (callback) => {
        setTimeout(() => callback({ host: 'PowerPoint' }), 100);
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
                  sliceCount: 1,
                  getSliceAsync: (index, sliceCallback) => {
                    setTimeout(() => {
                      sliceCallback({
                        status: 'Succeeded',
                        value: { data: new Uint8Array(100) }
                      });
                    }, 50);
                  },
                  closeAsync: (closeCallback) => {
                    setTimeout(() => closeCallback(), 10);
                  }
                }
              });
            }, 50);
          }
        }
      }
    };
  });

  await page.goto('http://localhost:3000/index.html');

  // Verify initials updated (confirming Office.onReady fired)
  const initials = page.locator('#initials-display');
  await expect(initials).toHaveText('JV');

  // Verify button text is pluralized
  const button = page.locator('#read-files-btn');
  await expect(button).toHaveText('Read Active Files');

  // Verify CSS centering and container margin
  const container = page.locator('.button-container');
  const marginTop = await container.evaluate((el) => window.getComputedStyle(el).marginTop);
  expect(marginTop).toBe('30px');

  // Click button and verify status messages
  await button.click();

  const status = page.locator('#status');
  await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./, { timeout: 5000 });
});
