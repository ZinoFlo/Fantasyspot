const { test, expect } = require('@playwright/test');

test.beforeEach(async ({ page }) => {
  // Mock Office.js before the page loads
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
                  getSliceAsync: (sliceIndex, sliceCallback) => {
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
                    }, 100);
                  }
                }
              });
            }, 500);
          }
        }
      }
    };
  });

  // Intercept the external office.js script to prevent network calls
  await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', (route) => {
    route.fulfill({
      status: 200,
      contentType: 'application/javascript',
      body: '// Mock Office.js'
    });
  });

  await page.goto('/');
});

test('initialization updates initials display', async ({ page }) => {
  const initials = page.locator('#initials-display');
  await expect(initials).toHaveText('JV');
});

test('clicking Read Active File(s) button triggers reading process', async ({ page }) => {
  // Wait for initialization
  await expect(page.locator('#initials-display')).toHaveText('JV');

  const readBtn = page.locator('#read-files-btn');
  const status = page.locator('#status');

  await readBtn.click();

  // Check initial status
  await expect(status).toHaveText('Reading active file(s)...');

  // Check file size and slice count message
  await expect(status).toHaveText(/File size: 100 bytes. Reading 2 slices.../);

  // Check intermediate progress (due to 500ms delay in getSliceAsync mock)
  await expect(status).toHaveText('Reading progress: 50%');

  // Check final success message
  await expect(status).toHaveText('Successfully read active file(s): 100 bytes.');
});
