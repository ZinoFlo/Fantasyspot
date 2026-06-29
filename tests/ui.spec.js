const { test, expect } = require('@playwright/test');

test.beforeEach(async ({ page }) => {
  // Mock Office.js
  await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
    await route.fulfill({
      status: 200,
      contentType: 'application/javascript',
      body: '// Mock Office.js',
    });
  });

  // Inject Office mock before the script loads
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
                        value: { data: new Array(50).fill(0) },
                      });
                    }, 500); // 500ms delay to capture progress status
                  },
                  closeAsync: (closeCallback) => {
                    closeCallback();
                  },
                },
              });
            }, 100);
          },
        },
      },
    };
  });

  await page.goto('/index.html');
});

test('initialization displays JV initials', async ({ page }) => {
  const initialsDisplay = page.locator('#initials-display');
  await expect(initialsDisplay).toHaveText('JV');
});

test('reading active file(s) displays progress and success', async ({ page }) => {
  // Wait for initialization
  await expect(page.locator('#initials-display')).toHaveText('JV');

  const readBtn = page.locator('#read-files-btn');
  await readBtn.click();

  const status = page.locator('#status');

  // Assert intermediate progress (Reading...)
  await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);

  // Assert specific progress updates
  await expect(status).toHaveText(/Reading progress: 50%/);

  // Assert success message
  await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
});
