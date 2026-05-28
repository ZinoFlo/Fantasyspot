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
                        value: { data: new Uint8Array(50) }
                      });
                    }, 200); // 200ms delay to help capture intermediate states if needed
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

  await page.goto('/');
});

test('initialization updates initials', async ({ page }) => {
  const initials = page.locator('#initials-display');
  await expect(initials).toHaveText('JV');
});

test('read active files functionality', async ({ page }) => {
  // Wait for initialization
  await expect(page.locator('#initials-display')).toHaveText('JV');

  const readBtn = page.locator('#read-files-btn');
  const status = page.locator('#status');

  await readBtn.click();

  // Check initial reading message
  await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);

  // Check file size/slices message
  await expect(status).toHaveText(/File size: 100 bytes\. Reading 2 slice\(s\)\.\.\./);

  // Check progress (assert on a reliable intermediate state)
  await expect(status).toHaveText(/Reading progress: 50%/);

  // Check success message (terminal state)
  await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
});

test('ui layout: button is centered', async ({ page }) => {
  const button = page.locator('#read-files-btn');
  const box = await button.boundingBox();
  const viewport = page.viewportSize();

  // Check if button center is roughly at viewport center
  const buttonCenter = box.x + box.width / 2;
  const viewportCenter = viewport.width / 2;

  expect(Math.abs(buttonCenter - viewportCenter)).toBeLessThan(5);
});
