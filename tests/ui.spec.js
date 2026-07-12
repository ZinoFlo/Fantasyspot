const { test, expect } = require('@playwright/test');

test.beforeEach(async ({ page }) => {
  // Mock Office.js before the page loads
  await page.addInitScript(() => {
    window.Office = {
      onReady: (callback) => {
        // Simulate a slight delay for initialization
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
            // Simulate reading a 128KB file (2 slices of 64KB)
            const mockFile = {
              size: 131072,
              sliceCount: 2,
              getSliceAsync: (sliceIndex, sliceCallback) => {
                setTimeout(() => {
                  sliceCallback({
                    status: 'Succeeded',
                    value: { data: new Uint8Array(65536) }
                  });
                }, 50);
              },
              closeAsync: (closeCallback) => {
                if (closeCallback) closeCallback();
              }
            };
            setTimeout(() => {
              callback({ status: 'Succeeded', value: mockFile });
            }, 50);
          }
        }
      }
    };
  });

  // Intercept the real office.js call to prevent it from loading and overriding our mock
  await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', route => {
    route.fulfill({ body: '// Mock Office.js' });
  });

  await page.goto('/');
});

test('should display initials JV after initialization', async ({ page }) => {
  const initials = page.locator('#initials-display');
  await expect(initials).toHaveText('JV');
});

test('should show correct status messages when reading file(s)', async ({ page }) => {
  // Ensure app is ready
  await expect(page.locator('#initials-display')).toHaveText('JV');

  const readBtn = page.locator('#read-files-btn');
  const status = page.locator('#status');

  await readBtn.click();

  // Initial status
  await expect(status).toHaveText('Reading file(s)...');

  // Progress status (using regex to catch either 50% or 100% or Success)
  // Since we have 2 slices, it should show 50% then 100% then Success.
  await expect(status).toHaveText(/Reading progress: (50|100)%|Successfully read active file\(s\): 131072 bytes\./);

  // Final status
  await expect(status).toHaveText('Successfully read active file(s): 131072 bytes.', { timeout: 5000 });
});
