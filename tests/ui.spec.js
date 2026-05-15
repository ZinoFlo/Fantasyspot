const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery UI', () => {
  test.beforeEach(async ({ page }) => {
    // Intercept Office.js script to prevent external network requests
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', route => {
      route.fulfill({ body: '// Mock Office.js' });
    });

    // Mock the Office object before the page loads
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          // Simulate the Office environment being ready with a slight delay
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
              // Simulate getFileAsync success
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
                      setTimeout(() => {
                        closeCallback();
                      }, 10);
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
  });

  test('should display "JV" initials and handle file reading', async ({ page }) => {
    // 1. Check initials
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');

    // 2. Click "Read Active Files" button
    const readBtn = page.locator('#read-files-btn');
    await expect(readBtn).toHaveText('Read Active Files');
    await readBtn.click();

    // 3. Verify status sequence
    const status = page.locator('#status');
    await expect(status).toHaveText(/Reading active file\(s\).../);

    // Wait for the final success message
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes./, { timeout: 5000 });
  });

  test('button-container should have correct margin-top', async ({ page }) => {
    const container = page.locator('.button-container');
    const marginTop = await container.evaluate(el => window.getComputedStyle(el).marginTop);
    expect(marginTop).toBe('30px');
  });
});
