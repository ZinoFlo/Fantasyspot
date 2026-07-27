const { test, expect } = require('@playwright/test');

test.describe('Office Add-in UI and File Reading', () => {
  test.beforeEach(async ({ page }) => {
    // Block real office.js loading to avoid external network calls and overriding the mock.
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
      await route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js load success'
      });
    });

    // Mock Office environment globals before index.js runs.
    await page.addInitScript(() => {
      window.Office = {
        HostType: { PowerPoint: 'PowerPoint' },
        FileType: { Compressed: 'Compressed' },
        AsyncResultStatus: { Succeeded: 'Succeeded', Failed: 'Failed' },
        context: {
          document: {}
        },
        onReady: (callback) => {
          // Trigger the callback asynchronously to match real behavior
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
          }, 50);
        }
      };
    });

    await page.goto('/');
  });

  test('initializes UI correctly with JV initials', async ({ page }) => {
    // Wait for initials-display to transition from '--' to 'JV'
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('reads file slices and displays progress successfully', async ({ page }) => {
    // Mock the document methods
    await page.evaluate(() => {
      window.Office.context.document.getFileAsync = (fileType, options, callback) => {
        // Trigger callback asynchronously
        setTimeout(() => {
          callback({
            status: 'Succeeded',
            value: {
              size: 200,
              sliceCount: 2,
              closeAsync: (cb) => {
                setTimeout(() => cb(), 50);
              },
              getSliceAsync: (sliceIndex, cb) => {
                setTimeout(() => {
                  cb({
                    status: 'Succeeded',
                    value: {
                      data: new Uint8Array([65, 66, 67]) // Mock data
                    }
                  });
                }, 100); // 100ms delay to catch intermediate reading progress
              }
            }
          });
        }, 50);
      };
    });

    // Make sure initials have loaded first
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const status = page.locator('#status');
    const readBtn = page.locator('#read-files-btn');

    // Click the Read Active File(s) button
    await readBtn.click();

    // Verify progress transitions and final success text
    await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);
    await expect(status).toHaveText(/Active file\(s\) size: 200 bytes\. Reading 2 slices\.\.\./);
    await expect(status).toHaveText(/Reading progress: (50|100)%/);
    await expect(status).toHaveText(/Successfully read active file\(s\): 200 bytes\./);
  });
});
