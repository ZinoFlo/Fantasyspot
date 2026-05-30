import { test, expect } from '@playwright/test';

test.describe('Eco-growth Discovery UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Mock Office.js
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async route => {
      await route.fulfill({ body: '// Mock Office.js' });
    });

    // Inject Office mock before scripts run
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          // Simulate Office environment ready
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
                    getSliceAsync: (index, cb) => {
                      setTimeout(() => {
                        cb({
                          status: 'Succeeded',
                          value: { data: new Uint8Array(50) }
                        });
                      }, 500); // Add delay to catch intermediate status
                    },
                    closeAsync: (cb) => {
                      if (cb) cb();
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
  });

  test('should display initials JV after initialization', async ({ page }) => {
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should update status message when reading active files', async ({ page }) => {
    // Wait for initialization
    await expect(page.locator('#initials-display')).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await readBtn.click();

    // 1. Initial "Reading active file(s)..."
    await expect(status).toHaveText('Reading active file(s)...');

    // 2. File size and slice count info
    await expect(status).toHaveText(/File size: 100 bytes. Reading 2 slices.../);

    // 3. Progress updates
    await expect(status).toHaveText('Reading progress: 50%');
    await expect(status).toHaveText('Reading progress: 100%');

    // 4. Success message
    await expect(status).toHaveText('Successfully read active file(s): 100 bytes.');
  });

  test('button should be centered in its container', async ({ page }) => {
    const container = page.locator('.button-container');
    const button = page.locator('#read-files-btn');

    const containerStyle = await container.evaluate(el => getComputedStyle(el).marginTop);
    expect(containerStyle).toBe('30px');

    const buttonMargin = await button.evaluate(el => {
      const style = getComputedStyle(el);
      return {
        marginLeft: style.marginLeft,
        marginRight: style.marginRight,
        marginTop: style.marginTop
      };
    });

    expect(buttonMargin.marginTop).toBe('0px');
    // For centered elements with margin: 0 auto, left and right margins should be roughly equal and non-zero
    const ml = parseFloat(buttonMargin.marginLeft);
    const mr = parseFloat(buttonMargin.marginRight);
    expect(ml).toBeGreaterThan(0);
    expect(Math.abs(ml - mr)).toBeLessThan(1);
  });
});
