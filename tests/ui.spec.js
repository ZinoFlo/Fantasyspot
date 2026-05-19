const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Capture console logs from the browser
    page.on('console', msg => console.log(`BROWSER [${msg.type()}]: ${msg.text()}`));

    // Mock Office.js to prevent external network requests and provide a controlled environment
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', route => {
      route.fulfill({ body: '// Mock Office.js' });
    });

    // Inject Office mock before the page scripts run
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          // Simulate Office environment readiness with a small delay
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
          }, 100);
        },
        HostType: { PowerPoint: 'PowerPoint' },
        FileType: { Compressed: 'Compressed' },
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        context: {
          document: {
            getFileAsync: (fileType, options, callback) => {
              // Simulate getFileAsync success
              setTimeout(() => {
                callback({
                  status: 'succeeded',
                  value: {
                    size: 1024,
                    sliceCount: 2,
                    getSliceAsync: (index, sliceCallback) => {
                      setTimeout(() => {
                        sliceCallback({
                          status: 'succeeded',
                          value: { data: new Uint8Array(512) }
                        });
                      }, 50);
                    },
                    closeAsync: (closeCallback) => {
                      setTimeout(() => closeCallback(), 10);
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

  test('should display correct title and initials', async ({ page }) => {
    await expect(page).toHaveTitle('Eco-growth Discovery');
    await expect(page.locator('h1')).toHaveText('Eco-growth Discovery');

    // Wait for Office.onReady to update initials
    const initials = page.locator('#initials-display');
    await expect(initials).toHaveText('JV');
  });

  test('should have a centered "Read Active Files" button with correct styling', async ({ page }) => {
    const button = page.locator('#read-files-btn');
    await expect(button).toBeVisible();
    await expect(button).toHaveText('Read Active Files');

    // Verify background color (#217346)
    const bgColor = await button.evaluate((el) => window.getComputedStyle(el).backgroundColor);
    // rgb(33, 115, 70) is #217346
    expect(bgColor).toBe('rgb(33, 115, 70)');

    // Verify centering
    const margin = await button.evaluate((el) => {
      const style = window.getComputedStyle(el);
      return {
        marginLeft: style.marginLeft,
        marginRight: style.marginRight,
        marginTop: style.marginTop
      };
    });

    expect(margin.marginTop).toBe('0px');
    const ml = parseFloat(margin.marginLeft);
    const mr = parseFloat(margin.marginRight);
    expect(ml).toBeGreaterThan(0);
    expect(Math.abs(ml - mr)).toBeLessThan(1); // Account for sub-pixel rendering
  });

  test('should show success status message when reading files', async ({ page }) => {
    const button = page.locator('#read-files-btn');
    const status = page.locator('#status');

    // Add a simple log to see if the button is clicked and if Office is initialized
    await page.evaluate(() => console.log("Button visible:", !!document.getElementById('read-files-btn')));

    await button.click();

    // Verify final success message
    await expect(status).toHaveText('Successfully read active file(s): 1024 bytes.', { timeout: 15000 });
  });
});
