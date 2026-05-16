const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // Mock Office.js before the page loads
    await page.addInitScript(() => {
      window.Office = {
        onReady: (callback) => {
          // Simulate PowerPoint host with a slight delay to mimic real behavior
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
          }, 100);
        },
        context: {
          document: {
            getFileAsync: (fileType, options, callback) => {
              // Simulate successful file retrieval
              callback({
                status: 'succeeded',
                value: {
                  size: 100,
                  sliceCount: 1,
                  getSliceAsync: (index, sliceCallback) => {
                    sliceCallback({
                      status: 'succeeded',
                      value: {
                        data: new Uint8Array(100)
                      }
                    });
                  },
                  closeAsync: (closeCallback) => {
                    closeCallback();
                  }
                }
              });
            }
          }
        },
        FileType: {
          Compressed: 'compressed'
        },
        AsyncResultStatus: {
          Succeeded: 'succeeded'
        },
        HostType: {
          PowerPoint: 'PowerPoint'
        }
      };
    });

    // Intercept the real office.js call to prevent it from loading and overriding our mock
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', route => {
      route.fulfill({
        status: 200,
        contentType: 'application/javascript',
        body: '// Mock Office.js'
      });
    });

    await page.goto('http://localhost:3000/index.html');
  });

  test('should display the correct button text and be wrapped in a container', async ({ page }) => {
    const buttonContainer = page.locator('.button-container');
    await expect(buttonContainer).toBeVisible();

    const readFilesBtn = page.locator('#read-files-btn');
    await expect(readFilesBtn).toBeVisible();
    await expect(readFilesBtn).toHaveText('Read Active Files');

    // Verify CSS for button container
    const marginTop = await buttonContainer.evaluate(el => window.getComputedStyle(el).marginTop);
    expect(marginTop).toBe('30px');

    // Verify button is centered (margin: 0 auto)
    const marginLR = await readFilesBtn.evaluate(el => {
      const style = window.getComputedStyle(el);
      return { left: style.marginLeft, right: style.marginRight };
    });
    // For block elements with margin: 0 auto, left and right margins should be equal and non-zero if centered
    expect(parseFloat(marginLR.left)).toBeGreaterThan(0);
    expect(Math.abs(parseFloat(marginLR.left) - parseFloat(marginLR.right))).toBeLessThan(1);
  });

  test('should read active files and update status on button click', async ({ page }) => {
    // Wait for initials to be set, indicating Office.onReady has fired
    const initialsDisplay = page.locator('#initials-display');
    await expect(initialsDisplay).toHaveText('JV');

    const readFilesBtn = page.locator('#read-files-btn');
    const status = page.locator('#status');

    await readFilesBtn.click();

    // Check for success message
    await expect(status).toHaveText(/Successfully read active file\(s\): 100 bytes\./);
  });

  test('should display initials JV', async ({ page }) => {
    const initialsDisplay = page.locator('#initials-display');
    await expect(initialsDisplay).toHaveText('JV');
  });
});
