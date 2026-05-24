import { test, expect } from '@playwright/test';

test.beforeEach(async ({ page }) => {
  page.on('console', msg => console.log(`BROWSER [${msg.type()}]: ${msg.text()}`));

  // Mock Office.js
  await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async route => {
    await route.fulfill({ body: '// Mock Office.js' });
  });

  await page.addInitScript(() => {
    window.Office = {
      onReady: (callback) => {
        setTimeout(() => {
          callback({ host: 'PowerPoint' });
        }, 100);
      },
      HostType: { PowerPoint: 'PowerPoint' },
      FileType: { Compressed: 'compressed' },
      AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
      context: {
        document: {
          getFileAsync: (fileType, options, callback) => {
            setTimeout(() => {
              callback({
                status: 'succeeded',
                value: {
                  size: 100,
                  sliceCount: 2,
                  getSliceAsync: (index, sliceCallback) => {
                    setTimeout(() => {
                      sliceCallback({
                        status: 'succeeded',
                        value: { data: new Array(50).fill(0) }
                      });
                    }, 50);
                  },
                  closeAsync: (closeCallback) => {
                    setTimeout(() => {
                      closeCallback({ status: 'succeeded' });
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
});

test('UI elements are present and correctly labeled', async ({ page }) => {
  await page.goto('/');

  // Check title
  await expect(page.locator('h1')).toHaveText('Eco-growth Discovery');

  // Check button text
  const readBtn = page.locator('#read-files-btn');
  await expect(readBtn).toHaveText('Read Active Files');

  // Check initials display
  const initials = page.locator('#initials-display');
  await expect(initials).toHaveText('JV');
});

test('Button container has correct styling', async ({ page }) => {
  await page.goto('/');

  const container = page.locator('.button-container');
  await expect(container).toHaveCSS('margin-top', '30px');

  const readBtn = page.locator('#read-files-btn');
  const marginLeft = await readBtn.evaluate(el => window.getComputedStyle(el).marginLeft);
  const marginRight = await readBtn.evaluate(el => window.getComputedStyle(el).marginRight);

  // Verify it's centered (margin 0 auto results in equal left/right margins)
  const ml = parseFloat(marginLeft);
  const mr = parseFloat(marginRight);
  expect(Math.abs(ml - mr)).toBeLessThan(1);
  expect(ml).toBeGreaterThan(0);
});

test('Reading files updates status correctly', async ({ page }) => {
  await page.goto('/');

  // Wait for initials to ensure Office.onReady has completed and listeners are attached
  const initials = page.locator('#initials-display');
  await expect(initials).toHaveText('JV');

  const readBtn = page.locator('#read-files-btn');
  await readBtn.click();

  const status = page.locator('#status');

  // Wait for any of the expected status messages to appear
  // This helps avoid flakiness if the first message is transient
  await expect(status).not.toHaveText('', { timeout: 10000 });

  const text = await status.textContent();
  console.log('Final status text observed:', text);

  // We check that the final state is reached and pluralization is correct
  await expect(status).toHaveText('Successfully read active file(s): 100 bytes.');
});
