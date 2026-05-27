const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery UI', () => {
    test.beforeEach(async ({ page }) => {
        // Mock Office.js library
        await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', async (route) => {
            await route.fulfill({ body: '// Mock Office.js' });
        });

        // Inject Office mock object before the page loads
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
                                            }, 200);
                                        },
                                        closeAsync: (closeCallback) => {
                                            if (closeCallback) closeCallback();
                                        }
                                    }
                                });
                            }, 200);
                        }
                    }
                }
            };
        });

        await page.goto('http://localhost:3000/index.html');
    });

    test('should initialize with correct initials and button text', async ({ page }) => {
        // Check for the initials "JV"
        await expect(page.locator('#initials-display')).toHaveText('JV');

        // Check for the pluralized button text
        await expect(page.locator('#read-files-btn')).toHaveText('Read Active Files');
    });

    test('should show progress and success message when reading file', async ({ page }) => {
        // Wait for initialization
        await expect(page.locator('#initials-display')).toHaveText('JV');

        // Click the "Read Active Files" button
        await page.click('#read-files-btn');

        // Verify initial status
        await expect(page.locator('#status')).toHaveText('Reading active file(s)...');

        // Verify progress update (50% after first slice)
        await expect(page.locator('#status')).toHaveText('Reading progress: 50%', { timeout: 1000 });

        // Verify final success message
        await expect(page.locator('#status')).toHaveText('Successfully read active file(s): 100 bytes.', { timeout: 2000 });
    });
});
