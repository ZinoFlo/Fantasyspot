const { test, expect } = require("@playwright/test");

test.describe("Eco-growth Discovery Taskpane UI Tests", () => {
  test.beforeEach(async ({ page }) => {
    // Intercept and mock office.js load requests to return a dummy script.
    // This allows our injected `Office` global via `addInitScript` to take precedence without external fetch.
    await page.route("https://appsforoffice.microsoft.com/lib/1/hosted/office.js", async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "application/javascript",
        body: "// Mock Office.js library loaded",
      });
    });

    // Mock Office object before page scripts run.
    await page.addInitScript(() => {
      window.Office = {
        HostType: {
          PowerPoint: "PowerPoint",
        },
        FileType: {
          Compressed: "compressed",
        },
        AsyncResultStatus: {
          Succeeded: "succeeded",
          Failed: "failed",
        },
        onReady: (callback) => {
          // Add a slight delay to allow scripts to load and attach event handlers.
          setTimeout(() => {
            callback({ host: "PowerPoint" });
          }, 100);
        },
        context: {
          document: {
            getFileAsync: (fileType, options, callback) => {
              // Introduce a delay to capture and assert intermediate UI progress messages.
              setTimeout(() => {
                callback({
                  status: "succeeded",
                  value: {
                    sliceCount: 2,
                    size: 10,
                    getSliceAsync: (sliceIndex, sliceCallback) => {
                      setTimeout(() => {
                        sliceCallback({
                          status: "succeeded",
                          value: {
                            data: new Uint8Array([1, 2, 3, 4, 5]),
                          },
                        });
                      }, 100);
                    },
                    closeAsync: (closeCallback) => {
                      setTimeout(() => {
                        closeCallback();
                      }, 50);
                    },
                  },
                });
              }, 100);
            },
          },
        },
      };
    });

    await page.goto("/");
  });

  test("initializes taskpane correctly and updates initials display to JV", async ({ page }) => {
    // Verify initial text before Office JS is fully ready (if any) or wait for transition.
    const initials = page.locator("#initials-display");
    await expect(initials).toHaveText("JV");
  });

  test("successfully reads active file(s) and displays status messages sequentially", async ({ page }) => {
    // Ensure the taskpane has initialized first.
    await expect(page.locator("#initials-display")).toHaveText("JV");

    const status = page.locator("#status");
    const readBtn = page.locator("#read-files-btn");

    // Click the Read Active File(s) button.
    await readBtn.click();

    // Assert intermediate status message is displayed during start.
    await expect(status).toHaveText(/Reading active file\(s\)\.\.\./);

    // Verify intermediate progress updates to 50% or 100% or final success due to step sequence.
    await expect(status).toHaveText(/(Reading progress: (50|100)%|Successfully read active file\(s\): 10 bytes\.)/);

    // Assert the final successful read message.
    await expect(status).toHaveText(/Successfully read active file\(s\): 10 bytes\./);
  });
});
