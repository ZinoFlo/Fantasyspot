/**
 * Eco-growth Discovery Project
 * Entry point for the Node.js project.
 */

Office.onReady((info) => {
  // Initialize for PowerPoint or when testing in a browser
  if (info.host === Office.HostType.PowerPoint || !info.host) {
    console.log("Eco-growth Discovery initialized.");

    // Retrieve and display initials
    const initialsDisplay = document.getElementById("initials-display");
    if (initialsDisplay) {
      // Default initials from the project template metadata (Julien Vink)
      let initials = "JV";

      // Note: Office.context.userProfile is primarily for Outlook.
      // In PowerPoint, user information is not directly exposed via simple properties.
      // We use "JV" as the primary identifier for this specialized add-in.

      initialsDisplay.textContent = initials;
    }

    const readBtn = document.getElementById("read-files-btn");
    if (readBtn) {
      readBtn.onclick = readActiveFile;
    }
  }
});

/**
 * Reads the active PowerPoint file as a byte stream.
 */
async function readActiveFile() {
  const status = document.getElementById("status");
  if (status) status.textContent = "Reading file...";

  if (typeof Office === "undefined" || !Office.context || !Office.context.document) {
    const errorMsg = "Office environment not detected. This function only works within a PowerPoint Add-in.";
    console.error(errorMsg);
    if (status) status.textContent = errorMsg;
    return;
  }

  let file;
  try {
    file = await new Promise((resolve, reject) => {
      Office.context.document.getFileAsync(Office.FileType.Compressed, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(result.error);
        }
      });
    });

    const sliceCount = file.sliceCount;
    const fileSize = file.size;
    let fileContent = new Uint8Array(fileSize);
    let offset = 0;

    for (let i = 0; i < sliceCount; i++) {
      const slice = await new Promise((resolve, reject) => {
        file.getSliceAsync(i, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value);
          } else {
            reject(result.error);
          }
        });
      });
      fileContent.set(slice.data, offset);
      offset += slice.data.length;
    }

    if (status) status.textContent = `Successfully read ${fileSize} bytes from the active presentation.`;
    console.log("File content read successfully:", fileContent);
  } catch (error) {
    const errorMsg = "Error reading file: " + (error.message || error);
    console.error(errorMsg);
    if (status) status.textContent = errorMsg;
  } finally {
    if (file) {
      await new Promise((resolve) => {
        file.closeAsync(() => {
          resolve();
        });
      });
    }
  }
}
