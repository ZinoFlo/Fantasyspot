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

    // Attach event listener to the button
    const readFilesBtn = document.getElementById("read-files-btn");
    if (readFilesBtn) {
      readFilesBtn.onclick = readActiveFile;
    }
  } else {
    const status = document.getElementById("status");
    if (status) {
      status.textContent = "This add-in is only supported in PowerPoint.";
    }
  }
});

/**
 * Reads the active PowerPoint file.
 */
async function readActiveFile() {
  const status = document.getElementById("status");
  if (!status) return;

  if (typeof Office === "undefined" || !Office.context || !Office.context.document) {
    status.textContent = "Office context is not available.";
    return;
  }

  status.textContent = "Reading file...";

  try {
    const file = await getFileAsync(Office.FileType.Compressed);
    const fileData = new Uint8Array(file.size);
    let offset = 0;

    try {
      for (let i = 0; i < file.sliceCount; i++) {
        const slice = await getSliceAsync(file, i);
        fileData.set(slice.value, offset);
        offset += slice.value.length;
      }
      status.textContent = `Successfully read file. Size: ${file.size} bytes.`;
      console.log("File read successfully.", fileData);
    } finally {
      file.closeAsync();
    }
  } catch (error) {
    status.textContent = `Error reading file: ${error.message}`;
    console.error(error);
  }
}

/**
 * Promisified version of Office.context.document.getFileAsync
 */
function getFileAsync(fileType) {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(fileType, { sliceSize: 65536 }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

/**
 * Promisified version of file.getSliceAsync
 */
function getSliceAsync(file, sliceIndex) {
  return new Promise((resolve, reject) => {
    file.getSliceAsync(sliceIndex, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}
