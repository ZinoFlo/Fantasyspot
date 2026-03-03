/**
 * Eco-growth Discovery Project
 * Entry point for the Node.js project.
 */

/**
 * Promisified wrapper for Office.context.document.getFileAsync
 */
function getFileAsync(fileType, options) {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(fileType, options, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
}

/**
 * Promisified wrapper for file.getSliceAsync
 */
function getSliceAsync(file, sliceIndex) {
  return new Promise((resolve, reject) => {
    file.getSliceAsync(sliceIndex, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
}

/**
 * Promisified wrapper for file.closeAsync
 */
function closeAsync(file) {
  return new Promise((resolve, reject) => {
    file.closeAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(result.error);
      }
    });
  });
}

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

    // Bind event listener for reading files
    const readFilesBtn = document.getElementById("read-files-btn");
    if (readFilesBtn) {
      readFilesBtn.onclick = readActiveFile;
    }
  }
});

/**
 * Reads the active PowerPoint file as a byte stream.
 */
async function readActiveFile() {
  const status = document.getElementById("status");
  if (!status) return;

  if (typeof Office === "undefined" || !Office.context || !Office.context.document) {
    status.innerText = "Office context not available. This function requires running within PowerPoint.";
    return;
  }

  status.innerText = "Reading file...";

  let file = null;
  try {
    // Office.FileType.Compressed returns the entire presentation as a .pptx file byte stream.
    file = await getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 });
    const sliceCount = file.sliceCount;
    const fileData = new Uint8Array(file.size);
    let offset = 0;

    for (let i = 0; i < sliceCount; i++) {
      const slice = await getSliceAsync(file, i);
      fileData.set(slice.data, offset);
      offset += slice.data.length;
      status.innerText = `Reading slice ${i + 1} of ${sliceCount}...`;
    }

    status.innerText = `Successfully read active file: ${file.size} bytes.`;
    console.log("File content aggregated.", fileData);
  } catch (error) {
    console.error("Error reading file:", error);
    status.innerText = `Error reading file: ${error.message || error}`;
  } finally {
    if (file) {
      try {
        await closeAsync(file);
      } catch (closeError) {
        console.error("Error closing file:", closeError);
      }
    }
  }
}
