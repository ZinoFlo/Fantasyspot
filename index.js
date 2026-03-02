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

    const readFilesBtn = document.getElementById("read-files-btn");
    if (readFilesBtn) {
      readFilesBtn.addEventListener("click", readActiveFile);
    }
  }
});

/**
 * Promisified Office.context.document.getFileAsync
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
 * Promisified file.getSliceAsync
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
 * Promisified file.closeAsync
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

/**
 * Reads the active PowerPoint file as a compressed byte stream.
 */
async function readActiveFile() {
  const status = document.getElementById("status");
  if (!status) return;

  if (!Office.context.document) {
    status.textContent = "Error: Office.context.document is not available.";
    return;
  }

  status.textContent = "Reading file...";

  try {
    const file = await getFileAsync(Office.FileType.Compressed);
    const sliceCount = file.sliceCount;
    let docData = new Uint8Array(file.size);
    let offset = 0;

    try {
      for (let i = 0; i < sliceCount; i++) {
        const slice = await getSliceAsync(file, i);
        docData.set(slice.data, offset);
        offset += slice.data.length;
        status.textContent = `Reading slice ${i + 1} of ${sliceCount}...`;
      }
      status.textContent = `File read successfully. Size: ${file.size} bytes.`;
      console.log("File data:", docData);
    } finally {
      await closeAsync(file);
    }
  } catch (error) {
    console.error(error);
    status.textContent = "Error reading file: " + error.message;
  }
}
