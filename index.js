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
      readFilesBtn.onclick = readActiveFile;
    }
  }
});

/**
 * Promisified version of Office.context.document.getFileAsync
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
 * Promisified version of file.getSliceAsync
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
 * Reads the active PowerPoint file as a compressed byte stream.
 */
async function readActiveFile() {
  const status = document.getElementById("status");
  if (!status) return;

  if (typeof Office === "undefined" || !Office.context || !Office.context.document) {
    status.textContent = "Error: Office.context.document is not available. Please ensure you are running this in PowerPoint.";
    return;
  }

  status.textContent = "Reading file...";

  let file = null;
  try {
    // Read the file as Compressed (entire byte stream)
    file = await getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 });
    const sliceCount = file.sliceCount;
    let fileContent = new Uint8Array(file.size);
    let offset = 0;

    for (let i = 0; i < sliceCount; i++) {
      const slice = await getSliceAsync(file, i);
      fileContent.set(slice.data, offset);
      offset += slice.data.length;
    }

    status.textContent = `Success! File read. Total size: ${file.size} bytes.`;
    console.log("File content read successfully.", fileContent);
  } catch (error) {
    status.textContent = `Error reading file: ${error.message || error}`;
    console.error(error);
  } finally {
    if (file) {
      file.closeAsync();
    }
  }
}
