/**
 * Eco-growth Discovery Project
 * Entry point for the Node.js project.
 */

// Promisified Office.context.document.getFileAsync
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

// Promisified file.getSliceAsync
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
 * Reads the active presentation file and displays its size.
 */
async function readActiveFile() {
  const status = document.getElementById("status");
  if (!status) return;

  if (!Office.context || !Office.context.document) {
    status.textContent = "Error: Office.context.document is not available. Are you running in a supported Office host?";
    return;
  }

  status.textContent = "Reading file...";

  let file;
  try {
    // Read the file in 64KB slices
    file = await getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 });
    const sliceCount = file.sliceCount;
    const fileSize = file.size;
    const data = new Uint8Array(fileSize);
    let offset = 0;

    for (let i = 0; i < sliceCount; i++) {
      const slice = await getSliceAsync(file, i);
      data.set(slice.data, offset);
      offset += slice.data.length;
    }

    status.textContent = `Successfully read file. Size: ${fileSize} bytes.`;
    console.log(`File read successfully. Total bytes: ${fileSize}`);
  } catch (error) {
    status.textContent = `Error reading file: ${error.message}`;
    console.error("Error reading file:", error);
  } finally {
    if (file) {
      file.closeAsync();
    }
  }
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
      readFilesBtn.addEventListener("click", readActiveFile);
    }
  }
});
