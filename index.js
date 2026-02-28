/**
 * Eco-growth Discovery Project
 * Entry point for the Node.js project.
 */

// Promisified Office.context.document.getFileAsync
const getFileAsync = (fileType, options) => {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(fileType, options, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
};

// Promisified file.getSliceAsync
const getSliceAsync = (file, sliceIndex) => {
  return new Promise((resolve, reject) => {
    file.getSliceAsync(sliceIndex, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
};

async function readActiveFile() {
  const statusDisplay = document.getElementById("status");
  if (!statusDisplay) return;

  if (typeof Office === "undefined" || !Office.context || !Office.context.document) {
    statusDisplay.textContent = "Error: Office context not available. This add-in must run within an Office application.";
    return;
  }

  statusDisplay.textContent = "Reading file...";

  let file = null;
  try {
    file = await getFileAsync(Office.FileType.Compressed);
    const sliceCount = file.sliceCount;
    const fileSize = file.size;
    const aggregatedData = new Uint8Array(fileSize);
    let offset = 0;

    for (let i = 0; i < sliceCount; i++) {
      const slice = await getSliceAsync(file, i);
      aggregatedData.set(slice.data, offset);
      offset += slice.data.length;
      statusDisplay.textContent = `Reading file: ${Math.round(((i + 1) / sliceCount) * 100)}%`;
    }

    statusDisplay.textContent = `Successfully read file (${fileSize} bytes).`;
    console.log("File data length:", aggregatedData.length);
  } catch (error) {
    console.error("Error reading file:", error);
    statusDisplay.textContent = `Error reading file: ${error.message}`;
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

    // Hook up button
    const readFilesBtn = document.getElementById("read-files-btn");
    if (readFilesBtn) {
      readFilesBtn.addEventListener("click", readActiveFile);
    }
  }
});
