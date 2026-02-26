/**
 * Eco-growth Discovery Project
 * Entry point for the Node.js project.
 */

/**
 * Promisified wrapper for getFileAsync.
 */
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

/**
 * Promisified wrapper for getSliceAsync.
 */
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

/**
 * Reads the active presentation's content and aggregates it into a single data array.
 */
async function readActiveFile() {
  const statusDisplay = document.getElementById("status");

  // Check if Office.js is available and initialized
  if (typeof Office === "undefined" || !Office.context || !Office.context.document) {
    if (statusDisplay) {
      statusDisplay.textContent = "Error: Office.js is not available. This action must be performed within PowerPoint.";
    }
    return;
  }

  if (statusDisplay) statusDisplay.textContent = "Opening file...";

  let file = null;
  try {
    // Request the file in compressed format
    const fileType = Office.FileType.Compressed;
    file = await getFileAsync(fileType);

    const sliceCount = file.sliceCount;
    const totalSize = file.size;

    if (statusDisplay) {
      statusDisplay.textContent = `Reading ${sliceCount} slices (${totalSize} bytes)...`;
    }

    // Use Uint8Array for efficient binary data aggregation
    const docData = new Uint8Array(totalSize);
    let offset = 0;

    for (let i = 0; i < sliceCount; i++) {
      const slice = await getSliceAsync(file, i);
      // Aggregate binary data from each slice
      docData.set(slice.data, offset);
      offset += slice.data.length;

      if (statusDisplay) {
        statusDisplay.textContent = `Progress: ${Math.round(((i + 1) / sliceCount) * 100)}%`;
      }
    }

    if (statusDisplay) {
      statusDisplay.textContent = `Successfully read active file (${docData.length} bytes).`;
    }
    console.log("File content aggregated:", docData);
    return docData;
  } catch (error) {
    if (statusDisplay) {
      statusDisplay.textContent = `Error reading file: ${error.message || error}`;
    }
    console.error("Error reading file:", error);
    throw error;
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

    // Attach event listener to the "Read Active Files" button
    const readFilesBtn = document.getElementById("read-files-btn");
    if (readFilesBtn) {
      readFilesBtn.addEventListener("click", () => {
        readActiveFile().catch((err) => {
          console.error("Failed to read active file:", err);
        });
      });
    }
  }
});
