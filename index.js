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

    // Set up the button click handler
    const readFilesBtn = document.getElementById("read-files-btn");
    if (readFilesBtn) {
      readFilesBtn.onclick = readFile;
    }
  }
});

/**
 * Reads the active presentation file using the Office JavaScript API.
 */
async function readFile() {
  const status = document.getElementById("status");
  status.textContent = "Reading active file...";

  try {
    const file = await getFilePromise(Office.FileType.Compressed, { sliceSize: 65536 });
    const sliceCount = file.sliceCount;
    status.textContent = `File loaded. Total slices: ${sliceCount}. Processing...`;

    const dataChunks = [];
    let totalSize = 0;

    for (let i = 0; i < sliceCount; i++) {
      status.textContent = `Reading slice ${i + 1} of ${sliceCount}...`;
      const slice = await getSlicePromise(file, i);
      dataChunks.push(slice.data);
      totalSize += slice.data.length;
    }

    status.textContent = `Successfully read active presentation file (${totalSize} bytes).`;
    console.log(`Read ${totalSize} bytes from the file across ${sliceCount} slices.`);

    await closeFilePromise(file);
  } catch (error) {
    status.textContent = "Error reading active file: " + error.message;
    console.error(error);
  }
}

/**
 * Promisified version of Office.context.document.getFileAsync
 */
function getFilePromise(fileType, options) {
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
function getSlicePromise(file, sliceIndex) {
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
 * Promisified version of file.closeAsync
 */
function closeFilePromise(file) {
  return new Promise((resolve) => {
    file.closeAsync(() => {
      resolve();
    });
  });
}
