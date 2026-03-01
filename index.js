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

      // Note: Office.context.userProfile is primarily for Office hosts like Outlook.
      // In PowerPoint, user information is not directly exposed via simple properties.
      // We use "JV" as the primary identifier for this specialized add-in.

      initialsDisplay.textContent = initials;
    }

    // Attach event listener to the button
    const readFilesBtn = document.getElementById("read-files-btn");
    if (readFilesBtn) {
      readFilesBtn.onclick = readActiveFile;
    }
  }
});

/**
 * Reads the content of the active PowerPoint file.
 */
async function readActiveFile() {
  const status = document.getElementById("status");
  if (!status) return;

  if (typeof Office === "undefined" || !Office.context || !Office.context.document) {
    status.innerText = "Error: Office.js is not initialized or not supported in this environment.";
    status.style.color = "red";
    return;
  }

  status.innerText = "Reading file...";
  status.style.color = "black";

  let file = null;

  try {
    file = await getFileAsync(Office.FileType.Compressed);
    const totalSlices = file.sliceCount;
    let fileContent = new Uint8Array(file.size);
    let offset = 0;

    for (let i = 0; i < totalSlices; i++) {
      const slice = await getSliceAsync(file, i);
      fileContent.set(slice.data, offset);
      offset += slice.data.length;
      status.innerText = `Reading file... ${Math.round(((i + 1) / totalSlices) * 100)}%`;
    }

    status.innerText = `Successfully read active file (${file.size} bytes).`;
    status.style.color = "green";
  } catch (error) {
    status.innerText = "Error reading file: " + (error.message || error);
    status.style.color = "red";
    console.error(error);
  } finally {
    if (file) {
      file.closeAsync();
    }
  }
}

/**
 * Promisified version of Office.context.document.getFileAsync.
 */
function getFileAsync(fileType) {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(fileType, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

/**
 * Promisified version of file.getSliceAsync.
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
