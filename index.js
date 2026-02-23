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

    // Read Active Files functionality
    const readFilesBtn = document.getElementById("read-files-btn");
    const status = document.getElementById("status");

    if (readFilesBtn) {
      readFilesBtn.onclick = () => {
        if (status) status.textContent = "Reading file...";

        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const myFile = result.value;
            const sliceCount = myFile.sliceCount;
            let slicesRead = 0;

            if (status) status.textContent = `File loaded. Reading ${sliceCount} slices...`;

            const getSlice = (index) => {
              myFile.getSliceAsync(index, (sliceResult) => {
                if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                  slicesRead++;
                  if (status) status.textContent = `Read slice ${slicesRead} of ${sliceCount}...`;

                  if (slicesRead < sliceCount) {
                    getSlice(slicesRead);
                  } else {
                    myFile.closeAsync();
                    if (status) status.textContent = `Active file read successfully (${sliceCount} slices).`;
                  }
                } else {
                  myFile.closeAsync();
                  if (status) status.textContent = "Error reading slice: " + sliceResult.error.message;
                }
              });
            };

            getSlice(0);
          } else {
            if (status) status.textContent = "Error reading active file: " + result.error.message;
          }
        });
      };
    }
  }
});
