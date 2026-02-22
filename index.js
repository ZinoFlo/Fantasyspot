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

    // "Read Active Files" functionality
    const readFilesBtn = document.getElementById("read-files-btn");
    const status = document.getElementById("status");
    if (readFilesBtn) {
      readFilesBtn.onclick = () => {
        status.textContent = "Reading presentation...";

        // In PowerPoint, we read the document content.
        // We use getFileAsync to get the compressed (pptx) content.
        Office.context.document.getFileAsync(Office.FileType.Compressed, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const file = result.value;
            status.textContent = `Success! Presentation read. Total size: ${file.size} bytes.`;
            file.closeAsync();
          } else {
            status.textContent = `Error: ${result.error.message}`;
            console.error(result.error);
          }
        });
      };
    }
  }
});
