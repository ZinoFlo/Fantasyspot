/**
 * Eco-growth Discovery Project
 * Entry point for the Node.js project.
 */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    console.log("Eco-growth Discovery initialized in PowerPoint.");
  }
});
