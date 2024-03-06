/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

function toggleTaskPane(event) {
  // Your logic to toggle the taskpane
  // This is often handled through the manifest rather than code.

  // Placeholder for any additional action
  console.log("Toggle taskpane action invoked");

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("toggleTaskPane", toggleTaskPane);

