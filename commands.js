// Copyright (c) Microsoft Corporation. Licensed under the MIT license.

Office.onReady(() => {
  // Office.js is ready
});

/**
 * Command handler for the ribbon button.
 * Must match <FunctionName>launchForm</FunctionName> in the manifest.
 */
function launchForm(event) {
  try {
    window.open("https://forms.office.com/e/0WMwRUR02J");
  } finally {
    // Always tell Outlook we're done, even if window.open fails
    event.completed();
  }
}

// Make the function globally accessible so Outlook can find it
if (typeof window !== "undefined") {
  window.launchForm = launchForm;
}
