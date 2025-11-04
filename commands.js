// Copyright (c) Microsoft Corporation. Licensed under the MIT license.

Office.onReady(() => {
  // Office.js is ready
});

function openForm(event) {
  try {
    window.open("https://forms.office.com/e/0WMwRUR02J");
  } finally {
    event.completed();
  }
}

// Make the function globally accessible to Outlook
if (typeof window !== "undefined") {
  window.openForm = openForm;  // Must match <FunctionName> in manifest
}
