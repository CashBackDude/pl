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

// Make the functions globally accessible so Outlook can find them
if (typeof window !== "undefined") {
  window.openForm = openForm;     // for you / testing
  window.launchForm = openForm;   // for Outlook, matches <FunctionName>
}
