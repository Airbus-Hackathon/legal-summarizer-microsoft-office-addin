/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("summarize").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {
    // Get the body of the current document
    const body = context.document.body;

    // Load the body and retrieve its text
    body.load("text");
    // Execute the request
    await context.sync();

    // Store the text content in a variable
    const documentContent = body.text;

    // Create a new document
    const newDoc = context.application.createDocument();

    const newBody = newDoc.body;

    // Insert the copied content into the new document
    newBody.insertText(documentContent, Word.InsertLocation.start);

    // Open the new document
    newDoc.open();

    await context.sync();
  });
}
