/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { summarize } from "../services/summarizer-api.service";

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("summarize").onclick = run;
  }
});

export async function run() {
  document.getElementById("loading-overlay").classList.remove("hidden");
  return Word.run(async (context) => {
    try {
      const selection = context.document.getSelection();

      selection.load("text");

      await context.sync();

      const selectionContent = selection.text;

      if (selectionContent.length === 0) {
        throw new Error(`Text selection length must be greater than 0, please increase your selection length.`);
      }

      if (selectionContent.split(/\W/).length > 1024) {
        throw new Error(`Our model has a limit of 1024 tokens, please reduce your selection length.`);
      }

      selection.insertBreak("Line", "After");

      const { summary } = await summarize("https://api.legal-summarizer.students-epitech.ovh", selectionContent).catch(
        (error) => {
          console.log(error);
          throw new Error(`An unexpected error occured, the AI service seems unavailable, please try again later.`);
        }
      );

      if (!summary.length) {
        throw new Error(`Invalid output summary format.`);
      }

      const paragraph = selection.insertParagraph(summary.join(" "), "After");
      paragraph.font.color = "red";

      document.getElementById("loading-overlay").classList.add("hidden");

      await context.sync();
    } catch (error) {
      console.log(error);
      document.getElementById("loading-overlay").classList.add("hidden");
      document.getElementById("alert").classList.remove("hidden");
      document.getElementById("alert").innerText = error.message;
      let timeout = null;
      timeout = setTimeout(() => {
        document.getElementById("alert").classList.add("hidden");
        timeout && clearTimeout(timeout);
      }, 5500);
    }
  });
}
