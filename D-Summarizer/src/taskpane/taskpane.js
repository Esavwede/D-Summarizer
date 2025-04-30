/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    runAgent();

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}

const GEMINI_API_KEY = "YOUR_GEMINI_API_KEY_HERE";

async function summarizeWithGemini(text) {
  const response = await fetch(
    "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=" +
      GEMINI_API_KEY,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        contents: [{ parts: [{ text: `Summarize this: ${text}` }] }],
      }),
    }
  );

  const data = await response.json();
  return data?.candidates?.[0]?.content?.parts?.[0]?.text || "No summary returned.";
}

async function runAgent() {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const textToSummarize = body.text.slice(0, 2000); // limit input size
    const summary = await summarizeWithGemini(textToSummarize);

    body.insertParagraph("Summary:", Word.InsertLocation.end);
    body.insertParagraph(summary, Word.InsertLocation.end);
    await context.sync();
  });
}
