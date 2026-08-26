

/*************************************************
 * AICoreStack Content Automation
 *
 * MASTER CODE — STEPS 1–6
 *
 * Completed:
 * STEP 3  — Apps Script → Google Sheets
 * STEP 4  — Calendar Reading / READY Content
 * STEP 5  — Today's Content / Approval
 * STEP 6C — Gemini API Connection
 * STEP 6D — Google Sheets → Gemini
 * STEP 6E — Gemini Structured Output
 *
 * CURRENTLY NOT INCLUDED:
 * STEP 7  — Image / Veo Generation
 * STEP 8  — Google Drive Asset Handling
 * STEP 9  — Meta / Instagram API
 * STEP 10 — Automatic Scheduling
 * STEP 11 — Approval + Safety Automation
 * STEP 12 — Daily Trigger
 *************************************************/


const CONFIG = {

  // Google Sheet tab name
  SHEET_NAME: "90 Content Calendar",

  // Header row
  HEADER_ROW: 1,

  // Working Gemini model
  GEMINI_MODEL: "gemini-3.5-flash-lite",

  // Script Properties key
  GEMINI_API_KEY_PROPERTY: "GEMINI_API_KEY"

};


/*************************************************
 * STEP 3
 * TEST GOOGLE SHEETS CONNECTION
 *************************************************/

function testConnection() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {

    throw new Error(
      `Sheet "${CONFIG.SHEET_NAME}" was not found.`
    );

  }

  Logger.log("=================================");
  Logger.log("STEP 3 — SHEET CONNECTION TEST");
  Logger.log("=================================");

  Logger.log(
    "Google Sheet connection successful."
  );

  Logger.log(
    `Sheet: ${sheet.getName()}`
  );

  Logger.log(
    `Rows: ${sheet.getLastRow() - 1}`
  );

  Logger.log(
    `Columns: ${sheet.getLastColumn()}`
  );

}


/*************************************************
 * STEP 4
 * READ ENTIRE CONTENT CALENDAR
 *************************************************/

function readContentCalendar() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {

    throw new Error(
      `Sheet "${CONFIG.SHEET_NAME}" was not found.`
    );

  }

  const data =
    sheet.getDataRange().getValues();

  if (data.length === 0) {

    return [];

  }

  const headers =
    data[CONFIG.HEADER_ROW - 1];

  const rows =
    data
      .slice(CONFIG.HEADER_ROW)
      .map(function(row) {

        const item = {};

        headers.forEach(function(header, index) {

          item[header] =
            row[index];

        });

        return item;

      });

  Logger.log(
    `Total content rows: ${rows.length}`
  );

  return rows;

}


/*************************************************
 * STEP 4
 * FIND READY CONTENT
 *************************************************/

function getReadyContent() {

  const rows =
    readContentCalendar();

  const ready =
    rows.filter(function(row) {

      return (
        String(row["Status"])
          .trim()
          .toUpperCase() === "READY"
      );

    });


  Logger.log("=================================");
  Logger.log("READY CONTENT");
  Logger.log("=================================");

  Logger.log(
    `READY items found: ${ready.length}`
  );


  ready.forEach(function(item, index) {

    Logger.log(
      `#${index + 1} | ` +
      `${item["Account"]} | ` +
      `${item["Content Type"]} | ` +
      `${item["Topic"]} | ` +
      `Status: ${item["Status"]}`
    );

  });


  return ready;

}


/*************************************************
 * STEP 5
 * FIND APPROVED CONTENT
 *************************************************/

function getApprovedContent() {

  const rows =
    readContentCalendar();


  const approved =
    rows.filter(function(row) {

      return (
        String(row["Approval"])
          .trim()
          .toUpperCase() === "YES"
      );

    });


  Logger.log("=================================");
  Logger.log("APPROVED CONTENT");
  Logger.log("=================================");

  Logger.log(
    `Approved items found: ${approved.length}`
  );


  approved.forEach(function(item, index) {

    Logger.log(
      `#${index + 1} | ` +
      `${item["Account"]} | ` +
      `${item["Content Type"]} | ` +
      `${item["Topic"]}`
    );

  });


  return approved;

}


/*************************************************
 * STEP 5
 * FIND TODAY'S CONTENT
 *************************************************/

function getTodaysContent() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEET_NAME);


  if (!sheet) {

    throw new Error(
      `Sheet "${CONFIG.SHEET_NAME}" was not found.`
    );

  }


  const data =
    sheet.getDataRange().getValues();


  const headers =
    data[0];


  const dateIndex =
    headers.indexOf("Date");


  if (dateIndex === -1) {

    throw new Error(
      "Date column was not found."
    );

  }


  const timezone =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSpreadsheetTimeZone();


  const todayString =
    Utilities.formatDate(
      new Date(),
      timezone,
      "yyyy-MM-dd"
    );


  Logger.log("=================================");
  Logger.log("TODAY'S CONTENT");
  Logger.log("=================================");

  Logger.log(
    `Today's date: ${todayString}`
  );

  Logger.log(
    `Spreadsheet timezone: ${timezone}`
  );


  const results = [];


  for (let i = 1; i < data.length; i++) {

    let rowDate =
      data[i][dateIndex];


    if (!rowDate) {
      continue;
    }


    let rowDateString = "";


    if (rowDate instanceof Date) {

      rowDateString =
        Utilities.formatDate(
          rowDate,
          timezone,
          "yyyy-MM-dd"
        );

    } else {

      rowDateString =
        String(rowDate).trim();


      if (
        /^\d{4}-\d{2}-\d{2}/
          .test(rowDateString)
      ) {

        rowDateString =
          rowDateString.substring(
            0,
            10
          );

      }

    }


    if (
      rowDateString ===
      todayString
    ) {

      const item = {};


      headers.forEach(function(header, index) {

        item[header] =
          data[i][index];

      });


      results.push(item);

    }

  }


  Logger.log(
    `Today's content: ${results.length}`
  );


  results.forEach(function(item, index) {

    Logger.log(
      `#${index + 1} | ` +
      `${item["Account"]} | ` +
      `${item["Content Type"]} | ` +
      `${item["Topic"]} | ` +
      `${item["Publish Time"]}`
    );

  });


  return results;

}


/*************************************************
 * HELPER
 * CHECK WHETHER TWO DATES ARE THE SAME
 *************************************************/

function isSameDate(date1, date2) {

  return (

    date1.getFullYear() ===
      date2.getFullYear()

    &&

    date1.getMonth() ===
      date2.getMonth()

    &&

    date1.getDate() ===
      date2.getDate()

  );

}


/*************************************************
 * STEP 6C
 * TEST GEMINI API CONNECTION
 *************************************************/

function testGeminiConnection() {

  Logger.log("=================================");
  Logger.log("STEP 6C — GEMINI API TEST");
  Logger.log("=================================");


  const apiKey =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        CONFIG.GEMINI_API_KEY_PROPERTY
      );


  if (!apiKey) {

    throw new Error(
      "Gemini API key was not found in Script Properties."
    );

  }


  Logger.log(
    "Gemini API key found in Script Properties."
  );


  Logger.log(
    `Model: ${CONFIG.GEMINI_MODEL}`
  );


  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    CONFIG.GEMINI_MODEL +
    ":generateContent?key=" +
    encodeURIComponent(apiKey);


  const payload = {

    contents: [

      {

        parts: [

          {
            text:
              "Reply with exactly: GEMINI CONNECTION SUCCESSFUL"
          }

        ]

      }

    ]

  };


  const options = {

    method: "post",

    contentType:
      "application/json",

    payload:
      JSON.stringify(payload),

    muteHttpExceptions:
      true

  };


  const response =
    UrlFetchApp.fetch(
      url,
      options
    );


  const responseCode =
    response.getResponseCode();


  const responseText =
    response.getContentText();


  Logger.log(
    `HTTP Response Code: ${responseCode}`
  );


  if (responseCode !== 200) {

    Logger.log("=================================");
    Logger.log("GEMINI API ERROR");
    Logger.log("=================================");

    Logger.log(responseText);

    throw new Error(
      `Gemini API returned HTTP ${responseCode}.`
    );

  }


  let result;


  try {

    result =
      JSON.parse(responseText);

  } catch (error) {

    Logger.log(responseText);

    throw new Error(
      "Could not parse Gemini response."
    );

  }


  const generatedText =
    result
      ?.candidates?.[0]
      ?.content?.parts?.[0]
      ?.text;


  Logger.log("=================================");
  Logger.log("GEMINI RESPONSE");
  Logger.log("=================================");


  Logger.log(
    generatedText || responseText
  );


  Logger.log("=================================");
  Logger.log(
    "GEMINI CONNECTION SUCCESSFUL"
  );
  Logger.log("=================================");

}


/*************************************************
 * STEP 6D
 * GOOGLE SHEETS → GEMINI TEST
 *************************************************/

function testGeminiWithTodaysContent() {

  Logger.log("=================================");
  Logger.log("STEP 6D — SHEETS → GEMINI TEST");
  Logger.log("=================================");


  // -----------------------------------------------
  // 1. Gemini API key
  // -----------------------------------------------

  const apiKey =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        CONFIG.GEMINI_API_KEY_PROPERTY
      );


  if (!apiKey) {

    throw new Error(
      "Gemini API key was not found in Script Properties."
    );

  }


  Logger.log(
    "Gemini API key found."
  );


  // -----------------------------------------------
  // 2. Sheet
  // -----------------------------------------------

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();


  const sheet =
    spreadsheet
      .getSheetByName(
        CONFIG.SHEET_NAME
      );


  if (!sheet) {

    throw new Error(
      `Sheet "${CONFIG.SHEET_NAME}" was not found.`
    );

  }


  Logger.log(
    `Sheet found: ${sheet.getName()}`
  );


  // -----------------------------------------------
  // 3. Read data
  // -----------------------------------------------

  const data =
    sheet.getDataRange().getValues();


  Logger.log(
    `Total rows found: ${data.length - 1}`
  );


  const headers =
    data[0];


  Logger.log("=================================");
  Logger.log("SHEET HEADERS");
  Logger.log("=================================");


  headers.forEach(function(header, index) {

    Logger.log(
      `${index + 1}: [${header}]`
    );

  });


  // -----------------------------------------------
  // 4. Required columns
  // -----------------------------------------------

  const requiredColumns = [

    "Date",
    "Account",
    "Content Type",
    "Topic",
    "Content + Generation Prompt",
    "Caption",
    "SEO Keywords",
    "5 Hashtags",
    "Audio / Editing Direction",
    "Repurpose / Collab Note",
    "Status",
    "Approval",
    "Asset URL",
    "Publish Time",
    "IG Status"

  ];


  requiredColumns.forEach(function(column) {

    if (
      headers.indexOf(column) === -1
    ) {

      throw new Error(
        `Column "${column}" was not found.`
      );

    }

  });


  Logger.log(
    "All required columns found successfully."
  );


  // -----------------------------------------------
  // 5. Today's date
  // -----------------------------------------------

  const timezone =
    spreadsheet.getSpreadsheetTimeZone();


  const todayString =
    Utilities.formatDate(
      new Date(),
      timezone,
      "yyyy-MM-dd"
    );


  Logger.log(
    `Today's date: ${todayString}`
  );


  Logger.log(
    `Spreadsheet timezone: ${timezone}`
  );


  // -----------------------------------------------
  // 6. Find today's READY content
  // -----------------------------------------------

  const dateIndex =
    headers.indexOf("Date");


  const statusIndex =
    headers.indexOf("Status");


  const todaysReadyContent = [];


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const row =
      data[i];


    let rowDate =
      row[dateIndex];


    if (!rowDate) {
      continue;
    }


    let rowDateString = "";


    if (rowDate instanceof Date) {

      rowDateString =
        Utilities.formatDate(
          rowDate,
          timezone,
          "yyyy-MM-dd"
        );

    } else {

      rowDateString =
        String(rowDate).trim();


      if (
        /^\d{4}-\d{2}-\d{2}/
          .test(rowDateString)
      ) {

        rowDateString =
          rowDateString.substring(
            0,
            10
          );

      }

    }


    const status =
      String(
        row[statusIndex]
      )
        .trim()
        .toUpperCase();


    if (

      rowDateString ===
        todayString

      &&

      status === "READY"

    ) {

      const item = {

        sheetRow: i + 1,

        date:
          row[dateIndex],

        account:
          row[
            headers.indexOf(
              "Account"
            )
          ],

        contentType:
          row[
            headers.indexOf(
              "Content Type"
            )
          ],

        topic:
          row[
            headers.indexOf(
              "Topic"
            )
          ],

        contentPrompt:
          row[
            headers.indexOf(
              "Content + Generation Prompt"
            )
          ],

        caption:
          row[
            headers.indexOf(
              "Caption"
            )
          ],

        seoKeywords:
          row[
            headers.indexOf(
              "SEO Keywords"
            )
          ],

        hashtags:
          row[
            headers.indexOf(
              "5 Hashtags"
            )
          ],

        audioDirection:
          row[
            headers.indexOf(
              "Audio / Editing Direction"
            )
          ],

        repurposeNote:
          row[
            headers.indexOf(
              "Repurpose / Collab Note"
            )
          ],

        status:
          row[
            statusIndex
          ],

        approval:
          row[
            headers.indexOf(
              "Approval"
            )
          ],

        assetUrl:
          row[
            headers.indexOf(
              "Asset URL"
            )
          ],

        publishTime:
          row[
            headers.indexOf(
              "Publish Time"
            )
          ],

        igStatus:
          row[
            headers.indexOf(
              "IG Status"
            )
          ]

      };


      todaysReadyContent.push(
        item
      );

    }

  }


  Logger.log("=================================");
  Logger.log(
    `Today's READY content: ${todaysReadyContent.length}`
  );
  Logger.log("=================================");


  if (
    todaysReadyContent.length === 0
  ) {

    Logger.log(
      "No READY content found for today."
    );

    return;

  }


  // -----------------------------------------------
  // 7. Select first READY item
  // -----------------------------------------------

  const content =
    todaysReadyContent[0];


  Logger.log("=================================");
  Logger.log("CONTENT SELECTED");
  Logger.log("=================================");


  Logger.log(
    `Sheet Row: ${content.sheetRow}`
  );


  Logger.log(
    `Account: ${content.account}`
  );


  Logger.log(
    `Content Type: ${content.contentType}`
  );


  Logger.log(
    `Topic: ${content.topic}`
  );


  // -----------------------------------------------
  // 8. Display prompt
  // -----------------------------------------------

  Logger.log("=================================");
  Logger.log(
    "CONTENT + GENERATION PROMPT"
  );
  Logger.log("=================================");


  Logger.log(
    `Topic: ${content.topic} ` +
    `${content.contentPrompt}`
  );


  // -----------------------------------------------
  // 9. Gemini prompt
  // -----------------------------------------------

  const prompt = `

Improve the following social media content
generation prompt.

Account:
${content.account}

Content Type:
${content.contentType}

Topic:
${content.topic}

Original Content + Generation Prompt:
${content.contentPrompt}

Return only the improved generation prompt.

`;


  Logger.log("=================================");
  Logger.log("SENDING CONTENT TO GEMINI");
  Logger.log("=================================");


  // -----------------------------------------------
  // 10. Gemini API
  // -----------------------------------------------

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    CONFIG.GEMINI_MODEL +
    ":generateContent?key=" +
    encodeURIComponent(apiKey);


  const payload = {

    contents: [

      {

        parts: [

          {
            text: prompt
          }

        ]

      }

    ]

  };


  const options = {

    method: "post",

    contentType:
      "application/json",

    payload:
      JSON.stringify(payload),

    muteHttpExceptions:
      true

  };


  const response =
    UrlFetchApp.fetch(
      url,
      options
    );


  const responseCode =
    response.getResponseCode();


  const responseText =
    response.getContentText();


  Logger.log(
    `Gemini HTTP Response: ${responseCode}`
  );


  if (responseCode !== 200) {

    Logger.log(
      "================================="
    );

    Logger.log(
      "GEMINI API ERROR"
    );

    Logger.log(
      "================================="
    );

    Logger.log(
      responseText
    );


    throw new Error(
      `Gemini API returned HTTP ${responseCode}.`
    );

  }


  // -----------------------------------------------
  // 11. Parse response
  // -----------------------------------------------

  let result;


  try {

    result =
      JSON.parse(responseText);

  } catch (error) {

    throw new Error(
      "Could not parse Gemini response."
    );

  }


  const generatedText =
    result
      ?.candidates?.[0]
      ?.content?.parts?.[0]
      ?.text;


  if (!generatedText) {

    throw new Error(
      "Gemini returned no generated content."
    );

  }


  Logger.log("=================================");
  Logger.log(
    "GEMINI IMPROVED PROMPT"
  );
  Logger.log("=================================");


  Logger.log(
    generatedText
  );


  Logger.log("=================================");
  Logger.log(
    "STEP 6D TEST COMPLETE"
  );
  Logger.log("=================================");


  Logger.log(
    "Google Sheet was NOT modified."
  );

}


/*************************************************
 * STEP 6E
 * GEMINI STRUCTURED OUTPUT TEST
 *************************************************/

function testGeminiStructuredOutput() {

  Logger.log("=================================");
  Logger.log(
    "STEP 6E — GEMINI STRUCTURED OUTPUT"
  );
  Logger.log("=================================");


  // -----------------------------------------------
  // 1. API key
  // -----------------------------------------------

  const apiKey =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        CONFIG.GEMINI_API_KEY_PROPERTY
      );


  if (!apiKey) {

    throw new Error(
      "Gemini API key was not found in Script Properties."
    );

  }


  Logger.log(
    "Gemini API key found."
  );


  // -----------------------------------------------
  // 2. Spreadsheet
  // -----------------------------------------------

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();


  const sheet =
    spreadsheet
      .getSheetByName(
        CONFIG.SHEET_NAME
      );


  if (!sheet) {

    throw new Error(
      `Sheet "${CONFIG.SHEET_NAME}" was not found.`
    );

  }


  Logger.log(
    `Sheet found: ${sheet.getName()}`
  );


  // -----------------------------------------------
  // 3. Read sheet
  // -----------------------------------------------

  const data =
    sheet.getDataRange().getValues();


  if (data.length < 2) {

    throw new Error(
      "No content rows were found."
    );

  }


  const headers =
    data[0];


  // -----------------------------------------------
  // 4. Required columns
  // -----------------------------------------------

  const requiredColumns = [

    "Date",
    "Account",
    "Content Type",
    "Topic",
    "Content + Generation Prompt",
    "Caption",
    "SEO Keywords",
    "5 Hashtags",
    "Audio / Editing Direction",
    "Repurpose / Collab Note",
    "Status",
    "Approval",
    "Asset URL",
    "Publish Time",
    "IG Status"

  ];


  requiredColumns.forEach(function(column) {

    if (
      headers.indexOf(column) === -1
    ) {

      throw new Error(
        `Column "${column}" was not found.`
      );

    }

  });


  Logger.log(
    "All required columns found successfully."
  );


  // -----------------------------------------------
  // 5. Today's date
  // -----------------------------------------------

  const timezone =
    spreadsheet.getSpreadsheetTimeZone();


  const todayString =
    Utilities.formatDate(
      new Date(),
      timezone,
      "yyyy-MM-dd"
    );


  Logger.log(
    `Today's date: ${todayString}`
  );


  Logger.log(
    `Spreadsheet timezone: ${timezone}`
  );


  // -----------------------------------------------
  // 6. Find today's READY content
  // -----------------------------------------------

  const dateIndex =
    headers.indexOf("Date");


  const statusIndex =
    headers.indexOf("Status");


  const todaysReadyContent = [];


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const row =
      data[i];


    let rowDate =
      row[dateIndex];


    if (!rowDate) {
      continue;
    }


    let rowDateString = "";


    if (rowDate instanceof Date) {

      rowDateString =
        Utilities.formatDate(
          rowDate,
          timezone,
          "yyyy-MM-dd"
        );

    } else {

      rowDateString =
        String(rowDate).trim();


      if (
        /^\d{4}-\d{2}-\d{2}/
          .test(rowDateString)
      ) {

        rowDateString =
          rowDateString.substring(
            0,
            10
          );

      }

    }


    const status =
      String(
        row[statusIndex]
      )
        .trim()
        .toUpperCase();


    if (

      rowDateString ===
        todayString

      &&

      status === "READY"

    ) {

      const item = {

        sheetRow: i + 1,

        date:
          row[dateIndex],

        account:
          row[
            headers.indexOf(
              "Account"
            )
          ],

        contentType:
          row[
            headers.indexOf(
              "Content Type"
            )
          ],

        topic:
          row[
            headers.indexOf(
              "Topic"
            )
          ],

        contentPrompt:
          row[
            headers.indexOf(
              "Content + Generation Prompt"
            )
          ],

        caption:
          row[
            headers.indexOf(
              "Caption"
            )
          ],

        seoKeywords:
          row[
            headers.indexOf(
              "SEO Keywords"
            )
          ],

        hashtags:
          row[
            headers.indexOf(
              "5 Hashtags"
            )
          ],

        audioDirection:
          row[
            headers.indexOf(
              "Audio / Editing Direction"
            )
          ],

        repurposeNote:
          row[
            headers.indexOf(
              "Repurpose / Collab Note"
            )
          ],

        status:
          row[
            statusIndex
          ],

        approval:
          row[
            headers.indexOf(
              "Approval"
            )
          ],

        assetUrl:
          row[
            headers.indexOf(
              "Asset URL"
            )
          ],

        publishTime:
          row[
            headers.indexOf(
              "Publish Time"
            )
          ],

        igStatus:
          row[
            headers.indexOf(
              "IG Status"
            )
          ]

      };


      todaysReadyContent.push(
        item
      );

    }

  }


  Logger.log("=================================");
  Logger.log(
    `Today's READY content: ${todaysReadyContent.length}`
  );
  Logger.log("=================================");


  if (
    todaysReadyContent.length === 0
  ) {

    Logger.log(
      "No READY content found for today."
    );

    return;

  }


  // -----------------------------------------------
  // 7. Select first READY item
  // -----------------------------------------------

  const content =
    todaysReadyContent[0];


  Logger.log("=================================");
  Logger.log("CONTENT SELECTED");
  Logger.log("=================================");


  Logger.log(
    `Sheet Row: ${content.sheetRow}`
  );


  Logger.log(
    `Account: ${content.account}`
  );


  Logger.log(
    `Content Type: ${content.contentType}`
  );


  Logger.log(
    `Topic: ${content.topic}`
  );


  // -----------------------------------------------
  // 8. Gemini instruction
  // -----------------------------------------------

  const instruction = `

You are an expert AI video and social media
content strategist.

Convert the following content calendar item
into structured production information.

IMPORTANT:

1. Keep the original creative idea.
2. Improve the generation prompt if necessary.
3. Keep the output practical for AI content generation.
4. Do not invent a brand name.
5. Do not invent factual claims.
6. Keep the CTA appropriate for the account.
7. Preserve the intended content type.
8. Return ONLY valid JSON.
9. Do not use Markdown.
10. Do not wrap the JSON in code fences.

CONTENT:

Account:
${content.account}

Content Type:
${content.contentType}

Topic:
${content.topic}

Content + Generation Prompt:
${content.contentPrompt}

Existing Caption:
${content.caption}

Existing SEO Keywords:
${content.seoKeywords}

Existing Hashtags:
${content.hashtags}

Audio / Editing Direction:
${content.audioDirection}

Repurpose / Collab Note:
${content.repurposeNote}


Return exactly this JSON structure:

{
  "video_prompt": "string",
  "on_screen_text": "string",
  "cta": "string",
  "caption": "string",
  "seo_keywords": "string",
  "hashtags": "string"
}

`;


  Logger.log("=================================");
  Logger.log(
    "SENDING CONTENT TO GEMINI"
  );
  Logger.log("=================================");


  // -----------------------------------------------
  // 9. Gemini API
  // -----------------------------------------------

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    CONFIG.GEMINI_MODEL +
    ":generateContent?key=" +
    encodeURIComponent(apiKey);


  const payload = {

    contents: [

      {

        parts: [

          {
            text: instruction
          }

        ]

      }

    ],

    generationConfig: {

      temperature: 0.4,

      responseMimeType:
        "application/json"

    }

  };


  const options = {

    method: "post",

    contentType:
      "application/json",

    payload:
      JSON.stringify(payload),

    muteHttpExceptions:
      true

  };


  const response =
    UrlFetchApp.fetch(
      url,
      options
    );


  const responseCode =
    response.getResponseCode();


  const responseText =
    response.getContentText();


  Logger.log(
    `Gemini HTTP Response: ${responseCode}`
  );


  if (responseCode !== 200) {

    Logger.log("=================================");
    Logger.log("GEMINI API ERROR");
    Logger.log("=================================");

    Logger.log(
      responseText
    );


    throw new Error(
      `Gemini API returned HTTP ${responseCode}.`
    );

  }


  // -----------------------------------------------
  // 10. Parse API response
  // -----------------------------------------------

  let geminiResponse;


  try {

    geminiResponse =
      JSON.parse(
        responseText
      );

  } catch (error) {

    Logger.log(
      responseText
    );


    throw new Error(
      "Unable to parse Gemini API response."
    );

  }


  // -----------------------------------------------
  // 11. Extract generated text
  // -----------------------------------------------

  const generatedText =
    geminiResponse
      ?.candidates?.[0]
      ?.content?.parts?.[0]
      ?.text;


  if (!generatedText) {

    Logger.log(
      JSON.stringify(
        geminiResponse,
        null,
        2
      )
    );


    throw new Error(
      "Gemini returned no generated content."
    );

  }


  // -----------------------------------------------
  // 12. Parse structured JSON
  // -----------------------------------------------

  let structured;


  try {

    structured =
      JSON.parse(
        generatedText
      );

  } catch (error) {

    Logger.log(
      "Gemini returned invalid JSON:"
    );


    Logger.log(
      generatedText
    );


    throw new Error(
      "Gemini structured output could not be parsed."
    );

  }


  // -----------------------------------------------
  // 13. Validate fields
  // -----------------------------------------------

  const expectedFields = [

    "video_prompt",
    "on_screen_text",
    "cta",
    "caption",
    "seo_keywords",
    "hashtags"

  ];


  expectedFields.forEach(function(field) {

    if (

      structured[field] ===
        undefined

      ||

      structured[field] ===
        null

    ) {

      throw new Error(
        `Gemini JSON is missing field: ${field}`
      );

    }

  });


  // -----------------------------------------------
  // 14. Display structured output
  // -----------------------------------------------

  Logger.log("=================================");
  Logger.log(
    "STRUCTURED GEMINI OUTPUT"
  );
  Logger.log("=================================");


  Logger.log(
    `Video Prompt:\n${structured.video_prompt}`
  );


  Logger.log(
    `On-Screen Text:\n${structured.on_screen_text}`
  );


  Logger.log(
    `CTA:\n${structured.cta}`
  );


  Logger.log(
    `Caption:\n${structured.caption}`
  );


  Logger.log(
    `SEO Keywords:\n${structured.seo_keywords}`
  );


  Logger.log(
    `Hashtags:\n${structured.hashtags}`
  );


  // -----------------------------------------------
  // 15. Raw JSON
  // -----------------------------------------------

  Logger.log("=================================");
  Logger.log(
    "RAW STRUCTURED JSON"
  );
  Logger.log("=================================");


  Logger.log(
    JSON.stringify(
      structured,
      null,
      2
    )
  );


  // -----------------------------------------------
  // 16. Final test status
  // -----------------------------------------------

  Logger.log("=================================");
  Logger.log(
    "STEP 6E TEST COMPLETE"
  );
  Logger.log("=================================");


  Logger.log(
    "Google Sheet was NOT modified."
  );


  Logger.log(
    "Veo was NOT called."
  );


  Logger.log(
    "Google Drive was NOT modified."
  );


  Logger.log(
    "Meta API was NOT called."
  );


  Logger.log("=================================");

}



6d:-
----
/*************************************************
 * AICoreStack Content Automation
 *
 * STEP 6D — Google Sheets → Gemini
 *
 * PURPOSE:
 * Read today's READY content from Google Sheets
 * and send the Content + Generation Prompt to Gemini.
 *
 * IMPORTANT:
 * This is a TEST version.
 *
 * It does NOT modify the Google Sheet.
 *
 * FLOW:
 *
 * Google Sheets
 *      ↓
 * Today's READY row
 *      ↓
 * Content + Generation Prompt
 *      ↓
 * Gemini 3.5 Flash-Lite
 *      ↓
 * Improved production prompt
 *      ↓
 * Execution Log
 *************************************************/


const CONFIG = {

  SHEET_NAME: "90 Content Calendar",

  HEADER_ROW: 1,

  GEMINI_MODEL: "gemini-3.5-flash-lite",

  GEMINI_ENDPOINT:
    "https://generativelanguage.googleapis.com/v1beta/models/",

  GEMINI_API_KEY_PROPERTY:
    "GEMINI_API_KEY"

};


/**
 * =================================================
 * STEP 6D
 *
 * Google Sheets → Gemini
 * =================================================
 */
function testGeminiWithTodaysContent() {

  Logger.log(
    "================================="
  );

  Logger.log(
    "STEP 6D — SHEETS → GEMINI TEST"
  );

  Logger.log(
    "================================="
  );


  /**
   * -----------------------------------------------
   * 1. GET GEMINI API KEY
   * -----------------------------------------------
   */

  const apiKey =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        CONFIG.GEMINI_API_KEY_PROPERTY
      );


  if (!apiKey) {

    throw new Error(
      "GEMINI_API_KEY was not found in Script Properties."
    );

  }


  Logger.log(
    "Gemini API key found."
  );


  /**
   * -----------------------------------------------
   * 2. GET GOOGLE SHEET
   * -----------------------------------------------
   */

  const spreadsheet =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheet =
    spreadsheet
      .getSheetByName(
        CONFIG.SHEET_NAME
      );


  if (!sheet) {

    throw new Error(
      `Sheet "${CONFIG.SHEET_NAME}" was not found.`
    );

  }


  Logger.log(
    `Sheet found: ${sheet.getName()}`
  );


  /**
   * -----------------------------------------------
   * 3. READ SHEET
   * -----------------------------------------------
   */

  const data =
    sheet
      .getDataRange()
      .getValues();


  if (data.length <= 1) {

    throw new Error(
      "No content rows were found."
    );

  }


  const headers =
    data[
      CONFIG.HEADER_ROW - 1
    ];


  Logger.log(
    `Total rows found: ${data.length - 1}`
  );


  /**
   * -----------------------------------------------
   * 4. FIND EXACT COLUMN NAMES
   * -----------------------------------------------
   */

  const dateIndex =
    headers.indexOf("Date");


  const accountIndex =
    headers.indexOf("Account");


  const contentTypeIndex =
    headers.indexOf("Content Type");


  const topicIndex =
    headers.indexOf("Topic");


  const promptIndex =
    headers.indexOf(
      "Content + Generation Prompt"
    );


  const statusIndex =
    headers.indexOf("Status");


  /**
   * -----------------------------------------------
   * 5. VERIFY COLUMNS
   * -----------------------------------------------
   */

  if (dateIndex === -1) {

    throw new Error(
      'Column "Date" was not found.'
    );

  }


  if (accountIndex === -1) {

    throw new Error(
      'Column "Account" was not found.'
    );

  }


  if (contentTypeIndex === -1) {

    throw new Error(
      'Column "Content Type" was not found.'
    );

  }


  if (topicIndex === -1) {

    throw new Error(
      'Column "Topic" was not found.'
    );

  }


  if (promptIndex === -1) {

    throw new Error(
      'Column "Content + Generation Prompt" was not found.'
    );

  }


  if (statusIndex === -1) {

    throw new Error(
      'Column "Status" was not found.'
    );

  }


  Logger.log(
    "All required columns found successfully."
  );


  /**
   * -----------------------------------------------
   * 6. GET TODAY'S DATE
   * -----------------------------------------------
   */

  const timezone =
    spreadsheet
      .getSpreadsheetTimeZone();


  const todayString =
    Utilities.formatDate(
      new Date(),
      timezone,
      "yyyy-MM-dd"
    );


  Logger.log(
    `Today's date: ${todayString}`
  );


  Logger.log(
    `Spreadsheet timezone: ${timezone}`
  );


  /**
   * -----------------------------------------------
   * 7. FIND TODAY'S READY CONTENT
   * -----------------------------------------------
   */

  const readyItems = [];


  for (
    let i = CONFIG.HEADER_ROW;
    i < data.length;
    i++
  ) {

    const row =
      data[i];


    /**
     * Read date
     */

    const rowDate =
      row[dateIndex];


    if (!rowDate) {

      continue;

    }


    let rowDateString = "";


    /**
     * Google Sheets date object
     */

    if (
      rowDate instanceof Date
    ) {

      rowDateString =
        Utilities.formatDate(
          rowDate,
          timezone,
          "yyyy-MM-dd"
        );

    }

    else {

      rowDateString =
        String(rowDate)
          .trim();


      if (
        /^\d{4}-\d{2}-\d{2}/
          .test(rowDateString)
      ) {

        rowDateString =
          rowDateString.substring(
            0,
            10
          );

      }

    }


    /**
     * Check date
     */

    if (
      rowDateString !== todayString
    ) {

      continue;

    }


    /**
     * Check status
     */

    const status =
      String(
        row[statusIndex]
      )
        .trim()
        .toUpperCase();


    if (
      status !== "READY"
    ) {

      continue;

    }


    /**
     * -------------------------------------------
     * STORE READY ITEM
     * -------------------------------------------
     */

    readyItems.push({

      rowNumber:
        i + 1,

      account:
        row[accountIndex],

      contentType:
        row[contentTypeIndex],

      topic:
        row[topicIndex],

      prompt:
        row[promptIndex]

    });

  }


  /**
   * -----------------------------------------------
   * 8. SHOW NUMBER OF READY ITEMS
   * -----------------------------------------------
   */

  Logger.log(
    "================================="
  );

  Logger.log(
    `Today's READY content: ${readyItems.length}`
  );

  Logger.log(
    "================================="
  );


  /**
   * -----------------------------------------------
   * NO READY CONTENT
   * -----------------------------------------------
   */

  if (
    readyItems.length === 0
  ) {

    Logger.log(
      "No READY content found for today."
    );

    return;

  }


  /**
   * -----------------------------------------------
   * 9. TEST ONLY FIRST ITEM
   *
   * We intentionally process only one item
   * during Step 6D testing.
   * -----------------------------------------------
   */

  const item =
    readyItems[0];


  Logger.log(
    "================================="
  );

  Logger.log(
    "CONTENT SELECTED"
  );

  Logger.log(
    "================================="
  );

  Logger.log(
    `Sheet Row: ${item.rowNumber}`
  );

  Logger.log(
    `Account: ${item.account}`
  );

  Logger.log(
    `Content Type: ${item.contentType}`
  );

  Logger.log(
    `Topic: ${item.topic}`
  );


  /**
   * -----------------------------------------------
   * 10. SHOW ORIGINAL PROMPT
   * -----------------------------------------------
   */

  Logger.log(
    "================================="
  );

  Logger.log(
    "CONTENT + GENERATION PROMPT"
  );

  Logger.log(
    "================================="
  );

  Logger.log(
    String(item.prompt)
  );


  /**
   * -----------------------------------------------
   * 11. BUILD GEMINI REQUEST
   * -----------------------------------------------
   */

  const geminiPrompt = `

You are an AI content production assistant.

We are building an automated content workflow
for AI video production and digital marketing.

Read the content and generation instructions
below and create a clean production-ready
generation prompt.

IMPORTANT RULES:

1. Preserve the original creative idea.
2. Do not change the intended marketing message.
3. Do not add unnecessary hype.
4. Keep the visual direction cinematic,
   realistic and professional.
5. Keep any requested duration.
6. Keep any requested aspect ratio.
7. Keep important visual details.
8. Do not add logos unless explicitly requested.
9. Do not add unrelated scenes.
10. Return ONLY the improved generation prompt.
11. Do not explain your changes.

ACCOUNT:
${item.account}

CONTENT TYPE:
${item.contentType}

TOPIC:
${item.topic}

CONTENT + GENERATION PROMPT:
${item.prompt}

Now create the final production-ready
generation prompt.
`;


  /**
   * -----------------------------------------------
   * 12. GEMINI API URL
   * -----------------------------------------------
   */

  const url =
    CONFIG.GEMINI_ENDPOINT +
    CONFIG.GEMINI_MODEL +
    ":generateContent";


  /**
   * -----------------------------------------------
   * 13. GEMINI REQUEST
   * -----------------------------------------------
   */

  const payload = {

    contents: [

      {

        role: "user",

        parts: [

          {

            text:
              geminiPrompt

          }

        ]

      }

    ],

    generationConfig: {

      temperature: 0.2,

      maxOutputTokens: 1000

    }

  };


  const options = {

    method: "POST",

    contentType:
      "application/json",

    headers: {

      "x-goog-api-key":
        apiKey

    },

    payload:
      JSON.stringify(payload),

    muteHttpExceptions:
      true

  };


  /**
   * -----------------------------------------------
   * 14. SEND TO GEMINI
   * -----------------------------------------------
   */

  Logger.log(
    "================================="
  );

  Logger.log(
    "SENDING CONTENT TO GEMINI..."
  );

  Logger.log(
    "================================="
  );


  const response =
    UrlFetchApp.fetch(
      url,
      options
    );


  const responseCode =
    response.getResponseCode();


  const responseText =
    response.getContentText();


  Logger.log(
    `Gemini HTTP Response: ${responseCode}`
  );


  /**
   * -----------------------------------------------
   * 15. HANDLE GEMINI ERROR
   * -----------------------------------------------
   */

  if (
    responseCode < 200 ||
    responseCode >= 300
  ) {

    Logger.log(
      "================================="
    );

    Logger.log(
      "GEMINI API ERROR"
    );

    Logger.log(
      responseText
    );

    Logger.log(
      "================================="
    );

    throw new Error(
      `Gemini API returned HTTP ${responseCode}.`
    );

  }


  /**
   * -----------------------------------------------
   * 16. PARSE GEMINI RESPONSE
   * -----------------------------------------------
   */

  const geminiData =
    JSON.parse(
      responseText
    );


  /**
   * -----------------------------------------------
   * 17. EXTRACT AI TEXT
   * -----------------------------------------------
   */

  let improvedPrompt = "";


  try {

    improvedPrompt =
      geminiData
        .candidates[0]
        .content
        .parts[0]
        .text;

  }

  catch (error) {

    Logger.log(
      "Unexpected Gemini response:"
    );

    Logger.log(
      responseText
    );

    throw new Error(
      "Could not extract Gemini response."
    );

  }


  /**
   * -----------------------------------------------
   * 18. SHOW GEMINI RESULT
   * -----------------------------------------------
   */

  Logger.log(
    "================================="
  );

  Logger.log(
    "GEMINI IMPROVED PROMPT"
  );

  Logger.log(
    "================================="
  );

  Logger.log(
    improvedPrompt
  );


  /**
   * -----------------------------------------------
   * 19. IMPORTANT SAFETY CHECK
   *
   * DO NOT MODIFY SHEET YET.
   * -----------------------------------------------
   */

  Logger.log(
    "================================="
  );

  Logger.log(
    "STEP 6D TEST COMPLETE"
  );

  Logger.log(
    "Google Sheet was NOT modified."
  );

  Logger.log(
    "================================="
  );

}


6E:6E — Structured Gemini output
----


/*************************************************
 * AICoreStack Content Automation
 *
 * STEP 6E — GEMINI STRUCTURED OUTPUT TEST
 *
 * This version:
 * Google Sheets → Gemini → Structured JSON
 *
 * NO Veo
 * NO Drive upload
 * NO Meta API
 * NO Sheet modification
 *************************************************/


const CONFIG = {

  SHEET_NAME: "90 Content Calendar",

  HEADER_ROW: 1,

  GEMINI_MODEL: "gemini-3.5-flash-lite",

  GEMINI_API_KEY_PROPERTY: "GEMINI_API_KEY"

};


/**
 * =================================================
 * STEP 6E
 * Test Gemini structured output using today's
 * first READY content.
 * =================================================
 */
function testGeminiStructuredOutput() {

  Logger.log("=================================");
  Logger.log("STEP 6E — GEMINI STRUCTURED OUTPUT");
  Logger.log("=================================");


  // -----------------------------------------------
  // 1. Get Gemini API key
  // -----------------------------------------------

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty(CONFIG.GEMINI_API_KEY_PROPERTY);


  if (!apiKey) {

    throw new Error(
      "Gemini API key was not found in Script Properties."
    );

  }


  Logger.log("Gemini API key found.");


  // -----------------------------------------------
  // 2. Find spreadsheet
  // -----------------------------------------------

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();


  const sheet =
    spreadsheet.getSheetByName(CONFIG.SHEET_NAME);


  if (!sheet) {

    throw new Error(
      `Sheet "${CONFIG.SHEET_NAME}" was not found.`
    );

  }


  Logger.log(
    `Sheet found: ${sheet.getName()}`
  );


  // -----------------------------------------------
  // 3. Read sheet
  // -----------------------------------------------

  const data =
    sheet.getDataRange().getValues();


  if (data.length < 2) {

    throw new Error(
      "No content rows were found."
    );

  }


  const headers = data[0];


  // -----------------------------------------------
  // 4. Required columns
  // -----------------------------------------------

  const requiredColumns = [

    "Date",
    "Account",
    "Content Type",
    "Topic",
    "Content + Generation Prompt",
    "Caption",
    "SEO Keywords",
    "5 Hashtags",
    "Audio / Editing Direction",
    "Repurpose / Collab Note",
    "Status",
    "Approval",
    "Asset URL",
    "Publish Time",
    "IG Status"

  ];


  requiredColumns.forEach(function(column) {

    if (headers.indexOf(column) === -1) {

      throw new Error(
        `Column "${column}" was not found.`
      );

    }

  });


  Logger.log(
    "All required columns found successfully."
  );


  // -----------------------------------------------
  // 5. Find today's date
  // -----------------------------------------------

  const timezone =
    spreadsheet.getSpreadsheetTimeZone();


  const todayString =
    Utilities.formatDate(
      new Date(),
      timezone,
      "yyyy-MM-dd"
    );


  Logger.log(
    `Today's date: ${todayString}`
  );


  Logger.log(
    `Spreadsheet timezone: ${timezone}`
  );


  // -----------------------------------------------
  // 6. Find today's READY content
  // -----------------------------------------------

  const dateIndex =
    headers.indexOf("Date");


  const statusIndex =
    headers.indexOf("Status");


  const todaysReadyContent = [];


  for (let i = 1; i < data.length; i++) {

    const row = data[i];


    let rowDate = row[dateIndex];


    if (!rowDate) {
      continue;
    }


    let rowDateString = "";


    if (rowDate instanceof Date) {

      rowDateString =
        Utilities.formatDate(
          rowDate,
          timezone,
          "yyyy-MM-dd"
        );

    } else {

      rowDateString =
        String(rowDate).trim();

      if (
        /^\d{4}-\d{2}-\d{2}/
          .test(rowDateString)
      ) {

        rowDateString =
          rowDateString.substring(0, 10);

      }

    }


    const status =
      String(row[statusIndex])
        .trim()
        .toUpperCase();


    if (
      rowDateString === todayString &&
      status === "READY"
    ) {

      const item = {

        sheetRow: i + 1,

        date: row[dateIndex],

        account:
          row[headers.indexOf("Account")],

        contentType:
          row[headers.indexOf("Content Type")],

        topic:
          row[headers.indexOf("Topic")],

        contentPrompt:
          row[
            headers.indexOf(
              "Content + Generation Prompt"
            )
          ],

        caption:
          row[headers.indexOf("Caption")],

        seoKeywords:
          row[headers.indexOf("SEO Keywords")],

        hashtags:
          row[headers.indexOf("5 Hashtags")],

        audioDirection:
          row[
            headers.indexOf(
              "Audio / Editing Direction"
            )
          ],

        repurposeNote:
          row[
            headers.indexOf(
              "Repurpose / Collab Note"
            )
          ],

        status:
          row[statusIndex],

        approval:
          row[headers.indexOf("Approval")],

        assetUrl:
          row[headers.indexOf("Asset URL")],

        publishTime:
          row[headers.indexOf("Publish Time")],

        igStatus:
          row[headers.indexOf("IG Status")]

      };


      todaysReadyContent.push(item);

    }

  }


  Logger.log(
    `Today's READY content: ${todaysReadyContent.length}`
  );


  if (todaysReadyContent.length === 0) {

    Logger.log(
      "No READY content found for today."
    );

    return;

  }


  // -----------------------------------------------
  // 7. Select first READY item
  // -----------------------------------------------

  const content =
    todaysReadyContent[0];


  Logger.log("=================================");
  Logger.log("CONTENT SELECTED");
  Logger.log("=================================");


  Logger.log(
    `Sheet Row: ${content.sheetRow}`
  );


  Logger.log(
    `Account: ${content.account}`
  );


  Logger.log(
    `Content Type: ${content.contentType}`
  );


  Logger.log(
    `Topic: ${content.topic}`
  );


  // -----------------------------------------------
  // 8. Build Gemini instruction
  // -----------------------------------------------

  const instruction = `

You are an expert AI video content strategist.

We are automating a social media content workflow.

Convert the following content calendar item into
structured production information.

IMPORTANT:

1. Keep the original creative idea.
2. Improve the video generation prompt if necessary.
3. Keep the output practical for AI video generation.
4. Do not invent a brand name.
5. Do not invent factual claims.
6. Keep the CTA appropriate for the account.
7. Preserve the intended content type.
8. Return ONLY valid JSON.
9. Do not use Markdown.
10. Do not wrap the JSON in code fences.

CONTENT:

Account:
${content.account}

Content Type:
${content.contentType}

Topic:
${content.topic}

Content + Generation Prompt:
${content.contentPrompt}

Existing Caption:
${content.caption}

Existing SEO Keywords:
${content.seoKeywords}

Existing Hashtags:
${content.hashtags}

Audio / Editing Direction:
${content.audioDirection}

Repurpose / Collab Note:
${content.repurposeNote}


Return exactly this JSON structure:

{
  "video_prompt": "string",
  "on_screen_text": "string",
  "cta": "string",
  "caption": "string",
  "seo_keywords": "string",
  "hashtags": "string"
}

`;


  Logger.log("=================================");
  Logger.log("SENDING CONTENT TO GEMINI");
  Logger.log("=================================");


  // -----------------------------------------------
  // 9. Gemini API URL
  // -----------------------------------------------

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    CONFIG.GEMINI_MODEL +
    ":generateContent?key=" +
    encodeURIComponent(apiKey);


  // -----------------------------------------------
  // 10. Gemini request
  // -----------------------------------------------

  const payload = {

    contents: [

      {

        parts: [

          {
            text: instruction
          }

        ]

      }

    ],

    generationConfig: {

      temperature: 0.4,

      responseMimeType:
        "application/json"

    }

  };


  const options = {

    method: "post",

    contentType:
      "application/json",

    payload:
      JSON.stringify(payload),

    muteHttpExceptions:
      true

  };


  const response =
    UrlFetchApp.fetch(
      url,
      options
    );


  const responseCode =
    response.getResponseCode();


  const responseText =
    response.getContentText();


  Logger.log(
    `Gemini HTTP Response: ${responseCode}`
  );


  // -----------------------------------------------
  // 11. Handle API errors
  // -----------------------------------------------

  if (responseCode !== 200) {

    Logger.log("=================================");
    Logger.log("GEMINI API ERROR");
    Logger.log("=================================");

    Logger.log(responseText);

    throw new Error(
      `Gemini API returned HTTP ${responseCode}.`
    );

  }


  // -----------------------------------------------
  // 12. Parse Gemini response
  // -----------------------------------------------

  let geminiResponse;


  try {

    geminiResponse =
      JSON.parse(responseText);

  } catch (error) {

    Logger.log(
      "Unable to parse Gemini API response."
    );

    Logger.log(responseText);

    throw new Error(
      "Gemini API returned invalid JSON."
    );

  }


  // -----------------------------------------------
  // 13. Extract generated text
  // -----------------------------------------------

  const generatedText =
    geminiResponse
      ?.candidates?.[0]
      ?.content?.parts?.[0]
      ?.text;


  if (!generatedText) {

    Logger.log(
      "Gemini response did not contain text."
    );

    Logger.log(
      JSON.stringify(
        geminiResponse,
        null,
        2
      )
    );

    throw new Error(
      "Gemini returned no generated content."
    );

  }


  // -----------------------------------------------
  // 14. Parse structured JSON
  // -----------------------------------------------

  let structured;


  try {

    structured =
      JSON.parse(generatedText);

  } catch (error) {

    Logger.log(
      "Gemini returned text that was not valid JSON."
    );

    Logger.log(
      generatedText
    );

    throw new Error(
      "Gemini structured output could not be parsed."
    );

  }


  // -----------------------------------------------
  // 15. Validate fields
  // -----------------------------------------------

  const expectedFields = [

    "video_prompt",
    "on_screen_text",
    "cta",
    "caption",
    "seo_keywords",
    "hashtags"

  ];


  expectedFields.forEach(function(field) {

    if (
      structured[field] === undefined ||
      structured[field] === null
    ) {

      throw new Error(
        `Gemini JSON is missing field: ${field}`
      );

    }

  });


  // -----------------------------------------------
  // 16. Display structured result
  // -----------------------------------------------

  Logger.log("=================================");
  Logger.log("STRUCTURED GEMINI OUTPUT");
  Logger.log("=================================");


  Logger.log(
    `Video Prompt:\n${structured.video_prompt}`
  );


  Logger.log(
    `On-Screen Text:\n${structured.on_screen_text}`
  );


  Logger.log(
    `CTA:\n${structured.cta}`
  );


  Logger.log(
    `Caption:\n${structured.caption}`
  );


  Logger.log(
    `SEO Keywords:\n${structured.seo_keywords}`
  );


  Logger.log(
    `Hashtags:\n${structured.hashtags}`
  );


  // -----------------------------------------------
  // 17. Display raw JSON
  // -----------------------------------------------

  Logger.log("=================================");
  Logger.log("RAW STRUCTURED JSON");
  Logger.log("=================================");


  Logger.log(
    JSON.stringify(
      structured,
      null,
      2
    )
  );


  // -----------------------------------------------
  // 18. Final status
  // -----------------------------------------------

  Logger.log("=================================");
  Logger.log("STEP 6E TEST COMPLETE");
  Logger.log("=================================");


  Logger.log(
    "Google Sheet was NOT modified."
  );


  Logger.log(
    "Veo was NOT called."
  );


  Logger.log(
    "Google Drive was NOT modified."
  );


  Logger.log(
    "Meta API was NOT called."
  );


  Logger.log("=================================");

}

| Stage       | Apps Script function           | What we test                 |
| ----------- | ------------------------------ | ---------------------------- |
| Step 3      | `testConnection`               | Google Sheet access          |
| Step 4      | `readContentCalendar`          | Read all 90 rows             |
| Step 4      | `getReadyContent`              | Find READY content           |
| Step 5      | `getApprovedContent`           | Find approved content        |
| Step 5      | `getTodaysContent`             | Find today's content         |
| Step 6C     | `testGeminiConnection`         | Gemini API connection        |
| Step 6D     | `testGeminiWithTodaysContent`  | Sheet → Gemini               |
| Step 6E     | `testGeminiStructuredOutput`   | Gemini → structured JSON     |
| **Step 7A** | `testGeminiImageGeneration`    | **Gemini → image**           |
| **Step 7B** | `testGeminiCarouselGeneration` | **Gemini → multiple images** |
| **Step 7C** | `testVeoConnection`            | **Veo API connection**       |
| **Step 7D** | `testVeoGeneration`            | **Gemini → Veo**             |
| Step 8      | `testDriveUpload`              | Save asset to Drive          |
| Step 9      | `testMetaConnection`           | Meta API connection          |
| Step 9      | `testInstagramPublish`         | Test Instagram publishing    |
| Step 10     | `testScheduledPublish`         | Scheduling                   |
| Step 11     | `testApprovalSafety`           | Approval protection          |
| Step 12     | `testDailyAutomation`          | Full workflow                |



