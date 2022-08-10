/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// There are several helpful constants
const columns = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

const right_headers = [
  "Name (incl. titles)",
  "Affiliation/Organisation and location",
  "Role",
  "Email",
  "Session Name",
  "Session Description",
  "Presentation Title",
  "Presentation Abstract",
  "Abstract URL",
  "Video URL",
];

const black_list_words = [
  "director",
  "department",
  "team",
  "group",
  "consortium",
  "project",
  "university",
  "institution",
  "program",
  "organization",
  "research",
  "network",
  "international",
  "medical",
  "center",
  "application",
  "organisation",
  "on behalf",
  "study",
  "genetic",
  "medicine",
  "topmed",
  "genom",
  "board",
  "institute",
  "science",
  "college",
  "accociat",
  "global",
  "develop",
  "health",
  "workplace",
  "workspace",
  "grupo",
  "committee",
  "hospital",
  "student",
  "associat",
  "clinic",
  "service",
  "society",
  "social",
  "collaborat",
  "national",
  "working",
  "contribut",
  "surgery",
  "covid",
  "candidate",
  "scient",
  "non role",
  "question",
  "answer",
  "unknown",
  "author",
  "invest",
  "general",
  "panel",
  "discus",
  "graduat",
  "mr.",
  "mrs.",
  "ms.",
  "technical",
  "leader",
  "senior",
  "other",
];

const roles = [
  "moderator",
  "speaker",
  "poster presenter",
  "panelist",
  "keynote speaker",
  "invited speaker",
  "abstract author",
];

// Loading spinner gif
const spinner = '<div class="loader" style="--b: 10px; --c:gray; width:50px; --n:20; --g:7deg"></div>';

// Function to check if a string is a valid email address
const validateEmail = (email) => {
  return String(email)
    .toLowerCase()
    .match(
      /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/
    );
};

// Method to get duplicates from an array
Array.prototype.getDuplicates = function () {
  let duplicates = {};
  for (let i = 0; i < this.length; i++) {
    if (duplicates.hasOwnProperty(this[i])) {
      duplicates[this[i]].push(i);
    } else if (this.lastIndexOf(this[i]) !== i) {
      duplicates[this[i]] = [i];
    }
  }
  return duplicates;
};

// Function to check if a URL is valid
function isValidHttpUrl(string) {
  let url;
  try {
    url = new URL(string);
  } catch (_) {
    return false;
  }
  return url.protocol === "http:" || url.protocol === "https:";
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Define the main buttons in the taskpane
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-welcome").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("refresh").onclick = refresh;
    document.getElementById("refresh2").onclick = refresh;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Main start of the script.
       */
      // Hide the welcome page and show the app page.
      document.getElementById("app-welcome").style.display = "none";
      document.getElementById("app-body").style.display = "flex";
      document.getElementById("loader").style.display = "flex";

      // Load data from the Excel sheet.
      let currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      let range = currentWorksheet.getUsedRange();
      range.load("rowCount, columnCount, values");
      await context.sync();
      let data = range.values;
      let rowCount = range.rowCount;
      let columnCount = range.columnCount;

      // If there is too much data, show an error message.
      if (data === null) {
        document.getElementById("loader").style.display = "none";
        document.getElementById("null-body").style.display = "flex";
        return;
      }

      // Run the tests and show the results.
      await test_headers(data);
      await test_data_range(columnCount, rowCount);
      await test_author_names(data, rowCount);
      await test_author_roles(data, rowCount);
      await test_author_emails(data, rowCount);
      await test_session_names(data, rowCount);
      await test_titles(data, rowCount);
      await test_urls(data, rowCount);
      await test_urls(data, rowCount);
      await test_duplicates(data, rowCount);
      await test_poster_sessions(data, rowCount);

      // If there is no error, show the success message.
      if (document.getElementById("errors-msg").style.display === "none") {
        document.getElementById("loader").style.display = "none";
        document.getElementById("no-errors-msg").style.display = "inline-grid";
      }
    });
  } catch (error) {
    console.error(error);
  }
}

export async function refresh() {
  try {
    await Excel.run(async (context) => {
      /**
       * Restart the tests
       */
      // Hide the results of the previous run
      document.getElementById("errors-msg").style.display = "none";
      document.getElementById("warning-msg").style.display = "none";
      document.getElementById("no-errors-msg").style.display = "none";
      // Clear the errors message
      document.getElementById("header-errors").innerHTML = "";
      document.getElementById("data-range-errors").innerHTML = "";
      document.getElementById("author-name-errors").innerHTML = "";
      document.getElementById("author-role-errors").innerHTML = "";
      document.getElementById("author-email-errors").innerHTML = "";
      document.getElementById("session-names-errors").innerHTML = "";
      document.getElementById("title-errors").innerHTML = "";
      document.getElementById("url-errors").innerHTML = "";
      document.getElementById("poster_session_errors").innerHTML = "";
      document.getElementById("duplicate-errors").innerHTML = "";
      // Clear the warnings message
      document.getElementById("url-warnings").innerHTML = "";
      document.getElementById("title-warnings").innerHTML = "";
      document.getElementById("duplicate-warnings").innerHTML = "";
      // Run the tests again
      await run();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function test_headers(data) {
  try {
    await Excel.run(async (context) => {
      /**
       * This test checks the wokrsheet headers
       */
      // Set the spinner while the test is running
      document.getElementById("header-errors").innerHTML = spinner;

      // Set the useful variables
      let headers = data[0];

      let errors = [];
      for (let i = 0; i < 10; i++) {
        // Check if the header match the expected header
        if (headers[i] !== right_headers[i]) {
          let url_cell_address = `${columns[i]}1`;
          errors.push([url_cell_address, `Header must be "${right_headers[i]}"`]);
        }
      }

      //If there are errors, create a card with the errors
      document.getElementById("header-errors").innerHTML = "";
      if (errors.length > 0) {
        await error_card_creator("header-errors", "HEADERS", errors);
        let card = document.getElementById("header-errors").getElementsByClassName("card")[0];
        card.innerHTML += `<div class="accept-button-container">
                             <span role="button" id="set-headers" class="accept-button">Set Right Headers</span>
                           </div>`;
        document.getElementById("set-headers").onclick = set_headers;
      }
    });
  } catch (error) {
    console.error(error);
  }
}

export async function test_data_range(columnCount, rowCount) {
  try {
    await Excel.run(async (context) => {
      /**
       * This test checks the used range of the worksheet
       * If there is any information outside the operating range, the errors will appear
       */
      // Set the spinner while the test is running
      document.getElementById("data-range-errors").innerHTML = spinner;

      // Get current worksheet
      let currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

      let errors = [];
      if (columnCount > 10) {
        // If there are some data out of the operating range, create an error message
        let error_range = currentWorksheet.getRange(`K1:${columns[columnCount - 1]}${rowCount}`);
        error_range.load("values");
        console.log("loading");
        await context.sync();
        console.log(error_range.values);
        let values = error_range.values;
        for (let i = 0; i < values.length; i++) {
          for (let j = 0; j < values[i].length; j++) {
            if (values[i][j] !== "") {
              errors.push([`${columns[j + 10]}${i + 1}`, `Data out of operating range`]);
            }
          }
        }
        console.log(errors);
      }

      //If there are errors, create a card with the errors
      document.getElementById("data-range-errors").innerHTML = "";
      if (errors.length > 0) {
        await error_card_creator("data-range-errors", "RANGE", errors);
        let card = document.getElementById("data-range-errors").getElementsByClassName("card")[0];
        card.innerHTML += `<div class="accept-button-container">
                             <span role="button" id="clear-range" class="accept-button">Clear all</span>
                           </div>`;
        document.getElementById("clear-range").onclick = clear_range;
      }
    });
  } catch (error) {
    console.error(error);
  }
}

export async function test_author_names(data, rowCount) {
  try {
    await Excel.run(async (context) => {
      /**
       * This test checks the author names
       */
      // Set the spinner while the test is running
      document.getElementById("author-name-errors").innerHTML = spinner;

      // Set useful variables
      let author_index = 0;
      let regex = /\d+/gm;

      let errors = [];
      for (let index = 1; index < rowCount; index++) {
        let author_name_address = `A${index + 1}`;
        let row = data[index];
        let author_name = row[author_index].toString().trim();

        //Test if the author name cell has a number in it
        if (regex.exec(author_name) !== null) {
          errors.push([author_name_address, `Author name "${author_name}" contains a number`]);
        }
        //Test if the author name cell is empty
        else if (author_name === "") {
          errors.push([author_name_address, `Author name is empty`]);
        }
        //Test if the author name cell don't have spaces in it
        else if (!author_name.includes(" ")) {
          errors.push([author_name_address, `Author name "${author_name}" doesn't contain spaces`]);
        }
        //Test if the author name cell is too short
        else if (author_name.length < 5) {
          errors.push([author_name_address, `Author name "${author_name}" is too short`]);
        }
        //Test if the author name cell is too long
        else if (author_name.length > 50) {
          errors.push([author_name_address, `Author name is too long`]);
        }
        //Test if the author name cell does not contain a black list word
        else {
          let lower_url_cell_value = author_name.toLowerCase();
          for (let j = 0; j < black_list_words.length; j++) {
            if (lower_url_cell_value.includes(black_list_words[j])) {
              errors.push([
                author_name_address,
                `Author name "${author_name}" contains a word from a blacklist: "${black_list_words[j]}"`,
              ]);
              break;
            }
          }
        }
      }
      //If there are errors, create a card with the errors
      document.getElementById("author-name-errors").innerHTML = "";
      if (errors.length > 0) {
        await error_card_creator("author-name-errors", "AUTHORS", errors);
      }
    });
  } catch (error) {
    console.error(error);
  }
}

export async function test_author_roles(data, rowCount) {
  try {
    await Excel.run(async (context) => {
      // Set the spinner while the test is running
      document.getElementById("author-role-errors").innerHTML = spinner;

      // Set useful variables
      let role_index = 2;

      let errors = [];
      for (let index = 1; index < rowCount; index++) {
        let author_role_address = `C${index + 1}`;
        let row = data[index];
        let author_role = row[role_index].toString().trim();
        let lower_author_role = author_role.toLowerCase();
        // Test if the author role cell is empty
        if (author_role === "") {
          errors.push([author_role_address, `Author role is empty`]);
        }
        // Test if the author role cell is invalid
        else if (!roles.includes(lower_author_role)) {
          errors.push([author_role_address, `Author role "${author_role}" is invalid`]);
        }
      }
      //If there are errors, create a card with the errors
      document.getElementById("author-role-errors").innerHTML = "";
      if (errors.length > 0) {
        await error_card_creator("author-role-errors", "ROLES", errors);
      }
    });
  } catch (error) {
    console.error(error);
  }
}

export async function test_author_emails(data, rowCount) {
  try {
    await Excel.run(async (context) => {
      // Set the spinner while the test is running
      document.getElementById("author-email-errors").innerHTML = spinner;

      // Set useful variables
      let email_index = 3;

      let errors = [];
      for (let index = 1; index < rowCount; index++) {
        let email_address = `D${index + 1}`;
        let row = data[index];
        let email = row[email_index].toString().trim();
        // Test if the author email cell is not empty
        if (email !== "") {
          if (!validateEmail(email)) {
            errors.push([email_address, `Author email "${email}" is invalid`]);
          }
        }
      }
      //If there are errors, create a card with the errors
      document.getElementById("author-email-errors").innerHTML = "";
      if (errors.length > 0) {
        await error_card_creator("author-email-errors", "EMAILS", errors);
      }
    });
  } catch (error) {
    console.error(error);
  }
}

export async function test_session_names(data, rowCount) {
  try {
    await Excel.run(async (context) => {
      // Set the spinner while the test is running
      document.getElementById("session-names-errors").innerHTML = spinner;

      // Set useful variables
      let session_name_index = 4;

      let errors = [];
      for (let index = 1; index < rowCount; index++) {
        let session_name_address = `E${index + 1}`;
        let row = data[index];
        let session_name = row[session_name_index].toString().trim();
        // Test if the session name cell is empty
        if (session_name === "") {
          errors.push([session_name_address, `Session name is empty`]);
        }
      }
      //If there are errors, create a card with the errors
      document.getElementById("session-names-errors").innerHTML = "";
      if (errors.length > 0) {
        await error_card_creator("session-names-errors", "SESSION", errors);
      }
    });
  } catch (error) {
    console.error(error);
  }
}

export async function test_titles(data, rowCount) {
  try {
    await Excel.run(async (context) => {
      // Set the spinner while the test is running
      document.getElementById("title-errors").innerHTML = spinner;
      document.getElementById("title-warnings").innerHTML = spinner;

      // Set useful variables
      let role_index = 2;
      let title_index = 6;

      let errors = [];
      let warns = [];
      for (let index = 1; index < rowCount; index++) {
        let title_address = `G${index + 1}`;
        let row = data[index];
        let role = row[role_index].toString().trim();
        let title = row[title_index].toString().trim();
        // Test if the title cell is not empty
        if (role !== "Moderator" && title === "") {
          errors.push([title_address, `Presentation title is empty`]);
        }
        // Test if the title cell is too short
        else if (role !== "Moderator" && title.length <= 5) {
          warns.push([title_address, `Presentation title is too short`]);
        }
        // Test if the title cell is empty for a moderator
        else if (role === "Moderator" && title !== "") {
          errors.push([title_address, "Presentation title should be empty (Moderator)"]);
        }
      }
      //If there are errors, create a card with the errors
      document.getElementById("title-errors").innerHTML = "";
      if (errors.length > 0) {
        await error_card_creator("title-errors", "TITLE", errors);
      }
      //If there are warnings, create a card with the warnings
      document.getElementById("title-warnings").innerHTML = "";
      if (warns.length > 0) {
        await warning_card_creator("title-warnings", "TITLE", warns);
      }
    });
  } catch (error) {
    console.error(error);
  }
}

export async function test_urls(data, rowCount) {
  try {
    await Excel.run(async (context) => {
      // Set the spinner while the test is running
      document.getElementById("url-errors").innerHTML = spinner;
      document.getElementById("url-warnings").innerHTML = spinner;

      // Set useful variables
      let role_index = 2;
      let url_index = 8;
      let video_url_index = 9;

      let errors = [];
      let warns = [];
      for (let index = 1; index < rowCount; index++) {
        let url_address = `I${index + 1}`;
        let video_url_address = `J${index + 1}`;
        let row = data[index];
        let role = row[role_index].toString().trim();
        let url = row[url_index].toString().trim();
        let video_url = row[video_url_index].toString().trim();

        // Test if the abstract url is valid for any role except moderator
        if (role !== "Moderator") {
          //check if the url not empty
          if (url === "") {
            errors.push([url_address, `Presentation URL is empty`]);
          }
          // Check if the url is valid
          else if (!isValidHttpUrl(url)) {
            errors.push([url_address, `Presentation URL is invalid`]);
          }
          // Check if the url do not lead to GitHub PDF viewer
          else if (url.includes("github")) {
            errors.push([url_address, `Presentation URL leads to the github PDF viewer`]);
          }
        }
        // Test if the abstract url is valid for a moderator
        else if (role === "Moderator" && url !== "") {
          if (!isValidHttpUrl(url)) {
            errors.push([url_address, `Presentation URL is invalid`]);
          }
          warns.push([url_address, `Double check if the moderator needs the URL`]);
        }

        // Test if the video url is valid for any role except moderator
        if (role !== "Moderator" && video_url !== "") {
          if (!isValidHttpUrl(video_url)) {
            errors.push([video_url_address, `Video URL is invalid`]);
          }
        }
        // Test if the video url is valid for a moderator
        else if (role === "Moderator" && video_url !== "") {
          errors.push([video_url_address, `Video URL must be empty for Moderator`]);
        }
      }
      //If there are errors, create a card with the errors
      document.getElementById("url-errors").innerHTML = "";
      if (errors.length > 0) {
        await error_card_creator("title-errors", "URL", errors);
      }
      //If there are warnings, create a card with the warnings
      document.getElementById("url-warnings").innerHTML = "";
      if (warns.length > 0) {
        await warning_card_creator("url-warnings", "URL", warns);
      }
    });
  } catch (error) {
    console.error(error);
  }
}

export async function test_duplicates(data, rowCount) {
  try {
    await Excel.run(async (context) => {
      // Set the spinner while the test is running
      document.getElementById("duplicate-errors").innerHTML = spinner;
      document.getElementById("duplicate-warnings").innerHTML = spinner;
      // Get the current worksheet
      let currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

      let i = rowCount;
      let concat_cell_errors = currentWorksheet.getRange("X2");
      // Concatenate Name, Affiliation, Role, Session Name and Presentation Title
      concat_cell_errors.formulas = [[`=CONCATENATE(A2:A${i}, B2:B${i}, C2:C${i}, E2:E${i}, G2:G${i})`]];
      let concat_cell_warnings = currentWorksheet.getRange("Y2");
      // Concatenate Name and Presentation Title
      concat_cell_warnings.formulas = [[`=CONCATENATE(A2:A${i}, G2:G${i})`]];
      await context.sync();
      let concat_range_errors = currentWorksheet.getRange(`X2:X${i}`);
      concat_range_errors.load("values");
      let concat_range_warnings = currentWorksheet.getRange(`Y2:Y${i}`);
      concat_range_warnings.load("values");
      await context.sync();
      // Search for error duplicates in the concatenated cells
      let errors_res = concat_range_errors.values.flat();
      let errors_dup = errors_res.getDuplicates();
      let errors_keys = Object.keys(errors_dup);
      let errors = [];
      for (const key of errors_keys) {
        let temp_errors = [];
        for (const index of errors_dup[key]) {
          temp_errors.push([`A${index + 2}`, `${data[index + 1][0]} | ${data[index + 1][2]}`]);
        }
        if (temp_errors.length > 1) {
          errors.push(temp_errors);
        }
      }
      // Search for warning duplicates in the concatenated cells
      let warnings_res = concat_range_warnings.values.flat();
      let warnings_dup = warnings_res.getDuplicates();
      let warnings_keys = Object.keys(warnings_dup);
      let main_roles = ["Poster Presenter", "Speaker", "Invited Speaker", "Keynote Speaker"];
      let warnings = [];
      for (const key of warnings_keys) {
        let temp_warnings = [];
        let is_main_roles = [];
        for (const index of warnings_dup[key]) {
          if (main_roles.includes(data[index + 1][2])) {
            is_main_roles.push(true);
          }
        }
        console.log(main_roles);
        console.log(is_main_roles);
        if (is_main_roles.length !== 0 && is_main_roles.every((element) => element === true)) {
          for (const index of warnings_dup[key]) {
            temp_warnings.push([`A${index + 2}`, `${data[index + 1][0]} | ${data[index + 1][2]}`]);
          }
          if (temp_warnings.length > 1) {
            warnings.push(temp_warnings);
          }
        }
      }
      concat_range_errors.clear();
      concat_range_warnings.clear();
      await context.sync();
      //If there are errors, create a card with the errors
      document.getElementById("duplicate-errors").innerHTML = "";
      if (errors.length > 0) {
        await duplicate_error_card_creator("duplicate-errors", "DUPLICATE", errors);
        let card = document.getElementById("duplicate-errors").getElementsByClassName("card")[0];
        card.innerHTML += `<div class="accept-button-container">
                             <span role="button"
                             id="remove-full-duplicates"
                             class="accept-button"
                             title="Removes fully matched rows">Remove full duplicates
                             <i class="ms-Icon ms-Icon--Warning ms-font-xxl" title="This action cannot be undone via Ctrl+Z"></i>
                             </span>
                           </div>
                           <div class="accept-button-container">
                             <span 
                             role="button" 
                             id="remove-presented-duplicates" 
                             class="accept-button" 
                             title="Removes rows that are matched by:\nName (incl. titles)\nAffiliation/Organisation and location\nRole\nSession Name\nPresentation Title">Remove presented duplicates
                             <i class="ms-Icon ms-Icon--Warning ms-font-xxl" title="This action cannot be undone via Ctrl+Z"></i>
                             </span>
                           </div>`;
        document.getElementById("remove-full-duplicates").onclick = remove_full_duplicates;
        document.getElementById("remove-presented-duplicates").onclick = remove_presented_duplicates;
      }
      //If there are warnings, create a card with the warnings
      document.getElementById("duplicate-warnings").innerHTML = "";
      if (warnings.length > 0) {
        await duplicate_warning_card_creator("duplicate-warnings", "SEVERAL MAIN ROLES BY PRESENTATION", warnings);
      }
    });
  } catch (error) {
    console.error(error);
  }
}

export async function test_poster_sessions(data, rowCount) {
  try {
    await Excel.run(async (context) => {
      // Set the spinner while the test is running
      document.getElementById("poster_session_errors").innerHTML = spinner;

      // Set useful variables
      let role_index = 2;
      let session_index = 4;
      let non_poster_roles = ["speaker", "invited speaker", "keynote speaker"];
      console.log("Starting test_poster_sessions");

      let poster_sessions = new Set();
      for (let i = 1; i < rowCount; i++) {
        if (data[i][role_index] === "Poster Presenter") {
          poster_sessions.add(data[i][session_index]);
        }
      }
      console.log(poster_sessions);

      let errors = [];
      for (let index = 1; index < rowCount; index++) {
        let author_cell_address = `A${index + 1}`;
        let row = data[index];
        let role = row[role_index].toString().trim();
        let session_cell_value = row[session_index].toString().trim();
        // Test if the title cell is not empty
        if (poster_sessions.has(session_cell_value)) {
          if (non_poster_roles.includes(role.toLowerCase())) {
            errors.push([author_cell_address, `${role} in a poster session`]);
          }
        }
      }
      //If there are errors, create a card with the errors
      document.getElementById("poster_session_errors").innerHTML = "";
      if (errors.length > 0) {
        await error_card_creator("poster_session_errors", "WRONG ROLE IN POSTER SESSION", errors);
      }
    });
  } catch (error) {
    console.error(error);
  }
}

export async function set_headers() {
  try {
    await Excel.run(async (context) => {
      // Get the current worksheet
      let currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      let range = currentWorksheet.getRange("A1:J1");
      // Set the headers
      range.values = [right_headers];
      await context.sync();
      document.getElementById("header-errors").innerHTML = "";
    });
  } catch (error) {
    console.error(error);
  }
}

export async function clear_range() {
  try {
    await Excel.run(async (context) => {
      // Get the current worksheet
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      // Clear the range
      let range = sheet.getRange("K:Z");
      range.clear();
      await context.sync();
      // Restart the test
      await test_data_range();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function remove_full_duplicates() {
  try {
    await Excel.run(async (context) => {
      console.log("In it");
      let currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      let range = currentWorksheet.getUsedRange();

      let deleteResult = range.removeDuplicates([0, 1, 2, 3, 4, 5, 6, 7, 8, 9], true);
      deleteResult.load();
      await context.sync();
      console.log(deleteResult.removed + " entries with duplicate names removed.");
      console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");

      // Restart the tests
      await refresh();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function remove_presented_duplicates() {
  try {
    await Excel.run(async (context) => {
      let currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      let range = currentWorksheet.getUsedRange();

      let deleteResult = range.removeDuplicates([0, 1, 2, 4, 6], true);
      deleteResult.load();
      await context.sync();

      console.log(deleteResult.removed + " entries with duplicate names removed.");
      console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");

      // Restart the test
      await refresh();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function error_card_creator(card_id, card_title, card_content) {
  try {
    await Excel.run(async (context) => {
      document.getElementById("errors-msg").style.display = "inline-grid";
      document.getElementById("loader").style.display = "none";
      let card = document.getElementById(card_id);
      card.style.display = "block";
      card.innerHTML = `<div class="card">
                          <div class="card-label">${card_title}
                          <sup class="error-label" title="Need to fix">ERROR</sup></div>
                          <div class="card-container"></div>`;
      let card_container = card.getElementsByClassName("card-container")[0];
      for (let i = 0; i < card_content.length; i++) {
        let on_click = `set_active('${card_content[i][0]}')`;
        console.log(on_click);
        card_container.innerHTML += `<div class="container-row">
                                       <div class="goto-cell-button" 
                                       onclick="${on_click}" 
                                       title="Go to cell">${card_content[i][0]}</div>
                                       <div class="row-explanation">${card_content[i][1]}</div>
                                     </div>`;
      }
      card.innerHTML += `</div></div>`;
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function duplicate_error_card_creator(card_id, card_title, card_content) {
  try {
    await Excel.run(async (context) => {
      document.getElementById("errors-msg").style.display = "inline-grid";
      document.getElementById("loader").style.display = "none";
      let card = document.getElementById(card_id);
      card.style.display = "block";
      let container = `<div class="card-container"></div>`;
      card.innerHTML = `<div class="card">
                          <div class="card-label">${card_title}
                          <sup class="error-label" title="Need to fix">ERROR</sup></div>
                          ${container.repeat(card_content.length)}`;
      for (let i = 0; i < card_content.length; i++) {
        let card_container = card.getElementsByClassName("card-container")[i];
        for (let j = 0; j < card_content[i].length; j++) {
          let deleted = j === 0 ? "" : '<div class="duplicate"></div>';
          let on_click = `set_active('${card_content[i][j][0]}')`;
          card_container.innerHTML += `<div class="container-row">
                                         ${deleted}
                                         <div class="goto-cell-button" 
                                         onclick="${on_click}" 
                                         title="Go to cell">${card_content[i][j][0]}</div>
                                         <div class="row-explanation duplicates">${card_content[i][j][1]}</div>
                                       </div>`;
        }
      }
      card.innerHTML += `</div></div>`;
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function duplicate_warning_card_creator(card_id, card_title, card_content) {
  try {
    await Excel.run(async (context) => {
      document.getElementById("warning-msg").style.display = "inline-grid";
      document.getElementById("loader").style.display = "none";
      let card = document.getElementById(card_id);
      card.style.display = "block";
      let container = `<div class="card-container"></div>`;
      card.innerHTML = `<div class="card">
                          <div class="card-label">${card_title}
                          <sup class="warning-label" title="Need to double-check">WARNING</sup></div>
                          ${container.repeat(card_content.length)}`;
      for (let i = 0; i < card_content.length; i++) {
        let card_container = card.getElementsByClassName("card-container")[i];
        for (let j = 0; j < card_content[i].length; j++) {
          let on_click = `set_active('${card_content[i][j][0]}')`;
          card_container.innerHTML += `<div class="container-row">
                                         <div class="goto-cell-button" 
                                         onclick="${on_click}" 
                                         title="Go to cell">${card_content[i][j][0]}</div>
                                         <div class="row-explanation duplicates">${card_content[i][j][1]}</div>
                                       </div>`;
        }
      }
      card.innerHTML += `</div></div>`;
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function warning_card_creator(card_id, card_title, card_content) {
  try {
    await Excel.run(async (context) => {
      document.getElementById("warning-msg").style.display = "inline-grid";
      document.getElementById("loader").style.display = "none";
      let card = document.getElementById(card_id);
      card.style.display = "block";
      card.innerHTML = `<div class="card">
                          <div class="card-label">${card_title}
                          <sup class="warning-label" title="Need to double-check">WARNING</sup></div>
                          <div class="card-container"></div>`;
      let card_container = card.getElementsByClassName("card-container")[0];
      for (let i = 0; i < card_content.length; i++) {
        let on_click = `set_active('${card_content[i][0]}')`;
        console.log(on_click);
        card_container.innerHTML += `<div class="container-row">
                                       <div class="goto-cell-button" onclick="${on_click}" title="Go to cell">${card_content[i][0]}</div>
                                       <div class="row-explanation">${card_content[i][1]}</div>
                                     </div>`;
      }
      card.innerHTML += `</div></div>`;
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
