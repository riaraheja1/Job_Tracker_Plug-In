/**
 * JOB TRACKER AI FILL — Google Apps Script + Gemini
 * ---------------------------------------------------
 * Adds a "Job Tracker" menu to your sheet with an "Add Job from Link" popup.
 * Paste a job posting URL — that's it — and it fills a full row automatically.
 *
 * ===================== 1. CUSTOMIZE THIS SECTION =====================
 * Edit these to match your sheet. Nothing else in the file needs to change
 * for normal customization (renamed/added/removed columns, new fixed values, etc).
 */

const CONFIG = {
  // Exact name of the tab your tracker lives on
  SHEET_NAME: 'August',

  // Gemini model to use (Flash-Lite is free-tier, fast, and built for
  // exactly this kind of extraction task)
  GEMINI_MODEL: 'gemini-3.5-flash-lite',

  // Free-text columns the AI extracts straight from the job posting
  AI_TEXT_COLUMNS: ['Company', 'Role', 'Location', 'Salary'],

  // Dropdown column(s) the AI extracts, constrained to your sheet's actual
  // dropdown options (read live from Data Validation, so it stays in sync
  // if you edit the dropdown list later)
  AI_DROPDOWN_COLUMNS: ['Type'],

  // Columns always set to the same fixed value on every new row — no AI,
  // no asking. Edit the values on the right if your workflow changes.
  // NOTE: these must exactly match text in your sheet's dropdown lists.
  FIXED_VALUE_COLUMNS: {
    'Application Status': 'Submitted',
    'Technical Interview': 'No Interview',
    'Behavioral Interview': 'No Interview',
    'Application Decision': 'Waiting'
  },

  // Platform Found is guessed from the URL's domain, not the AI. Add more
  // patterns here if you apply through other boards. First match wins;
  // DEFAULT is used when nothing matches.
  PLATFORM_URL_PATTERNS: [
    { match: 'linkedin.com', platform: 'LinkedIn' },
    { match: 'workatastartup.com', platform: 'YC' },
    { match: 'ycombinator.com', platform: 'YC' },
    { match: 'joinhandshake.com', platform: 'Handshake' },
    { match: 'handshake.com', platform: 'Handshake' }
  ],
  PLATFORM_DEFAULT: 'Company Website',

  // Columns this tool never touches — always left blank
  SKIP_COLUMNS: ['Notes', 'Username', 'Password'],

  // Optional extra formatting instructions per AI column
  COLUMN_INSTRUCTIONS: {
    'Salary': 'Format as a range like "$90,000-$110,000" or hourly like "$25/hr". If not listed on the posting, write "Not listed".',
    'Location': 'Format as "City, State". If the posting doesn\'t give a specific location, write "Not listed".'
  }
};

/** ===================== 2. MENU + API KEY SETUP ===================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Job Tracker')
    .addItem('Add Job from Link...', 'showJobDialog')
    .addSeparator()
    .addItem('Set Gemini API Key...', 'promptForApiKey')
    .addToUi();
}

function promptForApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Gemini API Key',
    'Paste your free Gemini API key (from aistudio.google.com/apikey). It is stored privately in this sheet\'s script settings, not in the code.',
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() === ui.Button.OK) {
    const key = result.getResponseText().trim();
    if (key) {
      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
      ui.alert('Saved. You can now use "Add Job from Link".');
    }
  }
}

function showJobDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setWidth(420)
    .setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Job from Link');
}

/** ===================== 3. MAIN ENTRY POINT (called from Sidebar.html) ===================== */

function processJobEntry(form) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('No Gemini API key set. Use Job Tracker > Set Gemini API Key first.');
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Could not find a tab named "' + CONFIG.SHEET_NAME + '". Check CONFIG.SHEET_NAME in Code.gs.');

  const headerMap = getHeaderMap(sheet);

  const jobUrl = (form.jobUrl || '').trim();
  let jobText = (form.jobText || '').trim();
  if (!jobText) {
    if (!jobUrl) throw new Error('Provide a job posting URL (or paste the description text).');
    jobText = fetchJobText(jobUrl);
  }

  // Pull live dropdown options for AI dropdown columns so extraction is
  // constrained to values that will actually be valid in your sheet.
  const dropdownOptions = {};
  CONFIG.AI_DROPDOWN_COLUMNS.forEach(function (col) {
    if (headerMap[col] !== undefined) {
      const opts = getDropdownOptions(sheet, headerMap[col]);
      if (opts) dropdownOptions[col] = opts;
    }
  });

  const aiColumns = CONFIG.AI_TEXT_COLUMNS.concat(CONFIG.AI_DROPDOWN_COLUMNS);
  const aiResult = callGeminiExtract(apiKey, jobText, aiColumns, dropdownOptions);

  const rowData = {};
  const warnings = [];
  CONFIG.AI_TEXT_COLUMNS.forEach(function (col) { rowData[col] = aiResult[col] || ''; });
  CONFIG.AI_DROPDOWN_COLUMNS.forEach(function (col) {
    const matched = matchToDropdown(aiResult[col], dropdownOptions[col]);
    rowData[col] = matched;
    if (!matched && aiResult[col]) warnings.push(col + ' ("' + aiResult[col] + '" didn\'t match a dropdown option — left blank)');
  });

  Object.keys(CONFIG.FIXED_VALUE_COLUMNS).forEach(function (col) {
    const fixedVal = CONFIG.FIXED_VALUE_COLUMNS[col];
    if (headerMap[col] !== undefined) {
      const opts = getDropdownOptions(sheet, headerMap[col]);
      if (opts) {
        const matched = matchToDropdown(fixedVal, opts);
        rowData[col] = matched;
        if (!matched) warnings.push(col + ' (fixed value "' + fixedVal + '" isn\'t one of your dropdown options — left blank; update FIXED_VALUE_COLUMNS in Code.gs)');
        return;
      }
    }
    rowData[col] = fixedVal;
  });

  rowData['Date Applied'] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  if (headerMap['Platform Found'] !== undefined) {
    const platformOptions = getDropdownOptions(sheet, headerMap['Platform Found']);
    rowData['Platform Found'] = inferPlatform(jobUrl, platformOptions);
  }

  appendRow(sheet, headerMap, rowData);

  let msg = 'Added: ' + (rowData['Company'] || 'New row') + ' — ' + (rowData['Role'] || '');
  if (warnings.length) msg += '. Note: ' + warnings.join('; ') + '.';
  return msg;
}

/** ===================== 4. HELPERS ===================== */

function getHeaderMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach(function (name, i) {
    if (name) map[name.toString().trim()] = i + 1; // 1-indexed column
  });
  return map;
}

// Reads the dropdown list from row 2 of a column (apply Data Validation to
// the whole column, not just one cell, so this stays reliable as rows grow).
function getDropdownOptions(sheet, colIndex) {
  const cell = sheet.getRange(2, colIndex);
  const rule = cell.getDataValidation();
  if (!rule) return null;
  const criteria = rule.getCriteriaType();
  if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    return rule.getCriteriaValues()[0];
  }
  if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    const range = rule.getCriteriaValues()[0];
    return range.getValues().flat().filter(String);
  }
  return null;
}

function inferPlatform(url, dropdownOptions) {
  const lowerUrl = (url || '').toLowerCase();
  let guess = CONFIG.PLATFORM_DEFAULT;
  for (let i = 0; i < CONFIG.PLATFORM_URL_PATTERNS.length; i++) {
    const rule = CONFIG.PLATFORM_URL_PATTERNS[i];
    if (lowerUrl.indexOf(rule.match) !== -1) {
      guess = rule.platform;
      break;
    }
  }
  return matchToDropdown(guess, dropdownOptions);
}

// Normalizes text for loose comparison: lowercase, strip everything but letters/numbers
function normalize(s) {
  return (s || '').toString().toLowerCase().replace(/[^a-z0-9]/g, '');
}

// Common wording the AI (or a URL guess) might use that doesn't literally match
// your dropdown text. Add more pairs here as you notice mismatches — the key
// is the normalized alias, the value is the normalized target it should map to.
const DROPDOWN_SYNONYMS = {
  'onsite': 'inperson', 'office': 'inperson', 'inoffice': 'inperson', 'onsiteoffice': 'inperson',
  'wfh': 'remote', 'workfromhome': 'remote', 'fullyremote': 'remote',
  'companysite': 'companywebsite', 'careerspage': 'companywebsite', 'careerssite': 'companywebsite'
};

// Snaps a value to the exact text of one of your dropdown options (case/punctuation
// insensitive), using DROPDOWN_SYNONYMS as a fallback. If nothing matches, returns
// '' (blank) rather than writing text that would violate Data Validation.
function matchToDropdown(value, options) {
  if (!options || options.length === 0) return value || '';
  const normVal = normalize(value);
  let match = options.find(function (o) { return normalize(o) === normVal; });
  if (match) return match;

  const alias = DROPDOWN_SYNONYMS[normVal];
  if (alias) {
    match = options.find(function (o) { return normalize(o) === alias; });
    if (match) return match;
  }
  return '';
}

// Job boards known to require login and/or render via JavaScript, so a plain
// fetch only returns the page shell, not the actual posting. Add more domains
// here as you run into them.
const BLOCKED_FETCH_DOMAINS = ['joinhandshake.com', 'handshake.com', 'linkedin.com'];

// Words that show up in almost any real job posting. If none of these appear,
// the fetched text is probably a login page, nav shell, or JS placeholder —
// not the actual description — so we ask for pasted text instead of silently
// feeding junk to the AI.
const JOB_SIGNAL_WORDS = ['responsibilit', 'qualificat', 'requirement', 'salary', 'compensation', 'experience', 'apply'];

// Workday career sites are structured as:
//   https://{tenant}.{datacenter}.myworkdayjobs.com/{site}/job/{location}/{Job-Title}_{ReqID}
// Their frontend loads job content from a JSON endpoint at:
//   https://{tenant}.{datacenter}.myworkdayjobs.com/wday/cxs/{tenant}/{site}/job/{location}/{Job-Title}_{ReqID}
// That endpoint is public (no login) since it's what powers the page itself,
// so we call it directly instead of the JS-rendered HTML page.
function tryWorkdayApi(url) {
  const match = url.match(/^https?:\/\/([a-z0-9-]+)\.([a-z0-9]+)\.myworkdayjobs\.com\/([^\/]+)\/job\/([^?#]+)/i);
  if (!match) return null;
  const tenant = match[1], dc = match[2], site = match[3], jobPath = match[4];
  const apiUrl = 'https://' + tenant + '.' + dc + '.myworkdayjobs.com/wday/cxs/' + tenant + '/' + site + '/job/' + jobPath;

  try {
    const resp = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    const data = JSON.parse(resp.getContentText());
    const info = data.jobPostingInfo;
    if (!info) return null;

    const descText = (info.jobDescription || '')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

    const parts = [info.title || '', info.location || ''];
    if (info.additionalLocations) parts.push(info.additionalLocations.join(', '));
    parts.push(descText);
    return parts.filter(String).join('\n');
  } catch (e) {
    return null;
  }
}

function fetchJobText(url) {
  const lowerUrl = url.toLowerCase();

  // Workday renders via JavaScript too, but its own frontend pulls job data
  // from a public JSON API. Use that directly instead of the empty JS shell.
  if (lowerUrl.indexOf('myworkdayjobs.com') !== -1) {
    const workdayText = tryWorkdayApi(url);
    if (workdayText) return workdayText.slice(0, 15000);
    throw new Error('Couldn\'t read this Workday posting via its data API (the URL format may differ from usual). Paste the job description text into the "Or paste job description text" box instead.');
  }

  const blockedMatch = BLOCKED_FETCH_DOMAINS.find(function (d) { return lowerUrl.indexOf(d) !== -1; });
  if (blockedMatch) {
    throw new Error('"' + blockedMatch + '" requires login and/or loads via JavaScript, so this can\'t fetch the real posting from a link. Copy the job description text and paste it into the "Or paste job description text" box instead.');
  }

  let response;
  try {
    response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  } catch (e) {
    throw new Error('Could not fetch that URL. Some sites block automated fetches — paste the job description text into the "Or paste job description text" box instead.');
  }
  const html = response.getContentText();
  const text = html
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const lowerText = text.toLowerCase();
  const hasSignal = JOB_SIGNAL_WORDS.some(function (w) { return lowerText.indexOf(w) !== -1; });

  if (text.length < 200 || !hasSignal) {
    throw new Error('That page didn\'t return readable job-posting text (likely blocked, requires login, or loads via JavaScript). Paste the job description text into the "Or paste job description text" box instead.');
  }
  return text.slice(0, 15000); // keep prompt a reasonable size
}

function callGeminiExtract(apiKey, jobText, aiColumns, dropdownOptions) {
  const properties = {};
  const required = [];

  aiColumns.forEach(function (col) {
    let desc = 'Extract the ' + col + ' from the job posting.';
    if (CONFIG.COLUMN_INSTRUCTIONS[col]) desc += ' ' + CONFIG.COLUMN_INSTRUCTIONS[col];
    if (dropdownOptions[col]) {
      desc += ' You MUST choose exactly one of these values: ' + dropdownOptions[col].join(', ') + '.';
    }
    properties[col] = { type: 'string', description: desc };
    required.push(col);
  });

  const schema = { type: 'object', properties: properties, required: required };

  const prompt = 'You are extracting structured fields from a job posting for a job application tracker.\n' +
    'If a field cannot be determined from the text, use "Not listed".\n\n' +
    'JOB POSTING TEXT:\n' + jobText;

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
    CONFIG.GEMINI_MODEL + ':generateContent?key=' + apiKey;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: schema
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const body = response.getContentText();
  if (code !== 200) {
    throw new Error('Gemini API error (' + code + '): ' + body.slice(0, 300));
  }

  const json = JSON.parse(body);
  const textOut = json.candidates[0].content.parts[0].text;
  return JSON.parse(textOut);
}

function appendRow(sheet, headerMap, dataObj) {
  const numCols = sheet.getLastColumn();
  const rowValues = new Array(numCols).fill('');
  Object.keys(headerMap).forEach(function (header) {
    if (dataObj.hasOwnProperty(header)) {
      rowValues[headerMap[header] - 1] = dataObj[header];
    }
  });
  const nextRow = getNextEmptyRow(sheet, headerMap);
  sheet.getRange(nextRow, 1, 1, numCols).setValues([rowValues]);
}

// Finds the first row (after the header) where your reference column (the
// first AI_TEXT_COLUMNS entry, e.g. Company) is empty. This avoids the common
// getLastRow() trap, where stray formatting/content far down the sheet in an
// unrelated column inflates what Sheets thinks the "last row" is.
function getNextEmptyRow(sheet, headerMap) {
  const refHeader = CONFIG.AI_TEXT_COLUMNS[0];
  const refCol = headerMap[refHeader] || 1;
  const numRows = Math.max(sheet.getMaxRows() - 1, 1);
  const values = sheet.getRange(2, refCol, numRows, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (!values[i][0]) return i + 2; // +2: row 1 is header, values[0] is row 2
  }
  return sheet.getMaxRows() + 1; // fallback: sheet is completely full
}
