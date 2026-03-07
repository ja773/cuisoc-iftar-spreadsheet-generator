function createIftarSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Orders sheets
  createGenderSheet(ss, "Brothers' Orders", "Brother");
  createGenderSheet(ss, "Sisters' Orders", "Sister");

  // Summaries
  createCombinedSummarySheet(ss); // Orders Summary (with meal + non-meal breakdown, totals)
  createGenderSummarySheet(ss, "Brother", "Brothers' Summary"); // no meal breakdown
  createGenderSummarySheet(ss, "Sister", "Sisters' Summary");   // no meal breakdown

  // NEW messages sheet
  createMessagesSheet(ss);

  // Sanitise Sheet1 of sensitive data
  sanitizeSheet1_(ss);
}



function createGenderSheet(ss, sheetName, genderNeedle) {
  const dataSheet = ss.getSheetByName("Sheet1");
  if (!dataSheet) throw new Error("Sheet1 not found");

  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(sheetName);

  sheet.setFrozenRows(2);

  sheet.getRange("A1").setValue(sheetName.toUpperCase());

  // Main | Spice | Side | Drink | Name | Gender | Dietary | Collected?
  sheet.getRange("A2").setValue("Main Item");
  sheet.getRange("B2").setValue("Spice Level");
  sheet.getRange("C2").setValue("Side");
  sheet.getRange("D2").setValue("Drink");
  sheet.getRange("E2").setValue("Name");
  sheet.getRange("F2").setValue("Gender");
  sheet.getRange("G2").setValue("Dietary");
  sheet.getRange("H2").setValue("Collected?");

  sheet.getRange("A1:H1")
    .setBackground("#12c9c6")
    .setFontWeight("bold")
    .setFontSize(16);

  sheet.getRange("A2:H2")
    .setBackground("#e0e0e0")
    .setFontWeight("bold");

  const lastRow = dataSheet.getLastRow();
  const lastCol = dataSheet.getLastColumn();
  if (lastRow < 2) return;

  const headers = dataSheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(h => String(h || "").trim());

  const colStatus = colIndexByHeader_(headers, "Status", 3);
  const colVariant = colIndexByHeader_(headers, "Line Item Variant", colA1ToIndex_("U"));
  const colGender = colIndexByHeader_(headers, "Gender", colA1ToIndex_("AU"));
  const colDiet = colIndexByHeader_(headers, "Dietary", colA1ToIndex_("AV"));
  const colProdFormName = colIndexByHeader_(headers, "Product Form: Name", null);

  const values = dataSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rowsOut = [];

  for (const row of values) {
    const status = String(row[colStatus - 1] || "").toLowerCase();
    if (status === "refunded") continue;

    const g = String(row[colGender - 1] || "").trim();
    if (!g) continue;
    if (g.indexOf(genderNeedle) === -1) continue;

    const rawVariant = String(row[colVariant - 1] || "").trim();
    if (!rawVariant) continue;

    const parts = rawVariant.split("/").map(p => p.trim());

    // ✅ STRICT: must be exactly 4 parts
    if (parts.length !== 4) continue;

    let [main, spice, side, drink] = parts;

    // Blank out NA values
    side = isNA_(side) ? "" : side;
    spice = isNA_(spice) ? "" : spice;
    drink = isNA_(drink) ? "" : drink;

    const name = colProdFormName
      ? String(row[colProdFormName - 1] || "").trim()
      : "";

    const dietary = String(row[colDiet - 1] || "").trim();

    rowsOut.push([main, spice, side, drink, name, g, dietary]);
  }

  // Sort by Name (column index 4 in rowsOut)
  rowsOut.sort((a, b) =>
    (a[4] || "").localeCompare(b[4] || "", "en", { sensitivity: "base" })
  );

  if (rowsOut.length) {
    sheet.getRange(3, 1, rowsOut.length, 7).setValues(rowsOut);
  }

  sheet.getRange("E1").setValue(rowsOut.length);

  // Checkboxes
  const maxRows = Math.max(350, rowsOut.length + 10);
  sheet.getRange(`H3:H${maxRows}`).insertCheckboxes();

  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$H3=TRUE')
    .setBackground("#90EE90")
    .setRanges([sheet.getRange(`A3:H${maxRows}`)])
    .build();

  const rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);

  const widths = [240, 160, 240, 180, 220, 140, 220, 120];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

function isNA_(s) {
  const t = String(s || "").trim().toLowerCase();
  return t === "" || t === "na" || t.startsWith("na -") || t.startsWith("na ");
}

function colIndexByHeader_(headers, headerName, fallbackIndex) {
  const idx = headers.findIndex(
    h => h.toLowerCase() === String(headerName).toLowerCase()
  );
  if (idx >= 0) return idx + 1;
  return fallbackIndex;
}

function colA1ToIndex_(letters) {
  let n = 0;
  const s = String(letters).toUpperCase().replace(/[^A-Z]/g, "");
  for (let i = 0; i < s.length; i++)
    n = n * 26 + (s.charCodeAt(i) - 64);
  return n;
}

/** Parse line item variant supporting 4-part main/side/spice/drink with 3-part fallback main/drink/spice. */
function parseVariant4_(raw) {
  const parts = String(raw || "").split("/").map(p => p.trim());
  if (parts.length >= 4) return { main: parts[0], side: parts[1], spice: parts[2], drink: parts[3] };
  if (parts.length === 3) return { main: parts[0], side: "", spice: parts[2], drink: parts[1] };
  return { main: "", side: "", spice: "", drink: "" };
}

function isNA_(s) {
  const t = String(s || "").trim().toLowerCase();
  return t === "" || t === "na" || t.startsWith("na -") || t.startsWith("na ");
}

/** Find column index by header text; if not found, return fallbackIndex (1-based) or null. */
function colIndexByHeader_(headers, headerName, fallbackIndex) {
  const idx = headers.findIndex(h => h.toLowerCase() === String(headerName).toLowerCase());
  if (idx >= 0) return idx + 1;
  return fallbackIndex;
}

function colA1ToIndex_(letters) {
  let n = 0;
  const s = String(letters).toUpperCase().replace(/[^A-Z]/g, "");
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n;
}

function createCombinedSummarySheet(ss) {
  const dataSheet = ss.getSheetByName("Sheet1");
  if (!dataSheet) return;

  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) return;

  const lineItems = dataSheet.getRange("U2:U" + lastRow).getValues();
  const statuses  = dataSheet.getRange("C2:C" + lastRow).getValues();
  const genders   = dataSheet.getRange("AU2:AU" + lastRow).getValues();

  const mainSpiceCounts = {};
  const sideCounts = {};
  const drinkCounts = {};
  const mealByMain = {};
  const nonMealByMain = {};

  let totalMeals = 0;
  let totalNonMeals = 0;

  for (let i = 0; i < lineItems.length; i++) {
    const status = (statuses[i][0] || "").toString().toLowerCase();
    if (status === "refunded") continue;

    // ✅ skip missing gender entirely
    const g = String(genders[i][0] || "").trim();
    if (!g) continue;

    const raw = lineItems[i][0];
    if (!raw) continue;

    const v = parseVariant_(raw);
    if (!v.main) continue;

    let main = v.main.trim();
    const spice = isNA_(v.spice) ? "" : v.spice.trim();
    const side = isNA_(v.side) ? "" : v.side.trim();
    const drink = isNA_(v.drink) ? "" : v.drink.trim();

    const isMeal = /^meal:/i.test(main);
    if (isMeal) {
      totalMeals++;
      const stripped = main.replace(/^meal:\s*/i, "").trim();
      mealByMain[stripped] = (mealByMain[stripped] || 0) + 1;
      main = stripped;
    } else {
      totalNonMeals++;
      nonMealByMain[main] = (nonMealByMain[main] || 0) + 1;
    }

    // Main food summary (by main+spice)
    const mainKey = main + "||" + spice;
    mainSpiceCounts[mainKey] = (mainSpiceCounts[mainKey] || 0) + 1;

    // Side summary (just by side)
    if (side) sideCounts[side] = (sideCounts[side] || 0) + 1;

    // Drinks summary
    if (drink) drinkCounts[drink] = (drinkCounts[drink] || 0) + 1;
  }

  let sheet = ss.getSheetByName("Orders Summary");
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet("Orders Summary");

  sheet.getRange("A1").setValue("Food Orders");
  sheet.getRange("C1").setValue("Side Orders");
  sheet.getRange("E1").setValue("Drink Orders");
  sheet.getRange("G1").setValue("MEAL Breakdown (by main item)");
  sheet.getRange("H1").setValue("Non-MEAL Breakdown (by main item)");

  const foodDisplay = Object.keys(mainSpiceCounts).sort().map(k => {
    const [main, spice] = k.split("||");
    return [foodDisplayLine_(main, spice, mainSpiceCounts[k])];
  });

  const sideDisplay = Object.keys(sideCounts).sort().map(s => [`${s} x${sideCounts[s]}`]);
  const drinkDisplay = Object.keys(drinkCounts).sort().map(d => [`${d} x${drinkCounts[d]}`]);

  const mealDisplay = Object.keys(mealByMain).sort().map(m => [`${m} MEAL x${mealByMain[m]}`]);
  const nonMealDisplay = Object.keys(nonMealByMain).sort().map(m => [`${m} x${nonMealByMain[m]}`]);

  if (foodDisplay.length)    sheet.getRange(2, 1, foodDisplay.length, 1).setValues(foodDisplay);
  if (sideDisplay.length)    sheet.getRange(2, 3, sideDisplay.length, 1).setValues(sideDisplay);
  if (drinkDisplay.length)   sheet.getRange(2, 5, drinkDisplay.length, 1).setValues(drinkDisplay);
  if (mealDisplay.length)    sheet.getRange(2, 7, mealDisplay.length, 1).setValues(mealDisplay);
  if (nonMealDisplay.length) sheet.getRange(2, 8, nonMealDisplay.length, 1).setValues(nonMealDisplay);

  const footerRow = Math.max(foodDisplay.length, sideDisplay.length, drinkDisplay.length, mealDisplay.length, nonMealDisplay.length) + 4;

  sheet.getRange(footerRow, 1).setValue("Total MEAL Orders:");
  sheet.getRange(footerRow, 2).setValue(totalMeals);

  sheet.getRange(footerRow + 1, 1).setValue("Total Non-MEAL Orders:");
  sheet.getRange(footerRow + 1, 2).setValue(totalNonMeals);

  sheet.getRange(footerRow + 2, 1).setValue("Extra Orders:");
  sheet.getRange(footerRow + 2, 2).setValue("");

  sheet.getRange(footerRow + 3, 1).setValue("Total Orders:");
  sheet.getRange(footerRow + 3, 2).setFormula(`=SUM(B${footerRow}:B${footerRow + 2})`);

  sheet.getRange("A1:H1").setFontWeight("bold").setBackground("#e0e0e0");
  [1,3,5,7,8].forEach(c => sheet.autoResizeColumn(c));
  sheet.setColumnWidth(1, Math.max(sheet.getColumnWidth(1), 320));
  sheet.setColumnWidth(3, Math.max(sheet.getColumnWidth(3), 260));
  sheet.setColumnWidth(5, Math.max(sheet.getColumnWidth(5), 220));
  sheet.setColumnWidth(7, Math.max(sheet.getColumnWidth(7), 320));
  sheet.setColumnWidth(8, Math.max(sheet.getColumnWidth(8), 320));
}



function createGenderSummarySheet(ss, gender, sheetName) {
  const dataSheet = ss.getSheetByName("Sheet1");
  if (!dataSheet) return;

  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) return;

  const lineItems = dataSheet.getRange("U2:U" + lastRow).getValues();
  const statuses  = dataSheet.getRange("C2:C" + lastRow).getValues();
  const genders   = dataSheet.getRange("AU2:AU" + lastRow).getValues();

  const mainSpiceCounts = {};
  const sideCounts = {};
  const drinkCounts = {};
  let totalMeals = 0;

  for (let i = 0; i < lineItems.length; i++) {
    const status = (statuses[i][0] || "").toString().toLowerCase();
    if (status === "refunded") continue;

    const g = String(genders[i][0] || "").trim();
    if (!g) continue;                 // ✅ skip missing gender
    if (!g.includes(gender)) continue;

    const raw = lineItems[i][0];
    if (!raw) continue;

    const v = parseVariant_(raw);
    if (!v.main) continue;

    let main = v.main.trim();
    const spice = isNA_(v.spice) ? "" : v.spice.trim();
    const side = isNA_(v.side) ? "" : v.side.trim();
    const drink = isNA_(v.drink) ? "" : v.drink.trim();

    if (/^meal:/i.test(main)) {
      totalMeals++;
      main = main.replace(/^meal:\s*/i, "").trim();
    }

    const mainKey = main + "||" + spice;
    mainSpiceCounts[mainKey] = (mainSpiceCounts[mainKey] || 0) + 1;

    if (side) sideCounts[side] = (sideCounts[side] || 0) + 1;
    if (drink) drinkCounts[drink] = (drinkCounts[drink] || 0) + 1;
  }

  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(sheetName);

  sheet.getRange("A1").setValue("Food Orders");
  sheet.getRange("C1").setValue("Side Orders");
  sheet.getRange("E1").setValue("Drink Orders");

  const foodDisplay = Object.keys(mainSpiceCounts).sort().map(k => {
    const [main, spice] = k.split("||");
    return [foodDisplayLine_(main, spice, mainSpiceCounts[k])];
  });
  const sideDisplay = Object.keys(sideCounts).sort().map(s => [`${s} x${sideCounts[s]}`]);
  const drinkDisplay = Object.keys(drinkCounts).sort().map(d => [`${d} x${drinkCounts[d]}`]);

  if (foodDisplay.length)  sheet.getRange(2, 1, foodDisplay.length, 1).setValues(foodDisplay);
  if (sideDisplay.length)  sheet.getRange(2, 3, sideDisplay.length, 1).setValues(sideDisplay);
  if (drinkDisplay.length) sheet.getRange(2, 5, drinkDisplay.length, 1).setValues(drinkDisplay);

  const footerRow = Math.max(foodDisplay.length, sideDisplay.length, drinkDisplay.length) + 4;
  sheet.getRange(footerRow, 1).setValue("Total MEAL Orders:");
  sheet.getRange(footerRow, 2).setValue(totalMeals);

  sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#e0e0e0");
  [1,3,5].forEach(c => sheet.autoResizeColumn(c));
  sheet.setColumnWidth(1, Math.max(sheet.getColumnWidth(1), 320));
  sheet.setColumnWidth(3, Math.max(sheet.getColumnWidth(3), 260));
  sheet.setColumnWidth(5, Math.max(sheet.getColumnWidth(5), 220));
  sheet.getRange(footerRow, 1).setFontWeight("bold");
}


function createMessagesSheet(ss) {
  // Create/reset messages sheet
  let sheet = ss.getSheetByName("Messages");
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet("Messages");

  // Make two big cells: A1 and B1
  sheet.setColumnWidth(1, 650); // A
  sheet.setColumnWidth(2, 650); // B
  sheet.setRowHeight(1, 900);   // very tall

  // Build left message from Orders Summary (includes meal breakdown list)
  const left = buildOrdersMessage_Combined(ss);

  // Build right message from Sisters' Summary (no meal breakdown list; includes totals + food/drink lists)
  const right = buildOrdersMessage_GenderSummary(ss, "Sisters' Summary", "SISTERS ORDERS");

  sheet.getRange("A1").setRichTextValue(left.rich);
  sheet.getRange("B1").setRichTextValue(right.rich);

  // Wrap + top align for readability
  sheet.getRange("A1:B1").setWrap(true).setVerticalAlignment("top");
}

/**
 * Build message for combined Orders Summary:
 * - ORDERS SUMMARY (bold)
 * - Food Orders (bold) + list
 * - Total MEAL Orders line + meal breakdown list
 * - Drinks Orders (bold) + list
 */
function buildOrdersMessage_Combined(ss) {
  const sum = ss.getSheetByName("Orders Summary");
  if (!sum) return { rich: SpreadsheetApp.newRichTextValue().setText("Orders Summary sheet not found.").build() };

  const foodList = readColumnUntilLabel(sum, 1, 2, "Total MEAL Orders:");
  const sideList = readNonEmptyColumn(sum, 3, 2);   // col C
  const drinkList = readNonEmptyColumn(sum, 5, 2);  // col E
  const mealBreakdown = readNonEmptyColumn(sum, 7, 2); // col G

  const totalMeals = findFooterValue(sum, "Total MEAL Orders:");

  const lines = [];
  lines.push("ORDERS SUMMARY");
  lines.push("");
  lines.push("Food Orders");
  lines.push(...foodList);
  lines.push("");
  lines.push(`Total MEAL Orders: ${totalMeals !== null ? totalMeals : ""}`);
  lines.push("Meals Breakdown By Item");
  lines.push(...mealBreakdown);
  lines.push("");
  lines.push("Side Orders");
  lines.push(...sideList);
  lines.push("");
  lines.push("Drinks Orders");
  lines.push(...drinkList);

  const text = lines.join("\n");
  const builder = SpreadsheetApp.newRichTextValue().setText(text);

  applyBoldToLine(builder, text, 1);
  applyBoldToSubstring(builder, text, "Food Orders");
  applyBoldToSubstring(builder, text, "Meals Breakdown By Item");
  applyBoldToSubstring(builder, text, "Side Orders");
  applyBoldToSubstring(builder, text, "Drinks Orders");

  return { rich: builder.build() };
}


/**
 * Build message for a gender summary sheet (Brothers' Summary / Sisters' Summary):
 * - Title line (bold)
 * - Food Orders (bold) + list
 * - Total MEAL Orders line (normal)
 * - Drinks Orders (bold) + list
 */
function buildOrdersMessage_GenderSummary(ss, sheetName, title) {
  const sum = ss.getSheetByName(sheetName);
  if (!sum) return { rich: SpreadsheetApp.newRichTextValue().setText(sheetName + " not found.").build() };

  const foodList = readColumnUntilLabel(sum, 1, 2, "Total MEAL Orders:");
  const sideList = readNonEmptyColumn(sum, 3, 2);   // col C
  const drinkList = readNonEmptyColumn(sum, 5, 2);  // col E
  const totalMeals = findFooterValue(sum, "Total MEAL Orders:");

  const lines = [];
  lines.push(title);
  lines.push("");
  lines.push("Food Orders");
  lines.push(...foodList);
  lines.push("");
  lines.push(`Total MEAL Orders: ${totalMeals !== null ? totalMeals : ""}`);
  lines.push("");
  lines.push("Side Orders");
  lines.push(...sideList);
  lines.push("");
  lines.push("Drinks Orders");
  lines.push(...drinkList);

  const text = lines.join("\n");
  const builder = SpreadsheetApp.newRichTextValue().setText(text);

  applyBoldToLine(builder, text, 1);
  applyBoldToSubstring(builder, text, "Food Orders");
  applyBoldToSubstring(builder, text, "Side Orders");
  applyBoldToSubstring(builder, text, "Drinks Orders");

  return { rich: builder.build() };
}



/** Read a single column as a list of non-empty strings starting from startRow. */
function readNonEmptyColumn(sheet, col, startRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return [];
  const values = sheet.getRange(startRow, col, lastRow - startRow + 1, 1).getValues().flat();
  return values
    .map(v => (v === null || v === undefined) ? "" : String(v).trim())
    .filter(v => v !== "");
}

/**
 * Find a footer value by scanning column A for a label and returning the column B value.
 * Returns null if not found.
 */
function findFooterValue(sheet, label) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return null;

  const colA = sheet.getRange(1, 1, lastRow, 1).getValues().flat();
  for (let r = 0; r < colA.length; r++) {
    if (String(colA[r]).trim() === label) {
      const val = sheet.getRange(r + 1, 2).getValue();
      return val === "" ? "" : val;
    }
  }
  return null;
}

/** Bold an entire line number (1-indexed) in the given text. */
function applyBoldToLine(builder, text, lineNumber) {
  const lines = text.split("\n");
  let start = 0;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const end = start + line.length;
    if (i + 1 === lineNumber) {
      builder.setTextStyle(
        start,
        end,
        SpreadsheetApp.newTextStyle().setBold(true).build()
      );
      return;
    }
    start = end + 1; // +1 for '\n'
  }
}

/** Bold the first occurrence of a substring in the text. */
function applyBoldToSubstring(builder, text, substr) {
  const idx = text.indexOf(substr);
  if (idx === -1) return;
  builder.setTextStyle(
    idx,
    idx + substr.length,
    SpreadsheetApp.newTextStyle().setBold(true).build()
  );
}

/** Read a single column as a list of non-empty strings starting from startRow, stopping when a footer label is hit. */
function readColumnUntilLabel(sheet, col, startRow, stopLabel) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return [];

  const values = sheet.getRange(startRow, col, lastRow - startRow + 1, 1).getValues().flat();
  const out = [];

  for (const v of values) {
    const s = (v === null || v === undefined) ? "" : String(v).trim();
    if (!s) continue;
    if (s === stopLabel) break;          // stop before footer section
    out.push(s);
  }
  return out;
}

function sanitizeSheet1_(ss) {
  const sheet = ss.getSheetByName("Sheet1");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  // Read headers (assumes headers are in row 1)
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());

  // We MUST NOT clear columns you use for generating everything:
  // C  = status (refunded)
  // U  = line item variant (Item/Spice/Drink)
  // AT:AV = name/gender/dietary (from your HSTACK filter)
  // AU = gender filter string
  const protectedLetters = new Set(["C", "U", "AT", "AU", "AV"]);
  const protectedIndexes = new Set([...protectedLetters].map(colA1ToIndex_)); // 1-based indexes

  // Keywords that typically indicate sensitive personal data (edit freely)
  const sensitiveHeaderRegex = /(billing|shipping|address|addr|postcode|post code|zip|eircode|city|county|state|country|phone|mobile|email|e-mail|customer|contact|first name|last name|surname|company|order notes|note)/i;

  // Collect columns to clear
  const colsToClear = [];
  for (let i = 0; i < headers.length; i++) {
    const colIndex = i + 1; // 1-based
    if (protectedIndexes.has(colIndex)) continue;

    const header = headers[i];
    if (!header) continue;

    if (sensitiveHeaderRegex.test(header)) {
      colsToClear.push(colIndex);
    }
  }

  // Clear values in those columns for all data rows (row 2 -> lastRow)
  // (We clear contents but keep formatting/column structure intact.)
  for (const colIndex of colsToClear) {
    sheet.getRange(2, colIndex, lastRow - 1, 1).clearContent();
  }

  // Optional: also clear notes/comments in cleared columns
  // (Uncomment if you want notes wiped too)
  // for (const colIndex of colsToClear) {
  //   sheet.getRange(2, colIndex, lastRow - 1, 1).clearNote();
  // }
}

/** Convert A1 column letters like "C" or "AU" to 1-based index */
function colA1ToIndex_(letters) {
  let n = 0;
  const s = String(letters).toUpperCase().replace(/[^A-Z]/g, "");
  for (let i = 0; i < s.length; i++) {
    n = n * 26 + (s.charCodeAt(i) - 64);
  }
  return n;
}

function isNA_(s) {
  const t = String(s || "").trim().toLowerCase();
  return t === "" || t === "na" || t.startsWith("na -") || t.startsWith("didn");
}

function parseVariant_(raw) {
  const parts = String(raw || "").split("/").map(p => p.trim());
  // 4-part: main/side/spice/drink
  if (parts.length >= 4) return { main: parts[0], spice: parts[1], side: parts[2], drink: parts[3] };
  // fallback 3-part: main/drink/spice
  if (parts.length === 3) return { main: parts[0], spice: "", side: parts[2], drink: parts[1] };
  return { main: "", side: "", spice: "", drink: "" };
}

function foodDisplayLine_(main, spice, count) {
  const sp = isNA_(spice) ? "" : String(spice).trim();
  if (!sp) return `${main} x${count}`;
  return `${main}: ${sp} x${count}`;
}