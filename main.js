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






function createGenderSheet(ss, sheetName, gender) {
  // Create a new sheet
  let sheet = ss.getSheetByName(sheetName);
  // If the sheet already exists, delete and recreate it.
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet(sheetName);


  // Freeze rows
  sheet.setFrozenRows(2);


  // Set headers
  sheet.getRange("A1").setValue(sheetName.toUpperCase());
  sheet.getRange("A2").setValue("Item");
  sheet.getRange("B2").setValue("Spice Level");
  sheet.getRange("C2").setValue("Drink");
  sheet.getRange("D2").setValue("Name");
  sheet.getRange("E2").setValue("Gender");
  sheet.getRange("F2").setValue("Dietary");
  sheet.getRange("G2").setValue("Collected?");


  // Total order info
  sheet.getRange("C1").setValue(`TOTAL ${sheetName.toUpperCase()} ⸻>`);
  sheet.getRange("E1").setFormula(`=SUMPRODUCT((A3:A350<>"")*(A3:A350<>"refunded"))`);


  // Header formatting
  sheet.getRange("A1:G1")
    .setBackground("#12c9c6")
    .setFontWeight("bold")
    .setFontSize(16);


  sheet.getRange("A2:G2")
    .setBackground("#e0e0e0")
    .setFontWeight("bold");


  // FILTER function based on gender, with SPLIT for Line Item Variant
  sheet.getRange("A3").setFormula(
    `=SORT(` +
      `HSTACK(` +
        `ARRAYFORMULA(SPLIT(` +
          `FILTER(Sheet1!U2:U, ISNUMBER(SEARCH("${gender}", Sheet1!AU2:AU)), Sheet1!C2:C<>"refunded"),` +
          `"/"` +
        `)),` +
        `FILTER(Sheet1!AT2:AV, ISNUMBER(SEARCH("${gender}", Sheet1!AU2:AU)), Sheet1!C2:C<>"refunded")` +
      `),` +
    `4, TRUE)`
  );




  // Add checkboxes in column G starting from G3
  sheet.getRange("G3:G350").insertCheckboxes();


  // Add conditional formatting for checked rows.
  const lastRow = 350; // Upper limit of orders
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$G3=TRUE')
    .setBackground("#90EE90") // Light green
    .setRanges([sheet.getRange(`A3:G${lastRow}`)])
    .build();


  const rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);


  // Set column widths (in pixels)
  for (let i = 1; i <= 7; i++) {
    sheet.setColumnWidth(i, 200);
  }
}


function createCombinedSummarySheet(ss) {
  const dataSheet = ss.getSheetByName("Sheet1");
  if (!dataSheet) return;


  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) return;


  const lineItems = dataSheet.getRange("U2:U" + lastRow).getValues(); // Item/Spice/Drink
  const statuses  = dataSheet.getRange("C2:C" + lastRow).getValues(); // refunded?


  const itemSpiceCounts = {}; // merges MEAL + non-MEAL (after stripping MEAL:)
  const drinkCounts = {};


  const mealByItem = {};      // MEAL-only by main item (stripped)
  const nonMealByItem = {};   // Non-MEAL-only by main item


  let totalMeals = 0;
  let totalNonMeals = 0;


  for (let i = 0; i < lineItems.length; i++) {
    const status = (statuses[i][0] || "").toString().toLowerCase();
    if (status === "refunded") continue;


    const raw = lineItems[i][0];
    if (!raw) continue;


    const parts = String(raw).split("/");
    if (parts.length < 3) continue;


    let item = parts[0].trim();   // may be "MEAL: X"
    const spice = parts[1].trim();
    const drink = parts[2].trim();


    const isMeal = /^meal:/i.test(item);
    if (isMeal) {
      totalMeals++;
      const stripped = item.replace(/^meal:\s*/i, "").trim();
      mealByItem[stripped] = (mealByItem[stripped] || 0) + 1;
      item = stripped; // merge into food counts
    } else {
      totalNonMeals++;
      nonMealByItem[item] = (nonMealByItem[item] || 0) + 1;
    }


    // Food aggregation (MEAL + non-MEAL merged) by item+spice
    const itemKey = item + "||" + spice;
    itemSpiceCounts[itemKey] = (itemSpiceCounts[itemKey] || 0) + 1;


    // Drinks aggregation (skip "None ...")
    if (
      drink &&
      drink.toLowerCase() !== "none (can only buy drink with meal)".toLowerCase()
    ) {
      drinkCounts[drink] = (drinkCounts[drink] || 0) + 1;
    }
  }


  // Create/reset summary sheet
  let sheet = ss.getSheetByName("Orders Summary");
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet("Orders Summary");


  // Headers (display-only columns)
  sheet.getRange("A1").setValue("Food Orders");
  sheet.getRange("C1").setValue("Drink Orders");
  sheet.getRange("E1").setValue("MEAL Breakdown (by main item)");
  sheet.getRange("F1").setValue("Non-MEAL Breakdown (by main item)");


  // Build display lists
  const foodDisplay = Object.keys(itemSpiceCounts)
    .sort()
    .map(key => {
      const [itemName, spice] = key.split("||");
      const count = itemSpiceCounts[key];
      return [`${itemName}: ${spice} x${count}`];
    });


  const drinkDisplay = Object.keys(drinkCounts)
    .sort()
    .map(drink => [`${drink} x${drinkCounts[drink]}`]);


  const mealDisplay = Object.keys(mealByItem)
    .sort()
    .map(itemName => [`${itemName} MEAL x${mealByItem[itemName]}`]);


  const nonMealDisplay = Object.keys(nonMealByItem)
    .sort()
    .map(itemName => [`${itemName} x${nonMealByItem[itemName]}`]);


  // Write columns
  if (foodDisplay.length)    sheet.getRange(2, 1, foodDisplay.length, 1).setValues(foodDisplay);
  if (drinkDisplay.length)   sheet.getRange(2, 3, drinkDisplay.length, 1).setValues(drinkDisplay);
  if (mealDisplay.length)    sheet.getRange(2, 5, mealDisplay.length, 1).setValues(mealDisplay);
  if (nonMealDisplay.length) sheet.getRange(2, 6, nonMealDisplay.length, 1).setValues(nonMealDisplay);


  // Footer block (labels + values)
  const footerRow = Math.max(foodDisplay.length, drinkDisplay.length, mealDisplay.length, nonMealDisplay.length) + 4;


  sheet.getRange(footerRow, 1).setValue("Total MEAL Orders:");
  sheet.getRange(footerRow, 2).setValue(totalMeals);


  sheet.getRange(footerRow + 1, 1).setValue("Total Non-MEAL Orders:");
  sheet.getRange(footerRow + 1, 2).setValue(totalNonMeals);


  sheet.getRange(footerRow + 2, 1).setValue("Extra Orders:");
  sheet.getRange(footerRow + 2, 2).setValue(""); // manual later


  sheet.getRange(footerRow + 3, 1).setValue("Total Orders:");
  // ✅ formula sum of the 3 cells above it (meal + nonmeal + extra)
  sheet.getRange(footerRow + 3, 2).setFormula(`=SUM(B${footerRow}:B${footerRow + 2})`);


  // Formatting + readable widths
  sheet.getRange("A1:F1").setFontWeight("bold").setBackground("#e0e0e0");
  [1, 3, 5, 6].forEach(c => sheet.autoResizeColumn(c));
  sheet.setColumnWidth(1, Math.max(sheet.getColumnWidth(1), 320));
  sheet.setColumnWidth(3, Math.max(sheet.getColumnWidth(3), 220));
  sheet.setColumnWidth(5, Math.max(sheet.getColumnWidth(5), 320));
  sheet.setColumnWidth(6, Math.max(sheet.getColumnWidth(6), 320));
  sheet.getRange(footerRow, 1, 4, 1).setFontWeight("bold");
}






function createGenderSummarySheet(ss, gender, sheetName) {
  const dataSheet = ss.getSheetByName("Sheet1");
  if (!dataSheet) return;


  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) return;


  const lineItems = dataSheet.getRange("U2:U" + lastRow).getValues();   // Item/Spice/Drink
  const statuses  = dataSheet.getRange("C2:C" + lastRow).getValues();   // refunded?
  const genders   = dataSheet.getRange("AU2:AU" + lastRow).getValues(); // gender text


  const itemSpiceCounts = {};
  const drinkCounts = {};
  let totalMeals = 0;


  for (let i = 0; i < lineItems.length; i++) {
    const status = (statuses[i][0] || "").toString().toLowerCase();
    if (status === "refunded") continue;


    const g = (genders[i][0] || "").toString();
    if (!g.includes(gender)) continue;


    const raw = lineItems[i][0];
    if (!raw) continue;


    const parts = String(raw).split("/");
    if (parts.length < 3) continue;


    let item = parts[0].trim();
    const spice = parts[1].trim();
    const drink = parts[2].trim();


    const isMeal = /^meal:/i.test(item);
    if (isMeal) {
      totalMeals++;
      item = item.replace(/^meal:\s*/i, "").trim(); // merge MEAL + non-MEAL
    }


    // Food aggregation by item+spice
    const itemKey = item + "||" + spice;
    itemSpiceCounts[itemKey] = (itemSpiceCounts[itemKey] || 0) + 1;


    // Drinks aggregation (skip "None ...")
    if (
      drink &&
      drink.toLowerCase() !== "none (can only buy drink with meal)".toLowerCase()
    ) {
      drinkCounts[drink] = (drinkCounts[drink] || 0) + 1;
    }
  }


  // Create/reset sheet
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(sheetName);


  // Headers
  sheet.getRange("A1").setValue("Food Orders");
  sheet.getRange("C1").setValue("Drink Orders");


  // Display lists
  const foodDisplay = Object.keys(itemSpiceCounts)
    .sort()
    .map(key => {
      const [itemName, spice] = key.split("||");
      const count = itemSpiceCounts[key];
      return [`${itemName}: ${spice} x${count}`];
    });


  const drinkDisplay = Object.keys(drinkCounts)
    .sort()
    .map(d => [`${d} x${drinkCounts[d]}`]);


  if (foodDisplay.length) sheet.getRange(2, 1, foodDisplay.length, 1).setValues(foodDisplay);
  if (drinkDisplay.length) sheet.getRange(2, 3, drinkDisplay.length, 1).setValues(drinkDisplay);


  // ✅ Footer as label/value in A/B (not merged into one cell)
  const footerRow = Math.max(foodDisplay.length, drinkDisplay.length) + 4;
  sheet.getRange(footerRow, 1).setValue("Total MEAL Orders:");
  sheet.getRange(footerRow, 2).setValue(totalMeals);


  // Formatting + readable widths
  sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#e0e0e0");
  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(3);
  sheet.setColumnWidth(1, Math.max(sheet.getColumnWidth(1), 320));
  sheet.setColumnWidth(3, Math.max(sheet.getColumnWidth(3), 220));


  // Footer label bold
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
  if (!sum) {
    return { rich: SpreadsheetApp.newRichTextValue().setText("Orders Summary sheet not found.").build() };
  }


  // ✅ Food list stops before footer labels
  const foodList = readColumnUntilLabel(sum, 1, 2, "Total MEAL Orders:"); // col A
  const drinkList = readNonEmptyColumn(sum, 3, 2);                        // col C
  const mealBreakdown = readNonEmptyColumn(sum, 5, 2);                    // col E


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
  lines.push("Drinks Orders");
  lines.push(...drinkList);


  const text = lines.join("\n");


  const builder = SpreadsheetApp.newRichTextValue().setText(text);


  // Bold: title + section headings + "Meals Breakdown By Item"
  applyBoldToLine(builder, text, 1);
  applyBoldToSubstring(builder, text, "Food Orders");
  applyBoldToSubstring(builder, text, "Meals Breakdown By Item");
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
  if (!sum) {
    return { rich: SpreadsheetApp.newRichTextValue().setText(sheetName + " not found.").build() };
  }


  // Food list: stop BEFORE the footer label appears in column A
  const foodList = readColumnUntilLabel(sum, 1, 2, "Total MEAL Orders:");
  const drinkList = readNonEmptyColumn(sum, 3, 2);


  const totalMeals = findFooterValue(sum, "Total MEAL Orders:");
  const totalMealsText = (totalMeals === null || totalMeals === undefined) ? "" : String(totalMeals);


  const lines = [];
  lines.push(title);
  lines.push("");
  lines.push("Food Orders");
  lines.push(...foodList);
  lines.push("");
  lines.push(`Total MEAL Orders: ${totalMealsText}`);
  lines.push("");
  lines.push("Drinks Orders");
  lines.push(...drinkList);


  const text = lines.join("\n");


  const builder = SpreadsheetApp.newRichTextValue().setText(text);


  // Bold title + headings
  applyBoldToLine(builder, text, 1);
  applyBoldToSubstring(builder, text, "Food Orders");
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

