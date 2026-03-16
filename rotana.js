function createIftarSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Orders sheets
  createGenderSheet(ss, "Brothers' Orders", "Brother");
  createGenderSheet(ss, "Sisters' Orders", "Sister");

  // Summaries
  createCombinedSummarySheet(ss);
  createGenderSummarySheet(ss, "Brother", "Brothers' Summary");
  createGenderSummarySheet(ss, "Sister", "Sisters' Summary");

  // Messages sheet
  createMessagesSheet(ss);

  // Sanitise Sheet1 of sensitive data
  sanitizeSheet1_(ss);
}

function createGenderSheet(ss, sheetName, gender) {
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(sheetName);

  sheet.setFrozenRows(2);

  sheet.getRange("A1").setValue(sheetName.toUpperCase());
  sheet.getRange("A2").setValue("Item");
  sheet.getRange("B2").setValue("Drink");
  sheet.getRange("C2").setValue("Name");
  sheet.getRange("D2").setValue("Gender");
  sheet.getRange("E2").setValue("Dietary");
  sheet.getRange("F2").setValue("Collected?");

  sheet.getRange("C1").setValue(`TOTAL ${sheetName.toUpperCase()} ⸻>`);
  sheet.getRange("E1").setFormula(`=SUMPRODUCT((A3:A350<>"")*(A3:A350<>"refunded"))`);

  sheet.getRange("A1:F1")
    .setBackground("#12c9c6")
    .setFontWeight("bold")
    .setFontSize(16);

  sheet.getRange("A2:F2")
    .setBackground("#e0e0e0")
    .setFontWeight("bold");

  // Line Item Variant format: Item/Drink
  // Example: MEAL: Chicken Shish (rice, garlic sauce, drink)/Vive Cola
  // Remove "MEAL: " from item display
  sheet.getRange("A3").setFormula(
    `=LET(` +
      `raw, FILTER(Sheet1!U2:U, ISNUMBER(SEARCH("${gender}", Sheet1!AU2:AU)), Sheet1!C2:C<>"refunded"),` +
      `s, ARRAYFORMULA(SPLIT(raw, "/")),` +
      `itemRaw, INDEX(s,,1),` +
      `item, ARRAYFORMULA(REGEXREPLACE(TRIM(itemRaw), "^MEAL:\\s*", "")),` +
      `drink, INDEX(s,,2),` +
      `core, HSTACK(item, drink),` +
      `rest, FILTER(Sheet1!AT2:AV, ISNUMBER(SEARCH("${gender}", Sheet1!AU2:AU)), Sheet1!C2:C<>"refunded"),` +
      `SORT(HSTACK(core, rest), 3, TRUE)` +
    `)`
  );

  sheet.getRange("F3:F350").insertCheckboxes();

  const lastRow = 350;
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F3=TRUE')
    .setBackground("#90EE90")
    .setRanges([sheet.getRange(`A3:F${lastRow}`)])
    .build();

  const rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);

  for (let i = 1; i <= 6; i++) sheet.setColumnWidth(i, 220);
}

function createCombinedSummarySheet(ss) {
  const dataSheet = ss.getSheetByName("Sheet1");
  if (!dataSheet) return;

  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) return;

  const lineItems = dataSheet.getRange("U2:U" + lastRow).getValues(); // Item/Drink
  const statuses = dataSheet.getRange("C2:C" + lastRow).getValues();  // refunded?

  const itemCounts = {};
  const drinkCounts = {};
  const itemBreakdown = {};

  let totalOrders = 0;

  for (let i = 0; i < lineItems.length; i++) {
    const status = (statuses[i][0] || "").toString().toLowerCase();
    if (status === "refunded") continue;

    const raw = lineItems[i][0];
    if (!raw) continue;

    const parts = String(raw).split("/");
    if (parts.length < 2) continue;

    let item = parts[0].trim().replace(/^meal:\s*/i, "");
    const drinkRaw = parts.slice(1).join("/").trim();
    const drink = isNA_(drinkRaw) ? "" : drinkRaw;

    totalOrders++;
    itemCounts[item] = (itemCounts[item] || 0) + 1;
    itemBreakdown[item] = (itemBreakdown[item] || 0) + 1;
    if (drink) drinkCounts[drink] = (drinkCounts[drink] || 0) + 1;
  }

  let sheet = ss.getSheetByName("Orders Summary");
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet("Orders Summary");

  sheet.getRange("A1").setValue("Food Orders");
  sheet.getRange("C1").setValue("Drink Orders");
  sheet.getRange("E1").setValue("Breakdown By Item");

  const foodDisplay = Object.keys(itemCounts)
    .sort()
    .map(item => [`${item} x${itemCounts[item]}`]);

  const drinkDisplay = Object.keys(drinkCounts)
    .sort()
    .map(drink => [`${drink} x${drinkCounts[drink]}`]);

  const breakdownDisplay = Object.keys(itemBreakdown)
    .sort()
    .map(item => [`${item} x${itemBreakdown[item]}`]);

  if (foodDisplay.length) sheet.getRange(2, 1, foodDisplay.length, 1).setValues(foodDisplay);
  if (drinkDisplay.length) sheet.getRange(2, 3, drinkDisplay.length, 1).setValues(drinkDisplay);
  if (breakdownDisplay.length) sheet.getRange(2, 5, breakdownDisplay.length, 1).setValues(breakdownDisplay);

  const footerRow = Math.max(foodDisplay.length, drinkDisplay.length, breakdownDisplay.length) + 4;
  sheet.getRange(footerRow, 1).setValue("Total Orders:");
  sheet.getRange(footerRow, 2).setValue(totalOrders);

  sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#e0e0e0");
  [1, 3, 5].forEach(c => sheet.autoResizeColumn(c));
  sheet.setColumnWidth(1, Math.max(sheet.getColumnWidth(1), 320));
  sheet.setColumnWidth(3, Math.max(sheet.getColumnWidth(3), 220));
  sheet.setColumnWidth(5, Math.max(sheet.getColumnWidth(5), 320));
  sheet.getRange(footerRow, 1).setFontWeight("bold");
}

function createGenderSummarySheet(ss, gender, sheetName) {
  const dataSheet = ss.getSheetByName("Sheet1");
  if (!dataSheet) return;

  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) return;

  const lineItems = dataSheet.getRange("U2:U" + lastRow).getValues();   // Item/Drink
  const statuses = dataSheet.getRange("C2:C" + lastRow).getValues();    // refunded?
  const genders = dataSheet.getRange("AU2:AU" + lastRow).getValues();   // gender text

  const itemCounts = {};
  const drinkCounts = {};
  let totalOrders = 0;

  for (let i = 0; i < lineItems.length; i++) {
    const status = (statuses[i][0] || "").toString().toLowerCase();
    if (status === "refunded") continue;

    const g = (genders[i][0] || "").toString();
    if (!g.includes(gender)) continue;

    const raw = lineItems[i][0];
    if (!raw) continue;

    const parts = String(raw).split("/");
    if (parts.length < 2) continue;

    let item = parts[0].trim().replace(/^meal:\s*/i, "");
    const drinkRaw = parts.slice(1).join("/").trim();
    const drink = isNA_(drinkRaw) ? "" : drinkRaw;

    totalOrders++;
    itemCounts[item] = (itemCounts[item] || 0) + 1;
    if (drink) drinkCounts[drink] = (drinkCounts[drink] || 0) + 1;
  }

  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(sheetName);

  sheet.getRange("A1").setValue("Food Orders");
  sheet.getRange("C1").setValue("Drink Orders");

  const foodDisplay = Object.keys(itemCounts)
    .sort()
    .map(item => [`${item} x${itemCounts[item]}`]);

  const drinkDisplay = Object.keys(drinkCounts)
    .sort()
    .map(drink => [`${drink} x${drinkCounts[drink]}`]);

  if (foodDisplay.length) sheet.getRange(2, 1, foodDisplay.length, 1).setValues(foodDisplay);
  if (drinkDisplay.length) sheet.getRange(2, 3, drinkDisplay.length, 1).setValues(drinkDisplay);

  const footerRow = Math.max(foodDisplay.length, drinkDisplay.length) + 4;
  sheet.getRange(footerRow, 1).setValue("Total Orders:");
  sheet.getRange(footerRow, 2).setValue(totalOrders);

  sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#e0e0e0");
  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(3);
  sheet.setColumnWidth(1, Math.max(sheet.getColumnWidth(1), 320));
  sheet.setColumnWidth(3, Math.max(sheet.getColumnWidth(3), 220));
  sheet.getRange(footerRow, 1).setFontWeight("bold");
}

function createMessagesSheet(ss) {
  let sheet = ss.getSheetByName("Messages");
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet("Messages");

  sheet.setColumnWidth(1, 650);
  sheet.setColumnWidth(2, 650);
  sheet.setRowHeight(1, 900);

  const left = buildOrdersMessage_Combined(ss);
  const right = buildOrdersMessage_GenderSummary(ss, "Sisters' Summary", "SISTERS ORDERS");

  sheet.getRange("A1").setRichTextValue(left.rich);
  sheet.getRange("B1").setRichTextValue(right.rich);

  sheet.getRange("A1:B1").setWrap(true).setVerticalAlignment("top");
}

function buildOrdersMessage_Combined(ss) {
  const sum = ss.getSheetByName("Orders Summary");
  if (!sum) {
    return { rich: SpreadsheetApp.newRichTextValue().setText("Orders Summary sheet not found.").build() };
  }

  const foodList = readColumnUntilLabel(sum, 1, 2, "Total Orders:");
  const drinkList = readNonEmptyColumn(sum, 3, 2);
  const breakdown = readNonEmptyColumn(sum, 5, 2);
  const totalOrders = findFooterValue(sum, "Total Orders:");

  const lines = [];
  lines.push("ORDERS SUMMARY");
  lines.push("");
  lines.push("Food Orders");
  lines.push(...foodList);
  lines.push("");
  lines.push(`Total Orders: ${totalOrders !== null ? totalOrders : ""}`);
  lines.push("Breakdown By Item");
  lines.push(...breakdown);
  lines.push("");
  lines.push("Drinks Orders");
  lines.push(...drinkList);

  const text = lines.join("\n");
  const builder = SpreadsheetApp.newRichTextValue().setText(text);

  applyBoldToLine(builder, text, 1);
  applyBoldToSubstring(builder, text, "Food Orders");
  applyBoldToSubstring(builder, text, "Breakdown By Item");
  applyBoldToSubstring(builder, text, "Drinks Orders");

  return { rich: builder.build() };
}

function buildOrdersMessage_GenderSummary(ss, sheetName, title) {
  const sum = ss.getSheetByName(sheetName);
  if (!sum) {
    return { rich: SpreadsheetApp.newRichTextValue().setText(sheetName + " not found.").build() };
  }

  const foodList = readColumnUntilLabel(sum, 1, 2, "Total Orders:");
  const drinkList = readNonEmptyColumn(sum, 3, 2);
  const totalOrders = findFooterValue(sum, "Total Orders:");
  const totalOrdersText = (totalOrders === null || totalOrders === undefined) ? "" : String(totalOrders);

  const lines = [];
  lines.push(title);
  lines.push("");
  lines.push("Food Orders");
  lines.push(...foodList);
  lines.push("");
  lines.push(`Total Orders: ${totalOrdersText}`);
  lines.push("");
  lines.push("Drinks Orders");
  lines.push(...drinkList);

  const text = lines.join("\n");
  const builder = SpreadsheetApp.newRichTextValue().setText(text);

  applyBoldToLine(builder, text, 1);
  applyBoldToSubstring(builder, text, "Food Orders");
  applyBoldToSubstring(builder, text, "Drinks Orders");

  return { rich: builder.build() };
}

function readNonEmptyColumn(sheet, col, startRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return [];
  const values = sheet.getRange(startRow, col, lastRow - startRow + 1, 1).getValues().flat();
  return values
    .map(v => (v === null || v === undefined) ? "" : String(v).trim())
    .filter(v => v !== "");
}

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
    start = end + 1;
  }
}

function applyBoldToSubstring(builder, text, substr) {
  const idx = text.indexOf(substr);
  if (idx === -1) return;
  builder.setTextStyle(
    idx,
    idx + substr.length,
    SpreadsheetApp.newTextStyle().setBold(true).build()
  );
}

function readColumnUntilLabel(sheet, col, startRow, stopLabel) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return [];

  const values = sheet.getRange(startRow, col, lastRow - startRow + 1, 1).getValues().flat();
  const out = [];

  for (const v of values) {
    const s = (v === null || v === undefined) ? "" : String(v).trim();
    if (!s) continue;
    if (s === stopLabel) break;
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

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());

  const protectedLetters = new Set(["C", "U", "AT", "AU", "AV"]);
  const protectedIndexes = new Set([...protectedLetters].map(colA1ToIndex_));

  const sensitiveHeaderRegex = /(billing|shipping|address|addr|postcode|post code|zip|eircode|city|county|state|country|phone|mobile|email|e-mail|customer|contact|first name|last name|surname|company|order notes|note)/i;

  const colsToClear = [];
  for (let i = 0; i < headers.length; i++) {
    const colIndex = i + 1;
    if (protectedIndexes.has(colIndex)) continue;

    const header = headers[i];
    if (!header) continue;

    if (sensitiveHeaderRegex.test(header)) colsToClear.push(colIndex);
  }

  for (const colIndex of colsToClear) {
    sheet.getRange(2, colIndex, lastRow - 1, 1).clearContent();
  }
}

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
  return t === "" || t.startsWith("didn't") || t.startsWith("didnt") || t.startsWith("na");
}