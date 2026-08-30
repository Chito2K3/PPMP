/**
 * ============================================================================
 * PPMP MOBILE PHARMACY PRODUCT EVALUATION SYSTEM - BACKEND ENGINE
 * ============================================================================
 * 
 * Google Apps Script backend for handling sheet initialization, master drug list
 * syncing, evaluator tab mirroring, score calculation, and consensus dashboard.
 * 
 * Author: Antigravity Team
 * Version: 1.0.0
 */

// Global Constants
var SHEET_MEDICINE_MASTER = "Medicine_Master";
var SHEET_EVALUATIONS_MASTER = "Evaluations_Master";
var SHEET_EVALUATOR_1 = "Evaluator 1";
var SHEET_EVALUATOR_2 = "Evaluator 2";
var SHEET_EVALUATOR_3 = "Evaluator 3";
var SHEET_SUMMARY = "Consolidated_Summary";
var SHEET_CHECKLIST_REPORT = "Checklist_Report";
var SHEET_CHECKLIST_HORIZONTAL = "Checklist_Report_Horizontal";
var SHEET_INDICATIVE_SOURCE = "indicative 1";

// Column Definitions for Evaluations_Master (1-indexed for Apps Script range operations)
var EVAL_HEADERS = [
  "Evaluation_ID", "Timestamp", "Evaluator", "Generic_Name", "Brand_Name", "Manufacturer",
  "P1_01_Brand_Name", "P1_02_Generic_Name", "P1_03_Dosage_Form_Strength", "P1_04_Manufacturer_Details",
  "P1_05_CPR_FDA_Registration", "P1_06_Batch_Lot_Number", "P1_07_Manufacturing_Date", "P1_08_Expiration_Date",
  "P1_09_Storage_Conditions", "P1_10_Rx_Symbol", "P1_11_Net_Content", "P1_12_Language_Legibility",
  "P1_13_Outer_Package", "P1_14_Inner_Package", "P1_15_Package_Insert", "P1_16_Barcode_QR",
  "P1_17_Tamper_Evident_Seal", "P1_18_Special_Warnings", "P1_19_FDA_Compliance",
  "P2_01_Container_Integrity", "P2_02_Closure_Seal", "P2_03_Blister_Packaging", "P2_04_Physical_Appearance",
  "P2_05_Dosing_Graduation", "P2_06_Dispensing_Ease",
  "Requires_Reconstitution", "P3_01_Reconstitution_Clarity", "P3_02_Dissolution_Time", "P3_03_Foaming_Behavior",
  "P3_04_Solution_Color", "P3_05_Diluent_Compatibility", "P3_06_Syringe_Passability", "P3_07_Post_Reconstitution_Stability",
  "P3_08_Filter_Needle_Req", "P3_09_Particulate_Absence",
  "Part_I_Score", "Part_II_Score", "Remarks", "Recommendation"
];

var SUMMARY_HEADERS = [
  "Generic_Name", "Brand_Name", "Manufacturer",
  "Eval1_Recommendation", "Eval2_Recommendation", "Eval3_Recommendation",
  "Eval1_PartI_Score", "Eval1_PartII_Score",
  "Eval2_PartI_Score", "Eval2_PartII_Score",
  "Eval3_PartI_Score", "Eval3_PartII_Score",
  "Consensus_Status", "Consensus_Detail", "Last_Updated"
];

var OFFICIAL_PART1_ITEMS = [
  "Product Name",
  "Dosage Form and Strength",
  "Pharmacologic Category",
  "Formulation / Composition",
  "Indication(s)",
  "Dosage and Mode of Administration",
  "Contraindication(s), Precaution(s), Warning(s)",
  "Drug\u2013Drug / Drug\u2013Food Interactions",
  "Adverse Drug Reaction(s)",
  "Overdose and Treatment Information",
  "Storage Condition(s)",
  "Net Content / Pack Size",
  "Name & Address of Marketing Authorization Holder",
  "Name & Address of Manufacturer",
  "Rx Symbol & Prescription Caution Statement (if applicable)",
  "ADR Reporting Statement",
  "Registration Number",
  "Batch / Lot Number",
  "Date of Manufacture & Expiration Date"
];

var OFFICIAL_PART2_ITEMS = [
  "Inner label is identical to the outer label",
  "Drug name, dosage form, strength, batch/lot number, manufacture date, and expiry date are clearly readable on the container or inner packaging",
  "For blister or aluminum foil packs, expiry date, drugs name and dosage form is printed on each individual unit",
  "No leakage observed in IV fluids or other parenteral products through closures (rubber stoppers, caps, seals) or infusion sets",
  "Rubber stoppers (single-port and dual/twin-port) of IV fluid containers are durable yet easy to puncture",
  "Ease of opening, dispensing, and overall container integrity"
];

/**
 * Custom menu created when opening the Google Spreadsheet.
 */
function onOpen() {
  createCustomMenu();
}

/**
 * Creates the custom UI menu in Google Sheets.
 */
function createCustomMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🏥 PPMP Evaluation")
    .addItem("🚀 Initialize Evaluation Workbook", "initializeEvaluationWorkbook")
    .addItem("📊 Initialize / Refresh Horizontal Report", "createHorizontalReportSheet")
    .addItem("📑 Initialize / Reset Printable Checklist Viewer", "createChecklistReportSheet")
    .addItem("🔄 Sync Master Drug List", "syncMasterDrugList")
    .addItem("📊 Refresh Consolidated Summary", "refreshConsolidatedSummary")
    .addSeparator()
    .addItem("📄 Export Current Checklist to PDF", "exportCurrentChecklistPdf")
    .addItem("⚡ Setup Auto-Sync Triggers", "setupTriggers")
    .addToUi();
}

/**
 * Initializes all required sheets, headers, formatting, and data validations.
 */
function initializeEvaluationWorkbook() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  try {
    // 1. Medicine_Master
    var medSheet = getOrCreateSheet(ss, SHEET_MEDICINE_MASTER);
    setupSheetHeaders(medSheet, ["Generic_Name", "Category_Type", "Procurement_Mode"], "#1B365D");
    
    // 2. Evaluations_Master
    var masterSheet = getOrCreateSheet(ss, SHEET_EVALUATIONS_MASTER);
    setupSheetHeaders(masterSheet, EVAL_HEADERS, "#1B365D");
    applyEvaluationsValidations(masterSheet);
    
    // 3-5. Evaluator tabs 1, 2, 3
    var evalSheets = [SHEET_EVALUATOR_1, SHEET_EVALUATOR_2, SHEET_EVALUATOR_3];
    evalSheets.forEach(function(sheetName) {
      var s = getOrCreateSheet(ss, sheetName);
      setupSheetHeaders(s, EVAL_HEADERS, "#2C3E50");
      applyEvaluationsValidations(s);
    });
    
    // 6. Consolidated_Summary
    var summarySheet = getOrCreateSheet(ss, SHEET_SUMMARY);
    setupSheetHeaders(summarySheet, SUMMARY_HEADERS, "#004B49");
    
    ui.alert("Success", "Workbook initialized successfully with all 6 required sheets, headers, and validations!", ui.ButtonSet.OK);
  } catch (err) {
    Logger.log("Error in initializeEvaluationWorkbook: " + err.toString());
    ui.alert("Error Initializing Workbook", err.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Helper to retrieve an existing sheet or create a new one.
 */
function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/**
 * Helper to set up sheet headers, header formatting, and freeze row 1.
 */
function setupSheetHeaders(sheet, headers, headerColor) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground(headerColor)
             .setFontColor("#FFFFFF")
             .setFontWeight("bold")
             .setHorizontalAlignment("center")
             .setVerticalAlignment("middle");
  
  sheet.setRowHeight(1, 35);
  sheet.setFrozenRows(1);
}

/**
 * Applies Data Validation rules (Dropdowns for Evaluator, Yes/No/N/A, Recommendation)
 */
function applyEvaluationsValidations(sheet) {
  var ruleEvaluator = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Evaluator 1", "Evaluator 2", "Evaluator 3"], true)
    .setAllowInvalid(false)
    .build();
    
  var ruleYesNoNA = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Yes", "No", "N/A"], true)
    .setAllowInvalid(false)
    .build();

  var ruleYesNo = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Yes", "No"], true)
    .setAllowInvalid(false)
    .build();

  var ruleRecommendation = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Recommended", "Not Recommended"], true)
    .setAllowInvalid(false)
    .build();

  var maxRows = 500; // Apply validation rule to first 500 rows for performance
  
  // Evaluator Column C (index 3)
  sheet.getRange(2, 3, maxRows, 1).setDataValidation(ruleEvaluator);
  
  // Part I questions cols G-Y (indices 7 to 25)
  sheet.getRange(2, 7, maxRows, 19).setDataValidation(ruleYesNoNA);
  
  // Part II questions cols Z-AE (indices 26 to 31)
  sheet.getRange(2, 26, maxRows, 6).setDataValidation(ruleYesNoNA);
  
  // Requires_Reconstitution col AF (index 32)
  sheet.getRange(2, 32, maxRows, 1).setDataValidation(ruleYesNo);
  
  // Part III questions cols AG-AO (indices 33 to 41)
  sheet.getRange(2, 33, maxRows, 9).setDataValidation(ruleYesNoNA);
  
  // Recommendation col AS (index 45)
  sheet.getRange(2, 45, maxRows, 1).setDataValidation(ruleRecommendation);
}

/**
 * Synchronizes generic drug names and categories from 'framework agreement list' or 'indicative 1'
 * sheet into 'Medicine_Master'. Supports side-by-side category tables (e.g., Column C, Column J, etc.).
 * Preserves any existing custom category and procurement mode metadata if present.
 */
function syncMasterDrugList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  // 1. Find source sheet: check 'framework agreement list', then 'indicative 1', or search
  var sourceSheet = ss.getSheetByName("framework agreement list");
  if (!sourceSheet) {
    sourceSheet = ss.getSheetByName("indicative 1");
  }
  if (!sourceSheet) {
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var sName = sheets[i].getName().trim().toLowerCase();
      if (sName.indexOf("framework agreement") !== -1 || sName.indexOf("indicative") !== -1 || sName.indexOf("ppmp") !== -1) {
        sourceSheet = sheets[i];
        break;
      }
    }
  }
  
  if (!sourceSheet) {
    ui.alert("Source Sheet Not Found", "Could not find a tab named 'framework agreement list' or 'indicative 1'. Please make sure your drug list tab exists in this workbook.", ui.ButtonSet.OK);
    return;
  }
  
  var targetSheet = getOrCreateSheet(ss, SHEET_MEDICINE_MASTER);
  
  // Read existing custom metadata (Category & Procurement) if present
  var existingMap = {};
  if (targetSheet.getLastRow() > 1) {
    var existingData = targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, 3).getValues();
    existingData.forEach(function(row) {
      var drug = row[0] ? row[0].toString().trim() : "";
      if (drug) {
        existingMap[drug] = {
          category: row[1] || "",
          procurement: row[2] || ""
        };
      }
    });
  }
  
  var lastRow = sourceSheet.getLastRow();
  var lastCol = sourceSheet.getLastColumn();
  
  if (lastRow < 4) {
    ui.alert("No Data Found", "Source sheet has fewer than 4 rows.", ui.ButtonSet.OK);
    return;
  }
  
  var fullData = sourceSheet.getRange(1, 1, lastRow, lastCol).getValues();
  var extractedItems = [];
  var seen = {};
  
  // Determine if this is the multi-column 'framework agreement list' layout or single-column layout
  var drugColIndices = [];
  
  // Scan headers (rows 1 to 8) to find columns with "item / service" or "type and nature"
  for (var c = 0; c < lastCol; c++) {
    var isDrugCol = false;
    for (var r = 0; r < Math.min(8, lastRow); r++) {
      var cellVal = fullData[r][c] ? fullData[r][c].toString().toLowerCase() : "";
      if ((cellVal.indexOf("item") !== -1 && cellVal.indexOf("service") !== -1) || cellVal.indexOf("type and nature") !== -1) {
        isDrugCol = true;
        break;
      }
    }
    // Hardcoded fallback check for Column C (index 2) and Column J (index 9) for Framework Agreement layout
    if (isDrugCol || c === 2 || c === 9) {
      // Exclude "Item No." and "Unit" columns
      var colHeaderSample = (fullData[6] && fullData[6][c] ? fullData[6][c].toString() : "") + " " + (fullData[7] && fullData[7][c] ? fullData[7][c].toString() : "");
      if (colHeaderSample.toLowerCase().indexOf("item no") === -1 && colHeaderSample.toLowerCase().indexOf("unit") === -1) {
        drugColIndices.push(c);
      }
    }
  }
  
  // Filter unique column indices
  drugColIndices = drugColIndices.filter(function(v, idx, self) { return self.indexOf(v) === idx; });
  
  // Fallback to Column A (index 0) if no multi-column headers match
  if (drugColIndices.length === 0) {
    drugColIndices = [0];
  }
  
  // Start row detection: if row 9 exists and has items, use start row 9 (index 8), else start row 4 (index 3)
  var startRowIndex = (lastRow >= 9) ? 8 : 3;
  
  drugColIndices.forEach(function(colIdx) {
    // Detect Category from Row 2 (index 1) above this column (checks up to 6 columns back for merged cells)
    var category = "";
    if (fullData[1]) {
      for (var offset = 0; offset <= 6; offset++) {
        var checkIdx = colIdx - offset;
        if (checkIdx >= 0 && fullData[1][checkIdx]) {
          var val = fullData[1][checkIdx].toString().trim();
          if (val && 
              val.toUpperCase().indexOf("PR ") === -1 && 
              val.toUpperCase().indexOf("FRAMEWORK") === -1 &&
              val.toUpperCase().indexOf("LIST") === -1) {
            category = val;
            break;
          }
        }
      }
    }
    if (!category) category = "Pharmaceutical";
    
    // Scan items down the column
    for (var r = startRowIndex; r < lastRow; r++) {
      var drugName = fullData[r][colIdx] ? fullData[r][colIdx].toString().trim() : "";
      var lowerName = drugName.toLowerCase();
      
      // Filter out non-drug text, headers, and numbers
      if (drugName && 
          lowerName.indexOf("type and nature") === -1 &&
          lowerName.indexOf("item / service") === -1 &&
          lowerName.indexOf("framework agreement") === -1 &&
          lowerName.indexOf("item no") === -1 &&
          !seen[lowerName]) {
        seen[lowerName] = true;
        
        var cat = existingMap[drugName] && existingMap[drugName].category ? existingMap[drugName].category : category;
        var proc = existingMap[drugName] && existingMap[drugName].procurement ? existingMap[drugName].procurement : "Framework Agreement / Public Bidding";
        
        extractedItems.push([drugName, cat, proc]);
      }
    }
  });
  
  // Sort extracted items alphabetically by generic name
  extractedItems.sort(function(a, b) { return a[0].localeCompare(b[0]); });
  
  // Clear existing content (except row 1) and write
  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, 3).clearContent();
  }
  
  if (extractedItems.length > 0) {
    targetSheet.getRange(2, 1, extractedItems.length, 3).setValues(extractedItems);
  }
  
  ui.alert("Sync Complete", "Successfully synchronized " + extractedItems.length + " unique generic drugs into " + SHEET_MEDICINE_MASTER + " from sheet tab '" + sourceSheet.getName() + "'!", ui.ButtonSet.OK);
}

/**
 * Calculates Part I and Part II compliance percentage scores.
 * Formula: Yes_count / (Total_Questions - NA_Count) * 100%
 */
function calculateScoresForRow(rowValues) {
  // Part I: Indices 6 to 24 (19 items)
  var p1Yes = 0;
  var p1NA = 0;
  for (var i = 6; i <= 24; i++) {
    var val = rowValues[i] ? rowValues[i].toString().trim().toUpperCase() : "";
    if (val === "YES") p1Yes++;
    if (val === "N/A" || val === "NA") p1NA++;
  }
  var p1Eligible = 19 - p1NA;
  var p1Score = p1Eligible > 0 ? ((p1Yes / p1Eligible) * 100).toFixed(1) + "%" : "100.0%";
  
  // Part II: Indices 25 to 30 (6 items)
  var p2Yes = 0;
  var p2NA = 0;
  for (var j = 25; j <= 30; j++) {
    var val2 = rowValues[j] ? rowValues[j].toString().trim().toUpperCase() : "";
    if (val2 === "YES") p2Yes++;
    if (val2 === "N/A" || val2 === "NA") p2NA++;
  }
  var p2Eligible = 6 - p2NA;
  var p2Score = p2Eligible > 0 ? ((p2Yes / p2Eligible) * 100).toFixed(1) + "%" : "100.0%";
  
  return {
    partIScore: p1Score,
    partIIScore: p2Score
  };
}

/**
 * Installed trigger handler or onEdit listener for sync & mirror logic.
 */
function onEvaluationsEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_EVALUATIONS_MASTER) return;
  
  var row = e.range.getRow();
  if (row < 2) return; // Skip header
  
  processMasterRow(sheet, row);
  refreshConsolidatedSummary();
  syncHorizontalReport(true);
}

/**
 * Processes a single row in Evaluations_Master: auto-calculates scores and mirrors to evaluator sheet.
 */
function processMasterRow(masterSheet, rowNum) {
  var rowRange = masterSheet.getRange(rowNum, 1, 1, EVAL_HEADERS.length);
  var values = rowRange.getValues()[0];
  
  // Auto-assign Evaluation ID if blank
  if (!values[0]) {
    values[0] = "EVAL-" + Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd-HHmmss") + "-" + Math.floor(Math.random() * 1000);
    masterSheet.getRange(rowNum, 1).setValue(values[0]);
  }
  
  // Auto-assign Timestamp if blank
  if (!values[1]) {
    values[1] = new Date();
    masterSheet.getRange(rowNum, 2).setValue(values[1]);
  }
  
  // Calculate Scores
  var scores = calculateScoresForRow(values);
  values[41] = scores.partIScore; // Part_I_Score col AP (42)
  values[42] = scores.partIIScore; // Part_II_Score col AQ (43)
  
  masterSheet.getRange(rowNum, 42).setValue(scores.partIScore);
  masterSheet.getRange(rowNum, 43).setValue(scores.partIIScore);
  
  // Mirror to Evaluator tab based on Evaluator column (Index 2)
  var evaluator = values[2] ? values[2].toString().trim() : "";
  if (evaluator === "Evaluator 1" || evaluator === "Evaluator 2" || evaluator === "Evaluator 3") {
    mirrorToEvaluatorSheet(evaluator, values);
  }
}

/**
 * Mirrors/updates a record in the specific Evaluator sheet (matching by Evaluation_ID).
 */
function mirrorToEvaluatorSheet(sheetName, rowValues) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var evalSheet = getOrCreateSheet(ss, sheetName);
  var evalId = rowValues[0];
  
  if (!evalId) return;
  
  var lastRow = evalSheet.getLastRow();
  var targetRow = -1;
  
  if (lastRow > 1) {
    var existingIds = evalSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var r = 0; r < existingIds.length; r++) {
      if (existingIds[r][0] && existingIds[r][0].toString() === evalId) {
        targetRow = r + 2;
        break;
      }
    }
  }
  
  if (targetRow > 1) {
    // Update existing row
    evalSheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    // Append new row
    evalSheet.appendRow(rowValues);
  }
}

/**
 * Rebuilds the Consolidated_Summary sheet.
 * Groups by Generic_Name + Brand_Name + Manufacturer and aggregates responses from Evaluators 1, 2, and 3.
 */
function refreshConsolidatedSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(SHEET_EVALUATIONS_MASTER);
  var summarySheet = getOrCreateSheet(ss, SHEET_SUMMARY);
  
  if (!masterSheet || masterSheet.getLastRow() < 2) {
    // Clear summary except headers
    if (summarySheet.getLastRow() > 1) {
      summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, SUMMARY_HEADERS.length).clear();
    }
    return;
  }
  
  var data = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, EVAL_HEADERS.length).getValues();
  
  // Dictionary structure: key = generic|brand|manufacturer
  var grouped = {};
  
  data.forEach(function(row) {
    var evalId = row[0];
    var timestamp = row[1];
    var evaluator = row[2] ? row[2].toString().trim() : "";
    var generic = row[3] ? row[3].toString().trim() : "";
    var brand = row[4] ? row[4].toString().trim() : "";
    var manufacturer = row[5] ? row[5].toString().trim() : "";
    var p1Score = row[41] || "";
    var p2Score = row[42] || "";
    var recommendation = row[44] ? row[44].toString().trim() : "";
    
    if (!generic && !brand) return;
    
    var key = (generic + "||" + brand + "||" + manufacturer).toLowerCase();
    
    if (!grouped[key]) {
      grouped[key] = {
        generic: generic,
        brand: brand,
        manufacturer: manufacturer,
        evaluators: {}
      };
    }
    
    // Store latest evaluation per evaluator
    if (evaluator) {
      grouped[key].evaluators[evaluator] = {
        recommendation: recommendation,
        p1Score: p1Score,
        p2Score: p2Score,
        timestamp: timestamp
      };
    }
  });
  
  var nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
  var summaryRows = [];
  var statusColors = [];
  
  Object.keys(grouped).forEach(function(k) {
    var item = grouped[k];
    var e1 = item.evaluators["Evaluator 1"] || { recommendation: "Pending", p1Score: "-", p2Score: "-" };
    var e2 = item.evaluators["Evaluator 2"] || { recommendation: "Pending", p1Score: "-", p2Score: "-" };
    var e3 = item.evaluators["Evaluator 3"] || { recommendation: "Pending", p1Score: "-", p2Score: "-" };
    
    var recs = [e1.recommendation, e2.recommendation, e3.recommendation];
    var recCount = 0;
    var recommendedCount = 0;
    var notRecommendedCount = 0;
    
    recs.forEach(function(r) {
      if (r === "Recommended") {
        recCount++;
        recommendedCount++;
      } else if (r === "Not Recommended") {
        recCount++;
        notRecommendedCount++;
      }
    });
    
    var consensusStatus = "";
    var consensusDetail = "";
    var bgColor = "#FFFFFF";
    
    if (recCount < 3) {
      consensusStatus = "Pending (" + recCount + "/3 Complete)";
      consensusDetail = recCount + " of 3 evaluators have submitted scores.";
      bgColor = "#F2F4F4"; // Light gray
    } else if (recommendedCount === 3) {
      consensusStatus = "Unanimous Recommended";
      consensusDetail = "All 3 evaluators voted Recommended.";
      bgColor = "#D4EFDF"; // Light green
    } else if (notRecommendedCount === 3) {
      consensusStatus = "Unanimous Not Recommended";
      consensusDetail = "All 3 evaluators voted Not Recommended.";
      bgColor = "#FADBD8"; // Light red
    } else {
      consensusStatus = "Split Decision";
      consensusDetail = recommendedCount + " Recommended, " + notRecommendedCount + " Not Recommended.";
      bgColor = "#FCF3CF"; // Light amber/yellow
    }
    
    summaryRows.push([
      item.generic, item.brand, item.manufacturer,
      e1.recommendation, e2.recommendation, e3.recommendation,
      e1.p1Score, e1.p2Score,
      e2.p1Score, e2.p2Score,
      e3.p1Score, e3.p2Score,
      consensusStatus, consensusDetail, nowStr
    ]);
    
    statusColors.push(bgColor);
  });
  
  // Write to Consolidated_Summary sheet
  if (summarySheet.getLastRow() > 1) {
    summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, SUMMARY_HEADERS.length).clear();
  }
  
  if (summaryRows.length > 0) {
    var targetRange = summarySheet.getRange(2, 1, summaryRows.length, SUMMARY_HEADERS.length);
    targetRange.setValues(summaryRows);
    
    // Apply background colors to rows
    for (var r = 0; r < statusColors.length; r++) {
      summarySheet.getRange(r + 2, 1, 1, SUMMARY_HEADERS.length).setBackground(statusColors[r]);
    }
    
    // Format text alignment
    summarySheet.getRange(2, 4, summaryRows.length, 10).setHorizontalAlignment("center");
  }
}

/**
 * Programmatically installs an onEdit trigger for automatic updates.
 */
function setupTriggers() {
  var ui = SpreadsheetApp.getUi();
  
  // Clear existing triggers for this project to prevent duplicates
  var allTriggers = ScriptApp.getProjectTriggers();
  allTriggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "onEvaluationsEdit") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Install trigger
  ScriptApp.newTrigger("onEvaluationsEdit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
    
  ui.alert("Triggers Setup", "Successfully installed 'onEvaluationsEdit' trigger on spreadsheet edit!", ui.ButtonSet.OK);
}

/**
 * Creates or resets the official Printable Checklist Viewer tab (Checklist_Report).
 * Dynamically plots evaluator responses (☑ / ☐) based on selected Drug & Evaluator.
 */
function createChecklistReportSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  try {
    var sheet = getOrCreateSheet(ss, SHEET_CHECKLIST_REPORT);
    sheet.clear();
    sheet.clearFormats();
    
    // Set column widths
    sheet.setColumnWidth(1, 50);   // Col A: Part / Item #
    sheet.setColumnWidth(2, 45);   // Col B: Item #
    sheet.setColumnWidth(3, 400);  // Col C: Criteria Description
    sheet.setColumnWidth(4, 55);   // Col D: Yes
    sheet.setColumnWidth(5, 55);   // Col E: No
    sheet.setColumnWidth(6, 55);   // Col F: N/A
    sheet.setColumnWidth(7, 180);  // Col G: Remarks
    
    // 1. Title Block
    sheet.getRange("A1:G1").merge()
      .setValue("PRODUCT EVALUATION CHECKLIST")
      .setFontWeight("bold")
      .setFontSize(14)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
      
    sheet.getRange("A2:G2").merge()
      .setValue("PHARMACY BIDS & AWARDS COMMITTEE — TECHNICAL EVALUATION REPORT")
      .setFontSize(9)
      .setFontColor("#555555")
      .setHorizontalAlignment("center");
      
    // 2. Interactive Selection Controls
    sheet.getRange("A3:B3").merge().setValue("🔍 Select Drug:").setFontWeight("bold");
    sheet.getRange("C3").setBackground("#E8F0FE").setFontWeight("bold");
    
    // Set data validation for Drug (from Medicine_Master)
    var medSheet = ss.getSheetByName(SHEET_MEDICINE_MASTER);
    if (medSheet && medSheet.getLastRow() > 1) {
      var drugRule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(medSheet.getRange("A2:A" + medSheet.getLastRow()), true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange("C3").setDataValidation(drugRule);
    }
    
    sheet.getRange("D3:E3").merge().setValue("👤 Evaluator:").setFontWeight("bold").setHorizontalAlignment("right");
    sheet.getRange("F3:G3").merge().setBackground("#E8F0FE").setFontWeight("bold");
    var evalRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Evaluator 1", "Evaluator 2", "Evaluator 3"], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange("F3:G3").setDataValidation(evalRule);
    
    // Default initial values
    if (medSheet && medSheet.getLastRow() > 1) {
      sheet.getRange("C3").setValue(medSheet.getRange("A2").getValue());
    }
    sheet.getRange("F3:G3").setValue("Evaluator 1");
    
    // 3. Metadata Header (Supplier / Generic / Brand)
    sheet.getRange("A5:B5").merge().setValue("Name of Supplier:").setFontWeight("bold");
    sheet.getRange("C5:G5").merge().setFormula('=IFERROR(INDEX(Evaluations_Master!F:F, MATCH(1, (Evaluations_Master!D:D=$C$3)*(Evaluations_Master!C:C=$F$3), 0)), "—")').setFontWeight("bold");
    
    sheet.getRange("A6:B6").merge().setValue("Generic Name:").setFontWeight("bold");
    sheet.getRange("C6:G6").merge().setFormula('=$C$3').setFontWeight("bold");
    
    sheet.getRange("A7:B7").merge().setValue("Brand Name:").setFontWeight("bold");
    sheet.getRange("C7:G7").merge().setFormula('=IFERROR(INDEX(Evaluations_Master!E:E, MATCH(1, (Evaluations_Master!D:D=$C$3)*(Evaluations_Master!C:C=$F$3), 0)), "—")').setFontWeight("bold");
    
    // 4. Part I Table Header
    var headerRange = sheet.getRange("A9:G9");
    headerRange.setValues([["Part", "#", "PRODUCT SAMPLE EVALUATION (Part I)", "Yes", "No", "N/A", "Remarks"]]);
    headerRange.setBackground("#1B365D").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("C9").setHorizontalAlignment("left");
    
    // 5. Part I Items (19 rows: Row 10 to 28)
    var p1Rows = [];
    var p1Formulas = [];
    for (var i = 0; i < OFFICIAL_PART1_ITEMS.length; i++) {
      var itemNum = i + 1;
      var colLetter = String.fromCharCode(71 + i); // Col G=71, H=72... for P1_01 to P1_19
      // Handle column lettering past Z if necessary (G to Y is 7 to 25)
      var colRef = getColumnLetter(7 + i); // 7 is col G
      
      var yesFormula = '=IF(IFERROR(INDEX(Evaluations_Master!' + colRef + ':' + colRef + ', MATCH(1, (Evaluations_Master!$D:$D=$C$3)*(Evaluations_Master!$C:$C=$F$3), 0)), "")="Yes", "☑", "☐")';
      var noFormula  = '=IF(IFERROR(INDEX(Evaluations_Master!' + colRef + ':' + colRef + ', MATCH(1, (Evaluations_Master!$D:$D=$C$3)*(Evaluations_Master!$C:$C=$F$3), 0)), "")="No", "☑", "☐")';
      var naFormula  = '=IF(IFERROR(INDEX(Evaluations_Master!' + colRef + ':' + colRef + ', MATCH(1, (Evaluations_Master!$D:$D=$C$3)*(Evaluations_Master!$C:$C=$F$3), 0)), "")="N/A", "☑", "☐")';
      
      p1Rows.push(["Part I", itemNum, OFFICIAL_PART1_ITEMS[i], "", "", "", ""]);
      p1Formulas.push([yesFormula, noFormula, naFormula]);
    }
    
    sheet.getRange(10, 1, 19, 7).setValues(p1Rows);
    for (var r = 0; r < p1Formulas.length; r++) {
      sheet.getRange(10 + r, 4, 1, 3).setFormulas([p1Formulas[r]]);
    }
    sheet.getRange(10, 1, 19, 2).setHorizontalAlignment("center");
    sheet.getRange(10, 4, 19, 3).setHorizontalAlignment("center").setFontSize(11);
    
    // Merge Part I column vertically
    sheet.getRange("A10:A28").merge().setVerticalAlignment("middle").setHorizontalAlignment("center").setFontWeight("bold");
    
    // 6. Part II Table Header & Rows (Row 29 header, rows 30 to 35)
    sheet.getRange("A29:G29").merge().setValue("Part II: Physical Packaging & Container Integrity").setBackground("#2C3E50").setFontColor("#FFFFFF").setFontWeight("bold");
    
    var p2Rows = [];
    var p2Formulas = [];
    for (var j = 0; j < OFFICIAL_PART2_ITEMS.length; j++) {
      var itemNum2 = j + 1;
      var colRef2 = getColumnLetter(26 + j); // Col Z=26 (P2_01)
      
      var yesFormula2 = '=IF(IFERROR(INDEX(Evaluations_Master!' + colRef2 + ':' + colRef2 + ', MATCH(1, (Evaluations_Master!$D:$D=$C$3)*(Evaluations_Master!$C:$C=$F$3), 0)), "")="Yes", "☑", "☐")';
      var noFormula2  = '=IF(IFERROR(INDEX(Evaluations_Master!' + colRef2 + ':' + colRef2 + ', MATCH(1, (Evaluations_Master!$D:$D=$C$3)*(Evaluations_Master!$C:$C=$F$3), 0)), "")="No", "☑", "☐")';
      var naFormula2  = '=IF(IFERROR(INDEX(Evaluations_Master!' + colRef2 + ':' + colRef2 + ', MATCH(1, (Evaluations_Master!$D:$D=$C$3)*(Evaluations_Master!$C:$C=$F$3), 0)), "")="N/A", "☑", "☐")';
      
      p2Rows.push(["Part II", itemNum2, OFFICIAL_PART2_ITEMS[j], "", "", "", ""]);
      p2Formulas.push([yesFormula2, noFormula2, naFormula2]);
    }
    
    sheet.getRange(30, 1, 6, 7).setValues(p2Rows);
    for (var r2 = 0; r2 < p2Formulas.length; r2++) {
      sheet.getRange(30 + r2, 4, 1, 3).setFormulas([p2Formulas[r2]]);
    }
    sheet.getRange(30, 1, 6, 2).setHorizontalAlignment("center");
    sheet.getRange(30, 4, 6, 3).setHorizontalAlignment("center").setFontSize(11);
    sheet.getRange("A30:A35").merge().setVerticalAlignment("middle").setHorizontalAlignment("center").setFontWeight("bold");
    
    // 7. Summary & Verdict Footer Block
    var startFooter = 37;
    sheet.getRange(startFooter, 1, 1, 7).merge().setValue("EVALUATION SUMMARY & RECOMMENDATION").setBackground("#1B365D").setFontColor("#FFFFFF").setFontWeight("bold");
    
    sheet.getRange(startFooter + 1, 1, 1, 2).merge().setValue("Part I Score:").setFontWeight("bold");
    sheet.getRange(startFooter + 1, 3).setFormula('=IFERROR(INDEX(Evaluations_Master!AP:AP, MATCH(1, (Evaluations_Master!D:D=$C$3)*(Evaluations_Master!C:C=$F$3), 0)), "—")').setFontWeight("bold");
    
    sheet.getRange(startFooter + 1, 4, 1, 2).merge().setValue("Part II Score:").setFontWeight("bold");
    sheet.getRange(startFooter + 1, 6, 1, 2).merge().setFormula('=IFERROR(INDEX(Evaluations_Master!AQ:AQ, MATCH(1, (Evaluations_Master!D:D=$C$3)*(Evaluations_Master!C:C=$F$3), 0)), "—")').setFontWeight("bold");
    
    sheet.getRange(startFooter + 2, 1, 1, 2).merge().setValue("Recommendation:").setFontWeight("bold");
    sheet.getRange(startFooter + 2, 3, 1, 5).merge().setFormula('=IFERROR(INDEX(Evaluations_Master!AS:AS, MATCH(1, (Evaluations_Master!D:D=$C$3)*(Evaluations_Master!C:C=$F$3), 0)), "—")').setFontWeight("bold").setFontSize(11);
    
    sheet.getRange(startFooter + 3, 1, 1, 2).merge().setValue("Overall Remarks:").setFontWeight("bold");
    sheet.getRange(startFooter + 3, 3, 1, 5).merge().setFormula('=IFERROR(INDEX(Evaluations_Master!AR:AR, MATCH(1, (Evaluations_Master!D:D=$C$3)*(Evaluations_Master!C:C=$F$3), 0)), "—")');
    
    // 8. Signatures Block
    var sigRow = startFooter + 5;
    sheet.getRange(sigRow, 1, 1, 3).merge().setFormula('="Evaluated By: " & IFERROR(INDEX(Evaluator_Accounts!C:C, MATCH($F$3, Evaluator_Accounts!B:B, 0)), "")').setFontWeight("bold");
    sheet.getRange(sigRow, 4, 1, 4).merge().setValue("Date & Signature: _______________________").setFontWeight("bold");
    sheet.getRange(sigRow + 1, 1, 1, 3).merge().setFormula('="Official Role: " & $F$3').setFontStyle("italic");
    
    // Apply grid borders to tables
    sheet.getRange("A9:G28").setBorder(true, true, true, true, true, true, "#333333", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange("A29:G35").setBorder(true, true, true, true, true, true, "#333333", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(startFooter, 1, 4, 7).setBorder(true, true, true, true, true, true, "#333333", SpreadsheetApp.BorderStyle.SOLID);
    
    ui.alert("Checklist Initialized", "The 'Checklist_Report' tab is now created!\n\nSelect any Drug in C3 and Evaluator in F3 to immediately view and print the populated evaluation.", ui.ButtonSet.OK);
  } catch (err) {
    Logger.log("Error in createChecklistReportSheet: " + err.toString());
    ui.alert("Error Creating Checklist Report", err.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Helper to convert 1-indexed column number to letter (e.g. 1 -> A, 27 -> AA)
 */
function getColumnLetter(columnNumber) {
  var temp = "";
  var letter = "";
  while (columnNumber > 0) {
    temp = (columnNumber - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    columnNumber = (columnNumber - temp - 1) / 26;
  }
  return letter;
}

/**
 * Exports the active Checklist_Report sheet view directly as a PDF file saved in Google Drive.
 */
function exportCurrentChecklistPdf() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName(SHEET_CHECKLIST_REPORT);
  
  if (!sheet) {
    ui.alert("Sheet Not Found", "Please initialize the Checklist Viewer first via the menu.", ui.ButtonSet.OK);
    return;
  }
  
  try {
    var drugName = sheet.getRange("C3").getValue() || "Evaluation";
    var evaluator = sheet.getRange("F3").getValue() || "Summary";
    var safeDrug = drugName.toString().replace(/[^a-zA-Z0-9_-]/g, "_").substring(0, 30);
    var fileName = "Checklist_" + safeDrug + "_" + evaluator + ".pdf";
    
    var pdfBlob = ss.getAs('application/pdf');
    pdfBlob.setName(fileName);
    
    var file = DriveApp.createFile(pdfBlob);
    
    ui.alert("PDF Exported Successfully", "PDF exported and saved to your Google Drive:\n\nFile: " + fileName + "\nLink: " + file.getUrl(), ui.ButtonSet.OK);
  } catch (err) {
    Logger.log("Error in exportCurrentChecklistPdf: " + err.toString());
    ui.alert("Error Exporting PDF", err.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Creates or updates the clean, presentation-ready Horizontal Evaluation Summary Report
 * (Checklist_Report_Horizontal). Mirrors Evaluations_Master horizontally but with official
 * human-readable headers, clean checkmarks (✓ / ✗ / —), scores, and evaluator names.
 */
function createHorizontalReportSheet() {
  syncHorizontalReport(false);
}

/**
 * Creates or updates the clean, presentation-ready Horizontal Evaluation Summary Report
 * (Checklist_Report_Horizontal). Mirrors Evaluations_Master horizontally but with official
 * human-readable headers, clean checkmarks (✓ / ✗ / —), scores, and evaluator names.
 */
function syncHorizontalReport(silent) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CHECKLIST_HORIZONTAL);
  
  if (!sheet) {
    if (silent) return;
    sheet = getOrCreateSheet(ss, SHEET_CHECKLIST_HORIZONTAL);
  }
  
  try {
    sheet.clear();
    sheet.clearFormats();
    
    // Build Evaluator Name lookup map from Evaluator_Accounts
    var evalMap = {};
    var acctsSheet = ss.getSheetByName("Evaluator_Accounts");
    if (acctsSheet && acctsSheet.getLastRow() > 1) {
      var acctsData = acctsSheet.getRange(2, 1, acctsSheet.getLastRow() - 1, 3).getValues();
      acctsData.forEach(function(row) {
        var role = row[1]; // Col B
        var name = row[2]; // Col C
        if (role && name) {
          evalMap[role.toString().trim()] = name.toString().trim();
        }
      });
    }
    
    // 1. Group Headers (Row 1)
    sheet.getRange("A1:G1").merge().setValue("IDENTIFICATION & METADATA")
      .setBackground("#1B365D").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
      
    sheet.getRange("H1:Z1").merge().setValue("PART I: LABELING & REGULATORY COMPLIANCE (19 CRITERIA)")
      .setBackground("#2C3E50").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
      
    sheet.getRange("AA1:AF1").merge().setValue("PART II: PHYSICAL PACKAGING & CONTAINER INTEGRITY (6 CRITERIA)")
      .setBackground("#34495E").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
      
    sheet.getRange("AG1:AJ1").merge().setValue("EVALUATION VERDICT & SCORES")
      .setBackground("#004B49").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
      
    // 2. Official Column Headers (Row 2)
    var row2Headers = [
      "#", "Generic Name", "Brand Name", "Manufacturer", "Evaluator Role", "Evaluator Name", "Date",
      // Part I items (19)
      "1. Product Name", "2. Dosage Form & Strength", "3. Pharmacologic Category", "4. Formulation / Composition",
      "5. Indication(s)", "6. Dosage & Mode of Admin", "7. Warnings & Precautions", "8. Drug Interactions",
      "9. Adverse Drug Reactions", "10. Overdose Info", "11. Storage Conditions", "12. Net Content / Pack Size",
      "13. Marketing Auth Holder", "14. Manufacturer Address", "15. Rx Caution Statement", "16. ADR Reporting",
      "17. Registration Number", "18. Batch / Lot Number", "19. Mfg & Expiry Date",
      // Part II items (6)
      "1. Inner Label Match", "2. Container Label Legibility", "3. Blister Pack Print",
      "4. Parenteral Leakage Check", "5. Rubber Stopper Puncture", "6. Dispensing Ease & Integrity",
      // Verdict
      "Part I Score", "Part II Score", "Recommendation", "Remarks"
    ];
    
    var hRange = sheet.getRange(2, 1, 1, row2Headers.length);
    hRange.setValues([row2Headers]);
    hRange.setBackground("#F1F5F9").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("B2:D2").setHorizontalAlignment("left");
    sheet.getRange("F2").setHorizontalAlignment("left");
    sheet.getRange("AJ2").setHorizontalAlignment("left");
    
    // Set Column Widths
    sheet.setColumnWidth(1, 40);   // #
    sheet.setColumnWidth(2, 220);  // Generic Name
    sheet.setColumnWidth(3, 140);  // Brand Name
    sheet.setColumnWidth(4, 140);  // Manufacturer
    sheet.setColumnWidth(5, 110);  // Evaluator Role
    sheet.setColumnWidth(6, 170);  // Evaluator Name
    sheet.setColumnWidth(7, 95);   // Date
    for (var c = 8; c <= 32; c++) {
      sheet.setColumnWidth(c, 75); // Part I & II criteria checkmark cols
    }
    sheet.setColumnWidth(33, 90);  // Part I Score
    sheet.setColumnWidth(34, 90);  // Part II Score
    sheet.setColumnWidth(35, 130); // Recommendation
    sheet.setColumnWidth(36, 220); // Remarks
    
    // Freeze top 2 header rows and left 4 identifying columns
    sheet.setFrozenRows(2);
    sheet.setFrozenColumns(4);
    
    // 3. Read and populate data from Evaluations_Master
    var masterSheet = ss.getSheetByName(SHEET_EVALUATIONS_MASTER);
    if (!masterSheet || masterSheet.getLastRow() <= 1) {
      if (!silent) {
        SpreadsheetApp.getUi().alert("Horizontal Report Initialized", "The 'Checklist_Report_Horizontal' tab is ready!\n\nNo evaluation records found yet. Once evaluations are submitted, this sheet will populate automatically.", SpreadsheetApp.getUi().ButtonSet.OK);
      }
      return;
    }
    
    var masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, EVAL_HEADERS.length).getValues();
    var reportRows = [];
    var recColors = [];
    
    for (var i = 0; i < masterData.length; i++) {
      var mRow = masterData[i];
      var evalId = mRow[0];
      if (!evalId) continue;
      
      var rawTimestamp = mRow[1];
      var dateStr = "";
      if (rawTimestamp instanceof Date) {
        dateStr = Utilities.formatDate(rawTimestamp, "GMT+8", "yyyy-MM-dd");
      } else if (rawTimestamp) {
        dateStr = rawTimestamp.toString().split(" ")[0];
      }
      
      var role = mRow[2] ? mRow[2].toString().trim() : "";
      var genericName = mRow[3] || "";
      var brandName = mRow[4] || "";
      var manufacturer = mRow[5] || "";
      var evalName = evalMap[role] || role;
      
      // Transform Part I items (indices 6 to 24 -> 19 items)
      var p1Transformed = [];
      for (var p1 = 6; p1 <= 24; p1++) {
        var val1 = mRow[p1] ? mRow[p1].toString().trim() : "";
        p1Transformed.push(formatCheckmark(val1));
      }
      
      // Transform Part II items (indices 25 to 30 -> 6 items)
      var p2Transformed = [];
      for (var p2 = 25; p2 <= 30; p2++) {
        var val2 = mRow[p2] ? mRow[p2].toString().trim() : "";
        p2Transformed.push(formatCheckmark(val2));
      }
      
      var p1Score = mRow[41] !== undefined ? mRow[41] : "";
      var p2Score = mRow[42] !== undefined ? mRow[42] : "";
      var remarks = mRow[43] || "";
      var recommendation = mRow[44] || "";
      
      var newRow = [
        (reportRows.length + 1), genericName, brandName, manufacturer, role, evalName, dateStr
      ].concat(p1Transformed).concat(p2Transformed).concat([p1Score, p2Score, recommendation, remarks]);
      
      reportRows.push(newRow);
      
      // Color badge for recommendation
      if (recommendation === "Recommended") {
        recColors.push("#D4EFDF"); // Light green
      } else if (recommendation === "Not Recommended") {
        recColors.push("#FADBD8"); // Light red
      } else {
        recColors.push("#FFFFFF");
      }
    }
    
    if (reportRows.length > 0) {
      var dataRange = sheet.getRange(3, 1, reportRows.length, row2Headers.length);
      dataRange.setValues(reportRows);
      
      // Apply zebra styling and centered alignment
      for (var r = 0; r < reportRows.length; r++) {
        var rowNum = 3 + r;
        var rowBg = (r % 2 === 0) ? "#FFFFFF" : "#F8FAFC";
        sheet.getRange(rowNum, 1, 1, row2Headers.length).setBackground(rowBg);
        
        // Highlight recommendation badge
        sheet.getRange(rowNum, 35).setBackground(recColors[r]).setFontWeight("bold");
      }
      
      // Center-align checklist checkmark columns (Col H to AF, which is 8 to 32)
      sheet.getRange(3, 1, reportRows.length, 1).setHorizontalAlignment("center");
      sheet.getRange(3, 5, reportRows.length, 1).setHorizontalAlignment("center");
      sheet.getRange(3, 7, reportRows.length, 28).setHorizontalAlignment("center").setFontSize(10);
      
      // Apply clean subtle grid borders
      dataRange.setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
    }
    
    if (!silent) {
      SpreadsheetApp.getUi().alert("Horizontal Report Generated", "The 'Checklist_Report_Horizontal' tab is updated with " + reportRows.length + " evaluation records!\n\nAll headers are human-readable, checkmarks are clean (✓ / ✗ / —), and evaluator names are populated.", SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (err) {
    Logger.log("Error in syncHorizontalReport: " + err.toString());
    if (!silent) {
      SpreadsheetApp.getUi().alert("Error Creating Horizontal Report", err.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}

/**
 * Helper to format raw Yes/No/N/A values into clean presentation symbols.
 */
function formatCheckmark(val) {
  if (val === "Yes") return "✓";
  if (val === "No") return "✗";
  if (val === "N/A") return "—";
  return val;
}


