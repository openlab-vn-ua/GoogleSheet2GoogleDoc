// Explode spreadsheet data to mutiple documents/spreadsheets
// Free and open sorce based on MIT license

// Help
// ------------------------------------

let APP_NAME = "GoogleSheet2GoogleDoc";
let VERSION  = "2.2";

function getHelpText() {
  let help = ""; let EOL = "\n"; let SQT = " ";
  help += APP_NAME + " " + "Ver:" + VERSION + EOL;
  help += "Explode spreadsheet data to multiple documents/spreadsheets." + SQT;
  help += "Uses first table on active sheet as a data source for filling multiple files by template(s)" + SQT;
  help += "First row on sheet should contain field names (until first empty column)." + SQT;
  help += "Each data row will produce own output file by template file." + SQT;
  help += "Each {{field}} in document will be replaced with data row cell content." + SQT;
  help += "Stops at end of sheet or first empty data row." + EOL;
  help += "" + EOL;
  help += "Template document url/id:" + EOL;
  help += "a). Extracted from note on first column" + EOL;
  help += "b). Extracted from special field !template inside a row, to make this row have alternate template" + EOL;
  help += "" + EOL;
  help += "Special columns:" + EOL;
  help += "a). !skip - if exist, this row is skipped if value contains TRUE, Y, y, true" + EOL;
  help += "b). !template - if exist, will specify template file url/id for this row (used if not empty)" + EOL;
  help += "c). !output - if exist, specify output file name (default will be {{templateName}}_filled_{{rowNum}}" + EOL;
  help += "d). !output.folder - if exist, will specify sub folder name to save files to ('.' = same as template file path)" + EOL;
  help += "e). !pdf.folder - if exist, will specify sub folder name to save .pdf files to ('.' = same as template file path)" + EOL;
//help += "f). !pdf.email - if exist, will specify email(s) to send file to (reserved)" + EOL;
  help += "" + EOL;
  help += "Special treatment of values:" + EOL;
  help += "a). multiline fields: in all lines started with '#).' # replaced by item number (starting form '1).'" + EOL;
  help += "" + EOL;
  help += "Other notes:" + EOL;
  help += "a). template can be document or spreadsheet (that is, generation of spreadsheets supported also!)" + EOL;
  help += "b). in spreadsheet, hidden sheets are skipped and formula cells are not changed" + EOL;
  help += "c). the system may generate .pdf files from output files, if !pdf.folder column exist (experimental)" + EOL;
  help += "d). the note of first field may contains additional options in form OptionName=value (each on separate line)" + EOL;
  help += "*** Supported options are: 'SavePDF=Y'" + EOL;
  return help;
}

// Utils
// ------------------------------------

function strTrim(str) {
  return str.toString().replace(/^[\s\uFEFF\xA0]+|[\s\uFEFF\xA0]+$/g, '');
}

function escapeRegExp(str) {
  if (str == null) { return str; }
  return str.toString().replace(/[^A-Za-z0-9_]/g, '\\$&');
}

function escapeReplacementString(str) {
  if (str == null) { return str; }
  return str.toString().replace(/\$/g, '$$$&');
}

function strReplaceAll(subject, search, replacement) {
  function escapeRegExp(str) { return str.toString().replace(/[^A-Za-z0-9_]/g, '\\$&'); }
  search = search instanceof RegExp ? search : new RegExp(escapeRegExp(search), 'g');
  return subject.replace(search, replacement);
}

if (String.prototype['replaceAll'] == null) { String.prototype['replaceAll'] = function(search, replacement) { return strReplaceAll(this, search, replacement) }; }

function escapeHtml(text) {
  text = text.replaceAll('&', escapeReplacementString('&amp;'));
  text = text.replaceAll('<', escapeReplacementString('&lt;'));
  text = text.replaceAll('>', escapeReplacementString('&gt;'));
  text = text.replaceAll('"', escapeReplacementString('&quot;'));
  return text;
}

function preformatHtmlNewlines(text) {
  text = text.replaceAll("\n", escapeReplacementString("<br/>\n"));
  return text;
}

function makeCssTag(text) {
  text = '<style>'+"\n"+escapeHtml(text)+"\n"+'</style>'+ "\n";
  return text;
}

// Helpers
// ------------------------------------

function isValueEmpty(str) {
  if (str == null) { return true; }
  if (str == '') { return true; }
  if (strTrim(str) == '') { return true; }
  return false;
}

function showError(msg) {
  let asp = SpreadsheetApp.getActiveSpreadsheet();
  if (asp != null) { asp.toast(msg, 'ERROR!'); }
}

function showInfo(msg) {
  let asp = SpreadsheetApp.getActiveSpreadsheet();
  if (asp != null) { asp.toast(msg); }
}

function getIdFromUrl(url) { 
  if (url == null) { return null; }
  let ids = url.toString().match(/[-\w]{25,}/);
  if ((ids != null) && (ids.length > 0)) { return ids[0].toString(); }
  return null;
}

function makeFromTextByFieldName(fieldName) {
  return '{{' + strTrim(fieldName) + '}}';
}
    
function makeToTextByFieldValue(fieldValue) {
  let textTo = fieldValue;
  
  if (textTo == null) { textTo = ""; }
  
  textTo = strTrim(textTo.toString());
  
  // mutilined text will have leading '#' replaced by numbers
  
  let partsTo = textTo.split('\n');
  if (partsTo.length > 1)
  {
    let NUMSEP = "). ";
    textTo = "";
    let pn = 1;
    for (let pi = 0; pi < partsTo.length; pi++) {
      let part = strTrim(partsTo[pi]);
      if (part.length > 0) {
        if (part[0] == '#') { part = (pn++).toString() + NUMSEP + part.substr(1); } // replace # with number
        textTo += part + '\n';
      }
    }
  }
  
  return strTrim(textTo);
}

function isSubFolderNameEmpty(subFolderName) {
    if ((subFolderName != null) && (subFolderName != "") && (subFolderName != ".")) {
      return false;
    } else {
      return true;
    }
}

function saveFileAsPDF(masterFile, outPDFSubFolderName) {
  let outFile = DriveApp.createFile(masterFile.getBlob().getAs(MimeType.PDF));
  outFile.setName(masterFile.getName()+".pdf");
  let masterDirs = masterFile.getParents();
  if (masterDirs.hasNext()) {
    let masterDir = masterDirs.next();
    outFile.moveTo(masterDir);

    if (!isSubFolderNameEmpty(outPDFSubFolderName)) {
      let theSubDirs = masterDir.getFoldersByName(outPDFSubFolderName);
      let theSubDir = null;
      if (theSubDirs.hasNext()) {
        theSubDir = theSubDirs.next();
      } else {
        theSubDir = masterDir.createFolder(outPDFSubFolderName);
      }

      if (theSubDir != null) {
        outFile.moveTo(theSubDir);
      }
    }
  }
}

function moveFileToSubFolder(masterFile, targetSubFolderName) {
  let masterDirs = masterFile.getParents();
  if (masterDirs.hasNext()) {
    let masterDir = masterDirs.next();

    if (!isSubFolderNameEmpty(targetSubFolderName)) {
      let theSubDirs = masterDir.getFoldersByName(targetSubFolderName);
      let theSubDir = null;
      if (theSubDirs.hasNext()) {
        theSubDir = theSubDirs.next();
      } else {
        theSubDir = masterDir.createFolder(targetSubFolderName);
      }

      if (theSubDir != null) {
        masterFile.moveTo(theSubDir);
      }
    }
  }
}

// Proc
// ------------------------------------

function procMerge(fieldNames, fieldValues, defaultTemplateId, mergeOptions) {
  let FAIL = -1;
  
  let result = 0;
  
  if (defaultTemplateId != null) { defaultTemplateId = getIdFromUrl(defaultTemplateId); }
  
  let fieldIndexTemplateId = null;
  let fieldIndexSkipFlag = null;
  let fieldIndexOutputName = null;
  let fieldIndexOutPDFPath = null;
  let fieldIndexOutResPath = null;
  
  let fieldNamesCount = 0;
  for (let f = 0; f < fieldNames.length; f++) {
    if (isValueEmpty(fieldNames[f])) {
      // empty cell
      break;
    } else {
      fieldNamesCount++;
      if (fieldNames[f] == '!template') {
        fieldIndexTemplateId = f;
      }
      if (fieldNames[f] == '!skip') {
        fieldIndexSkipFlag = f;
      }
      if (fieldNames[f] == '!output') {
        fieldIndexOutputName = f;
      }
      if (fieldNames[f] == '!pdf.folder') {
        fieldIndexOutPDFPath = f;
      }
      if (fieldNames[f] == '!output.folder') {
        fieldIndexOutResPath = f;
      }
    }
  }
  
  if ((defaultTemplateId == null) && (fieldIndexTemplateId == null))
  { 
    let EOL = '\n';
    showError
    (
      ""
      +"No templates defined."+EOL
      +"Put URL of template on note on first header cell"+EOL
      +"or add column '!template' with url of templates for individual rows"
    ); 
    return FAIL; 
  }
  
  if (fieldNamesCount <= 0) { showError("No field names found in first row"); return FAIL; }

  // defaults  
  let setupFromRowIndex = 1; // assuming row 0 is a header row
  let setupSavePDF = false;
  let setupLimit = null;

  if ((mergeOptions != null) && (mergeOptions.fromRowIndex != null)) { setupFromRowIndex = mergeOptions.fromRowIndex; }
  if ((mergeOptions != null) && (mergeOptions.savePDF != null)) { setupSavePDF = mergeOptions.savePDF; }
  if ((mergeOptions != null) && (mergeOptions.limitRowCount != null)) { setupLimit = mergeOptions.limitRowCount; }

  if (setupFromRowIndex >= fieldValues.length) { showInfo("No rows to proccess"); return 0; }
  if ((setupLimit != null) && (setupLimit <= 0)) { showInfo("No rows allowed to proccess"); return 0; }

  for (let i = setupFromRowIndex; i < fieldValues.length; i++) {

    if ((setupLimit != null) && (result >= setupLimit)) { return result; }
  
    let fieldRow = fieldValues[i];
    
    if (fieldRow == null) { return result; }
      
    let isEmptyRow = true;
    for (let f = 0; f < fieldNamesCount; f++) {
      if (isValueEmpty(fieldRow[f])) {
        // empty cell
      } else {
        isEmptyRow = false;
        break;
      }
    }
    
    if (isEmptyRow) { break; } // run until first empty row
    
    if (fieldIndexSkipFlag != null) {
      let sf = fieldRow[fieldIndexSkipFlag];
      if (isValueEmpty(sf)) {
        // ok, go on
      }
      else if ((sf === true) || (sf == 'Y') || (sf == 'y') || (sf == 'TRUE') || (sf == 'true') || (sf === true)) {
        continue; // skip line
      }
    }

    let rowTemplateId = defaultTemplateId;
    
    if (fieldIndexTemplateId != null) {
      if (!isValueEmpty(fieldRow[fieldIndexTemplateId])) {
        rowTemplateId = getIdFromUrl(fieldRow[fieldIndexTemplateId]);
      }
    }
    
    let templateFile = DriveApp.getFileById(getIdFromUrl(rowTemplateId));
  
    if (templateFile == null) { showError("Template file not found"); return FAIL; }
    
    let mergedName = templateFile.getName()+"_filled_"+i;
    
    if (fieldIndexOutputName != null) {
      if (!isValueEmpty(fieldRow[fieldIndexOutputName])) {
        mergedName = fieldRow[fieldIndexOutputName];
      }
    }

    let outPDFPath = null; // null, '', '.' = in same place where template is found

    if (fieldIndexOutPDFPath != null) {
      if (!isValueEmpty(fieldRow[fieldIndexOutPDFPath])) {
        outPDFPath = fieldRow[fieldIndexOutPDFPath];
      }
    }

    let outResPath = null; // null, '', '.' = in same place where template is found

    if (fieldIndexOutResPath != null) {
      if (!isValueEmpty(fieldRow[fieldIndexOutResPath])) {
        outResPath = fieldRow[fieldIndexOutResPath];
      }
    }

    let mergedFile = templateFile.makeCopy();
    mergedFile.setName(mergedName);

    if (mergedFile.getMimeType() == MimeType.GOOGLE_SHEETS) {

      let DO_FAST_UPDATE = true;

      let statsCellsUpdated = 0;

      let mergedDoc = SpreadsheetApp.openById(mergedFile.getId());
    
      let sheets = mergedDoc.getSheets();
      
      for (let sheet of sheets) {

        if (sheet.isSheetHidden()) { continue; } // skip hidden sheets

        let range = sheet.getDataRange();
        let values = range.getValues();
        let formulas = range.getFormulas(); // same dimention as values

        for (let irow in values) {
          for (let icol in values[irow]) {
            if ((formulas[irow] != null) && (!isValueEmpty(formulas[irow][icol])))
            { 
              continue; // skip formulas
            }

            let theText = values[irow][icol];

            if (typeof(theText).toString().toLowerCase() != "string")
            {
              continue; // skip non-string fiels
            }

            if (isValueEmpty(theText)) { continue; }

            theText = theText.toString();
            let newText = theText;

            for (let f = 0; f < fieldNamesCount; f++) {
              if (isValueEmpty(fieldNames[f])) { continue; }
              
              let textFrom = makeFromTextByFieldName(fieldNames[f]);
              let textTo = makeToTextByFieldValue(fieldRow[f]);
              
              newText = strReplaceAll(newText.toString(), textFrom, escapeReplacementString(textTo));
            }

            if (DO_FAST_UPDATE)
            {
              if (theText != newText)
              { 
                values[irow][icol] = newText;
                statsCellsUpdated++;
              }
            }
            else
            {
              if (theText != newText)
              { 
                let nrow = 1+Number(irow);
                let ncol = 1+Number(icol);
                range.getCell(nrow, ncol).setValue(newText);
                statsCellsUpdated++;
              }
            }
          }
        }

        if (DO_FAST_UPDATE)
        {
          // range.setValues(values); // breaks formulas

          // https://stackoverflow.com/questions/54775597
          var data = values.map(function(row, i) {
            return row.map(function(col, j) {
              return formulas[i][j] || col;
            });
          });

          range.setValues(data);
        }
      }
      
      SpreadsheetApp.flush();

      if (setupSavePDF) { saveFileAsPDF(mergedFile, outPDFPath); }
      if (!isSubFolderNameEmpty(outResPath)) { moveFileToSubFolder(mergedFile, outResPath); }
      
      //mergedDoc.close();

      console.log (
        "Filled spreadsheet:" + "{"+mergedFile.getName()+"/"+mergedFile.getId()+"}" 
      + " " + "cells updated:" + statsCellsUpdated
      );

      result++;
    }
    else // GOOGLE_DOCUMENT
    {
      let mergedDoc = DocumentApp.openById(mergedFile.getId()); 
    
      let body = mergedDoc.getBody();
      
      for (let f = 0; f < fieldNamesCount; f++) {
        if (isValueEmpty(fieldNames[f])) { continue; }
        
        let textFrom = makeFromTextByFieldName(fieldNames[f]);
        let textTo = makeToTextByFieldValue(fieldRow[f]);
        
        body.replaceText(escapeRegExp(textFrom), textTo); // Do not need to escape replacemet string here? // $ does not have spec meaning in replacement
      }
      
      mergedDoc.saveAndClose();

      if (setupSavePDF) { saveFileAsPDF(mergedFile, outPDFPath); }
      if (!isSubFolderNameEmpty(outResPath)) { moveFileToSubFolder(mergedFile, outResPath); }
      
      console.log (
        "Filled document:" + "{"+mergedFile.getName()+"/"+mergedFile.getId()+"}" 
      );

      result++;
    }
  }
  
  return result;
}

// Main
// ------------------------------------

function doMergeProc(fromRowIndex, limitRowCount) {
  let FAIL = -1;

  let sheet = SpreadsheetApp.getActiveSheet();//current sheet
  if (sheet == null) { return; }

  let data = sheet.getDataRange();
  let numRows = data.getNumRows();

  if (numRows <= 0) { 
    // nothing to do
    showError('Nothing to do: table in upper left corner not found');
    return FAIL; 
  }

  let mainTemplateId = null;

  let options = { };

  if (fromRowIndex != null) { options.fromRowIndex = fromRowIndex; }
  if (limitRowCount != null) { options.limitRowCount = limitRowCount; }

  let cmdLine = data.getNote();
  
  if (cmdLine != null) {
    let cmdLineParts = cmdLine.split("\n").map(function(s) { return strTrim(s); });

    cmdLineParts.map(function(s) { if (getIdFromUrl(s) != null) { mainTemplateId = s; }});
    cmdLineParts.map(function(s) { if (s.toUpperCase() == 'SAVEPDF=Y') { options.savePDF = true;}});
  }
  
  let values = data.getValues();
  let fieldNames = values[0];//First row of the sheet must be the the field names

  let count = procMerge(fieldNames, values, mainTemplateId, options);

  return count;
}

function doMerge() {
  showInfo('Starting replacing {{field_names}} with values from table in copies of template document');
  let count = doMergeProc();
  if ((count != null) && (count >= 0))
  { 
    showInfo('Completed '+count+' file(s) created'); 
  }
}

function doShowHelp() {
  let html = preformatHtmlNewlines(escapeHtml(getHelpText()));
  html = makeCssTag('body { font-size: small; font-family: Tahoma; }') + html;
  let hout = HtmlService.createHtmlOutput(html).setTitle(APP_NAME + ' ' + 'How to use');

  SpreadsheetApp.getUi().showSidebar(hout); // Or DocumentApp or SlidesApp or FormApp.
}

// Long runner

/*
function doMergeWorker(stepIndex) {
  let REPORT_INTERVAL = 5;
  
  if (stepIndex == 0) {
    showInfo('Starting replacing {{field_names}} with values from table in copies of template document (process begin)');
  }

  let count = doMergeProc(stepIndex + 1, 1);

  if ((count != null) && (count > 0))
  { 
    if (((stepIndex + 1) % REPORT_INTERVAL) == 0)
    {
      showInfo('Proccessed '+(stepIndex+1)+' rows'); 
    }

    return true; // ok, go on
  }
  else if ((count != null) && (count == 0))
  {
    showInfo('Completed '+stepIndex+' rows(s) proccessed'); 
    return false;
  }
  else
  {
    return false; // error
  }
}

//let mainRunner = new Runner(function(i) { console.log('RunnerStep:DEBUG:'+i); return i < 3; });
//let mainRunner = new Runner(doMergeWorker);

function doMergeByRunner() {
  mainRunner.run();
}
*/

// Integration

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the doMerge() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (spreadsheet == null) { return; }
  let entries = [
    {
      name : "Fill docs with {{template}} fields",
      functionName : "doMerge"
    },
    {
      name : "Show help on how to use",
      functionName : "doShowHelp"
    }
  ];
  spreadsheet.addMenu(APP_NAME, entries);
};
