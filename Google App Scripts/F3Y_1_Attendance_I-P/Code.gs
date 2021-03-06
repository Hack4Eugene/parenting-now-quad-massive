/*
* Author: Josiah Young - digitaldavinci1618
* Created: 4-6-2018
* Last Modified: 4-17-2018
* Last Modified By: Josiah Young - digitaldavinci1618
*
* Summary: 
* This script prompts the user for a specifically formatted Google Drive Spreadsheet used 
* by Parenting Now and populates a Google Doc "form" template with report data.
* It integrates with and runs on Google Drive App Script Host. 
*
* @License: https://spdx.org/licenses/MIT.html
* 
* Copyright 2018 Parenting Now https://parentingnow.org/
* 
* Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated
* documentation files (the "Software"), to deal in the Software without restriction, including without limitation
* the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, 
* and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
* 
* The above copyright notice and this permission notice shall be included in all copies or substantial 
* portions of the Software.
* 
* THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED
* TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL 
* THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF 
* CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS 
* IN THE SOFTWARE.
*/

function onOpen() {
  DocumentApp.getUi()
      .createMenu('Parenting Now')
      .addItem('Populate from report', 'selectReport')
      .addToUi();
}

function selectReport() {
  
    // get file from dialog
    var html = HtmlService.createHtmlOutputFromFile('Picker.html')
        .setWidth(840)
        .setHeight(665)
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    DocumentApp.getUi().showModalDialog(html, 'Select Spreadsheet');
}

function getOAuthToken() {
    DriveApp.getRootFolder();
    return ScriptApp.getOAuthToken();
}

function populateFromReport(fileId) {
  
  var ui = DocumentApp.getUi();
  
  // get fields from report - debug alert display
  var report = SpreadsheetApp.openById(fileId).getActiveSheet().getDataRange().getValues();
  
  var reportValues = [];
  var reportKeys = [];
  
  // iterate through spreadsheet report and build array of hashes of people data
  for (var r = 0; r < report.length; r++) {
    if (r != 0) { // skip header row for value collection
        reportValues[r] = {};
    }
    
    for (var c = 0; c < report[r].length; c++) {
      if (r == 0) { 
        // capture column headers 
        reportKeys.push(report[r][c]);
      } else {
        // capture people data
        reportValues[r][reportKeys[c]] = report[r][c];
      }
    }
  }
 
  // remove first array index because we skipped it since it was header data and its undefined (blank)
  reportValues.shift();
  
  DocumentApp.getUi().alert(reportValues.length);
  
  // now start working with the form that needs filling in
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  
  // set up global values
  if (reportValues.length) {
      body.replaceText("<<FGroupID>>", reportValues[0]["fGroupID"]);
      body.replaceText("<<fFacilitatorFullName>>", reportValues[0]["fFacilitatorFullName"]);
  }

  var table = body.getTables()[0]; // first table in the document is the form we want to fill in

  // add some rows if we are out of rows
  var neededRows = reportValues.length - 12 < 0 ? 0 : reportValues.length - 12; 
  var insertRow = 39; // specific to this document format
  var insertRowNumber = 13; // row number label
  var templateRowUpperG = table.getRow(3).copy();  // Gray - this is the first tablerow that is for text entry (past the header tablerow)
  var templateRowMiddleG = table.getRow(4).copy(); // Gray - this is the lower tablerow that is for text entry 
  var templateRowLowerG = table.getRow(5).copy(); // Gray - this is the lower tablerow that is for text entry 
  templateRowUpperG.getCell(1).setText( "" ); 
  templateRowUpperG.getCell(3).setText( "" );
  templateRowMiddleG.getCell(1).setText( "" );  // this is <<ClFirst>> and <<ClLast>> & <<PFirst>> <<PLast>>
  templateRowLowerG.getCell(1).setText( "" );  // this is <<fChiFirName>>
    
  var templateRowUpperW = table.getRow(6).copy();  // White - this is the first tablerow that is for text entry (past the header tablerow)
  var templateRowMiddleW = table.getRow(7).copy();  // White - this is the first tablerow that is for text entry (past the header tablerow)
  var templateRowLowerW = table.getRow(8).copy(); // White - this is the lower tablerow that is for text entry
  templateRowUpperW.getCell(1).setText( "" ); 
  templateRowUpperW.getCell(3).setText( "" );
  templateRowMiddleW.getCell(1).setText( "" );  // this is <<ClFirst>> and <<ClLast>> & <<PFirst>> <<PLast>>
  templateRowLowerW.getCell(1).setText( "" );  // this is <<fChiFirName>
  
  var style = {};
  style[DocumentApp.Attribute.BOLD] = false;

  for (var newRow = 0; newRow <= neededRows; newRow += 6) {
  
    // Upper White TableRow
    templateRowUpperG.getCell(0).setText(insertRowNumber++).setAttributes(style);
    table.insertTableRow(insertRow, templateRowUpperG.copy());
    // Middle White TableRow
    table.insertTableRow(insertRow + 1, templateRowMiddleG.copy())
    // Lower White TableRow
    table.insertTableRow(insertRow + 2, templateRowLowerG.copy());
    
    
    // Upper Gray TableRow
    templateRowUpperW.getCell(0).setText(insertRowNumber++).setAttributes(style);
    table.insertTableRow(insertRow + 3, templateRowUpperW.copy());
    // Middle Gray TableRow
    table.insertTableRow(insertRow + 4, templateRowMiddleW.copy());
    // Lower Gray TableRow
    table.insertTableRow(insertRow + 5, templateRowLowerW.copy());
    
    insertRow += 6;
  }
  
  // loop from tablerow "1" (actually tableRow(6) because of the table header) until cell 0 has an equals sign in it - "Ave parent attendance =".
  // these are the tablerows that we want to fill in with data.'
  // these "rows" are actually made up of two tablerows each
  var reportIndex = 0;
  for (var i = 3; i < reportValues.length + 6; i += 3) {
      if (reportIndex < reportValues.length) {
        var fAdultPhoneCell = String(reportValues[reportIndex]["fAdultPhoneCell"]);
      if (fAdultPhoneCell) {
        if (fAdultPhoneCell.length == 7) {
          fAdultPhoneCell = fAdultPhoneCell.substr(0,3) + "-\n" + fAdultPhoneCell.substr(3);
        }
        fAdultPhoneCell = reportValues[reportIndex]["fAdultPhoneAreaCodeCel"] + "-" + fAdultPhoneCell;
      } else {
        fAdultPhoneCell = "";
      }
    
      table.getRow(i).getCell(1).setText( reportValues[reportIndex]["ClFirst"] + " " + reportValues[reportIndex]["ClLast"] );
      table.getRow(i).getCell(3).setText( fAdultPhoneCell );  // this is <<fAdultPhoneCell>>
      table.getRow(i+1).getCell(1).setText( reportValues[reportIndex]["PFirst"] + " " + reportValues[reportIndex]["PLast"] );  // this is <<ClFirst>> and <<ClLast>> & <<PFirst>> <<PLast>>
      table.getRow(i+2).getCell(1).setText( reportValues[reportIndex]["fChiFirName"] );  // this is <<fChiFirName>>
   
      reportIndex++;
    } else {
      break;
    }
  }
}
