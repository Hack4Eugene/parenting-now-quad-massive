/*
* Author: Josiah Young - digitaldavinci1618
* Created: 4-6-2018
* Last Modified: 4-8-2018
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
 
  
  // now start working with the form that needs filling in
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  
  // set up global values
  if (reportValues.length) {
      body.replaceText("<<FGroupID>>", reportValues[0]["fGroupID"]);
      body.replaceText("<<fFacilitatorFullName>>", reportValues[0]["fFacilitatorFullName"]);
  }
  
  var table = body.getTables()[0]; // first table in the document is the form we want to fill in 
  var templateRowUpper = body.getTables()[0].getRow(3);  // this is the first tablerow that is for text entry (past the header tablerow)
  
  var reportIndex = 0;
  // loop from tablerow "1" (actually tableRow(3) because of the table header) until cell 0 has an equals sign in it - "Ave parent attendance =".
  // these are the tablerows that we want to fill in with data.
  // these "rows" are actually made up of two tablerows each.
  for (var i = 3; i < table.getNumRows(); i++) {
    if (table.getRow(i).getCell(0).findText("=") || reportIndex >= reportValues.length) { // end of fillable rows in form or we are out of data to fill in
        break;
    } else if (table.getRow(i).getCell(0).findText("[0-9]+")) {
        // upper table row - odd
        table.getRow(i).getCell(1).setText( reportValues[reportIndex]["ClFirst"] + " " + reportValues[reportIndex]["ClLast"] );  // this is ClFirst and ClLast
        table.getRow(i).getCell(3).setText( reportValues[reportIndex]["fAdultPhoneAreaCodeCel"] + "-" + reportValues[reportIndex]["fAdultPhoneCell"] );  // this is <<fAdultPhoneAreaCodeCel>>-<<fAdultPhoneCell>>
        
    } else {
        // lower table row - even
        table.getRow(i).getCell(1).setText( reportValues[reportIndex]["PFirst"] + " " + reportValues[reportIndex]["PLast"] );  // this is <<PFirst>> <<PLast>>
        reportIndex++;
    }
  }
  // Notes for later development after the challenge...
  // ui.alert(templateRow.getCell(0).editAsText().getText());  // this is the row number
  //table.appendTableRow(body.findText("<<ClFirst>> <<ClLast>>").getElement().getParent().getParent().getParent().asTableRow().copy());
}