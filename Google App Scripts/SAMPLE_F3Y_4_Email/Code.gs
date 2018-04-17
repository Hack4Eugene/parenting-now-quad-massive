/*
* Author: Josiah Young - digitaldavinci1618
* Created: 4-6-2018
* Last Modified: 4-17-2018
* Last Modified By: Josiah Young - digitaldavinci1618
*
* Summary: 
* This script prompts the user for a specifically formatted Google Drive Spreadsheet used 
* by Parenting Now and populates a Google Spreadsheet Email template with report data.
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
  SpreadsheetApp.getUi()
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
    SpreadsheetApp.getUi().showModalDialog(html, 'Select Spreadsheet');
}

function getOAuthToken() {
    DriveApp.getRootFolder();
    return ScriptApp.getOAuthToken();
}

function populateFromReport(fileId) {
  
  var ui = SpreadsheetApp.getUi();
  
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
  var emailSheet = SpreadsheetApp.getActiveSheet();
  
  // delete existing rows 
  emailSheet.deleteRows(2, 3);
 
  // enter values into email sheet
  for (var i = 0; i < reportValues.length; i++) {
    var names = reportValues[i]["ClFirst"] + " " + reportValues[i]["ClLast"];
    if (reportValues[i]["PFirst"]) {
        names += " & " + reportValues[i]["PFirst"] + " " + reportValues[i]["PLast"];
    }
    emailSheet.appendRow([
        names,
        reportValues[i]["ClEmail"],
        reportValues[i]["Pemail"]
    ]);
  }
  
  // format sheet
  emailSheet.setRowHeights(14, reportValues.length + 4, 34);
  var range = emailSheet.getRange("A2:C2");
  
  range.copyFormatToRange(emailSheet, 1, 3, 10, reportValues.length + 4);
  emailSheet.getDataRange().setFontSize(10);
}
