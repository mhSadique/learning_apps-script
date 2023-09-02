# I will use this repo to take notes on what I learned about Apps Script (G Suite)

## Code examples for `DocumentApp` with descriptions below:

- The code below create Google Doc, add some text content to it, and sends the link to the newly created file to a specified email address

```js
function myFunction() {
  const doc = DocumentApp.create("Business blueprint");
  const body = doc.getBody();
  body.appendParagraph("Hello, Sadique");
  // const email = Session.getActiveUser().getEmail();
  const email = "habibullah1307@gmail.com";
  const subject = doc.getName();
  const emailBody = `This is the new doc ${doc.getUrl()}`;
  console.log(email);
  GmailApp.sendEmail(email, subject, emailBody);
}
```

- How you put something in the console:

```js
function practice() {
  for (let i = 0; i < 10; i++) {
    Logger.log(`Logged ${i + 1} times`); // this is equivalent to console.log(), but runs on the backend
  }
}
```

- Use `getActiveDocument()` method to find information about the doc using script that are `container-bound`:

```js
function myFunction() {
  const doc = DocumentApp.getActiveDocument();
  Logger.log(doc.getId());
  Logger.log(doc.getName());
  Logger.log(doc.getUrl());
  Logger.log(doc.getBody().getParagraphs()[0].getText());
}
```

- But use `openById(id)` or `openByUrl(url)` method to find information about the doc when you are using `standalone` script:

```js
function myFunction() {
  const doc = DocumentApp.openById(
    "1__Ffz8mVRd1JVHskB9MlVTXa4ygQ18rGOHQSem7kby8"
  );
  Logger.log(doc.getUrl());
  Logger.log(doc.getName());
}
```

## Code examples for `SpreadsheetApp` with descriptions below:

- Simply create spreadsheet with the code below:

```js
function myFunction() {
  const ss = SpreadsheetApp.create("My monthly expense tracker");
}
```

- You can define the number of rows and columns when you create one:

```js
function myFunction() {
  const ss = SpreadsheetApp.create("My monthly expense tracker, 50, 20");
}
```

- How you select cells using notations and/or cell values:

```js
function fun1() {
  const ssId = "1A8mPDyaw2RRSaLlf1q_l1HTWEI8lsSWmEw7p8U1e2Kc";
  const ss = SpreadsheetApp.openById(ssId);
  const firstSheet = ss.getSheets()[0];

  // Notations
  firstSheet.getRange("A1").setBackground("red"); // select only one cell
  firstSheet.getRange("A3:A8").setBackground("green"); // select the cells in between (vertically)
  firstSheet.getRange("A3:F3").setBackground("tomato"); // select the cells in between (horizontally)
  firstSheet.getRange("G3:I13").setBackground("yellow"); // select a rectangle

  firstSheet.getRange("2:2").setBackground("green"); // select only 2nd row
  firstSheet.getRange("2:9").setBackground("yellow"); // select 2nd row to 9th row

  // Cell values
  firstSheet.getRange(4, 8, 6).setBackground("yellow"); // (row, column, numRows) // select 6 cells vertically from the intersection of 4th row and 8th column
  firstSheet.getRange(4, 8, 6, 3).setBackground("yellow"); // (row, column, numRows, numCols) // select 6 cells vertically and 3 columns horizontally from the intersection of 4th row and 8th column
  console.log(ss.getName());
}
```

- How you select ranges, get and set their values:

```js
function fun() {
  const ssId = "1A8mPDyaw2RRSaLlf1q_l1HTWEI8lsSWmEw7p8U1e2Kc";
  const ss = SpreadsheetApp.openById(ssId);
  const firstSheet = ss.getSheets()[0];
  const range = firstSheet.getRange(1, 1, 3, 4); // (row, col, numRows, numCols)
  let values = range.getValues();

  // we swap the values of 2nd and 3rd rows
  // REMEMBER, the two dimensional array must match the range you define
  const newValues = [values[0], values[2], values[1]];
  range.setValues(newValues);
  values = range.getValues();
  console.log(values);
}
```

- How you create a table in a Document using SpreadSheet data:

```js
function createTableInDocUsingSheetData() {
  const ssId = "1A8mPDyaw2RRSaLlf1q_l1HTWEI8lsSWmEw7p8U1e2Kc";
  const ss = SpreadsheetApp.openById(ssId);
  const firstSheet = ss.getSheets()[0];

  // get values of the cells, which is a multidimensional array
  const dataToPutInTheTable = firstSheet.getRange(1, 1, 3, 4).getValues();
  console.log("dataToPutInTheTable", dataToPutInTheTable);

  // create a doc where we will create a table
  const doc = DocumentApp.create("Sample Data");

  // select the body of the doc
  const body = doc.getBody();

  // create a paragraph with text as the name of the SpreadSheet and set it as heading1
  body
    .insertParagraph(0, ss.getName())
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  // put the data into the table and format it
  const table = body.appendTable(dataToPutInTheTable);
  table.getRow(0).editAsText().setBold(true);
}
```

- How you create table in a Document with dynamic SpreadSheet data:

```js
function createTableInDocUsingSheetDataThatAreDynamic() {
  const ssId = "1A8mPDyaw2RRSaLlf1q_l1HTWEI8lsSWmEw7p8U1e2Kc";
  const ss = SpreadsheetApp.openById(ssId);
  const firstSheet = ss.getSheets()[0];

  // get values of the cells, which is a multidimensional array
  // when you use getLastRow() and getLastColumn(), remember the adjust the row and col values in the getRange(row, col, numRows, numCols) method
  const dataToPutInTheTable = firstSheet
    .getRange(1, 1, firstSheet.getLastRow(), firstSheet.getLastColumn())
    .getValues();
  console.log("dataToPutInTheTable", dataToPutInTheTable);

  const docId = "1r_s3Cr2URp3o-v5Ec5OW3DU4DK_nPeGAEOwZxmQatW4";
  const doc = DocumentApp.openById(docId);

  // select the body of the doc
  const body = doc.getBody();

  // create a paragraph with text as the name of the SpreadSheet and set it as heading1
  body
    .insertParagraph(0, ss.getName())
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  // put the data into the table and format it
  const table = body.appendTable(dataToPutInTheTable);
  table.getRow(0).editAsText().setBold(true);
}
```

- Example tracker that reads data from a sheet and creates a table on a Doc and then put the doc's info in another sheet in the same SpreadSheet

```js
function trackSheet() {
  const ss = SpreadsheetApp.openById(
    "1A8mPDyaw2RRSaLlf1q_l1HTWEI8lsSWmEw7p8U1e2Kc"
  );

  const doc = DocumentApp.openById(
    "1bSVKI9BPDQGVdWVqMwaBqZ61C8ghaEJoMWPDKS7JBD8"
  );
  const docBody = doc.getBody();

  const sheet1 = ss.getSheetByName("Sheet1");
  const tracking = ss.getSheetByName("Tracking");

  const rowData = sheet1
    .getRange(1, 1, sheet1.getLastRow(), sheet1.getLastColumn())
    .getValues();
  console.log(rowData);

  docBody
    .appendParagraph("New Table #" + tracking.getLastRow())
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  const table = docBody.appendTable(rowData);
  console.log("table", table);
  const adder = tracking.appendRow([
    doc.getName(),
    doc.getId(),
    doc.getUrl(),
    Date(),
  ]);
  console.log("adder", adder);
}
```

- How to autoresize columns or rows in a sheet

```js
function autoResize() {
  const ss = SpreadsheetApp.openById(
    "1A8mPDyaw2RRSaLlf1q_l1HTWEI8lsSWmEw7p8U1e2Kc"
  );
  const trackingSheet = ss.getSheetByName("Tracking");
  // trackingSheet.autoResizeColumn(1); // resize only one specific column
  trackingSheet.autoResizeColumns(1, 4); // resize a range of columns
}
```

- How to clear the content of the sheet

```js
function clearSheet() {
  const ss = SpreadsheetApp.openById(
    "1A8mPDyaw2RRSaLlf1q_l1HTWEI8lsSWmEw7p8U1e2Kc"
  );
  const sheetToDelete = ss.getSheetByName("Sheet3");
  sheetToDelete.clear();
}
```

- How to clear the format a sheet

```js
function clearFormat() {
  const ss = SpreadsheetApp.openById(
    "1A8mPDyaw2RRSaLlf1q_l1HTWEI8lsSWmEw7p8U1e2Kc"
  );
  const sheetToClearFormat = ss.getSheetByName("Sheet3");

  // sheetToClearFormat.clearFormats();
  // the line below is equivalent to the one above
  sheetToClearFormat.clear({ formatOnly: true, contentsOnly: false });
}
```

- How to copy a sheet's data to another sheet in a different Spreadsheet:

```js
function copySheetDataToAnotherSheet() {
  const ss = SpreadsheetApp.openById(
    "1A8mPDyaw2RRSaLlf1q_l1HTWEI8lsSWmEw7p8U1e2Kc"
  );
  const sheetToCopyFrom = ss.getSheetByName("Sheet1");

  const ssToCopyTo = SpreadsheetApp.openById(
    "1hZazGMYHBSGLR1HfV8tpc9gvXhX8arLPYlfHuKoM7O8"
  );
  sheetToCopyFrom.copyTo(ssToCopyTo);
}
```

- How to `deleteColumn(4), deleteColumns(2, 1), getLastColumn(), getMaxColumns(), getName(), getParent().getName(), hideSheet(), showSheet(), insertColumns(2)` in a sheet in Spreadsheet

```js
function doOtherStuff() {
  const ss = SpreadsheetApp.openById(
    "1A8mPDyaw2RRSaLlf1q_l1HTWEI8lsSWmEw7p8U1e2Kc"
  );
  const sheet = ss.getSheetByName("Sheet1");

  // sheet.deleteColumn(4); // delete one column
  // sheet.deleteColumns(2, 1); // delete the columns in between

  const lastColumn = sheet.getLastColumn(); // get the last column that is populated with data
  const maxColumn = sheet.getMaxColumns(); // get maximum number of columns added to the sheet
  const sheetName = sheet.getName();
  const sheetParent = sheet.getParent().getName();

  // sheet.hideSheet(); // hide the sheet
  // sheet.showSheet(); // show the sheet

  // sheet.insertColumns(2); // there are other methods for inserting columns for inserting before and after - check them out

  console.log({ lastColumn, maxColumn, sheetName, sheetParent });
}
```
