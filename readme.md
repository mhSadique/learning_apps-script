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
