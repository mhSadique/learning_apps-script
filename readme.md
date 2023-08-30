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

- Below is how you can select cells using notations and/or cell values:

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
  firstSheet.getRange(4, 8, 6).setBackground("yellow"); // select 6 cells vertically from the intersection of 4th row and 8th column
  firstSheet.getRange(4, 8, 6, 3).setBackground("yellow"); // select 6 cells vertically and 3 columns horizontally from the intersection of 4th row and 8th column
  console.log(ss.getName());
}
```
