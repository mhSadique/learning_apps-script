# I will use this repo to take notes on what I learned about Apps Script (G Suite)

## Code examples with descriptions below:

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
