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
