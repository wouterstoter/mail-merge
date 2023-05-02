# Mail Merge
Create mass emails using GMail, Google Sheets and Google App Script, with variables and advanced scripting options

## Use case

## Setup
The easiest way to setup the system is by copying [this spreadsheet](https://docs.google.com/spreadsheets/d/1boihzJ4OGOHytMMEi9k6nDC4EFIOD96UzFYxVGtGuPc/copy). If you do not want to do that, you can also create an empty spreadsheet or open an existing one and click 'Extensions > App Script'. Now you can copy the contents of the file [Code.gs](Code.gs) into the script, and you're ready to go.

## Authorization
The first time you use the script in a new document you need to authorize the script. This gives it permission to do the things in your Google Account it needs to do. The first time you try to use any of the options, you'll get a popup, telling you to authorize it.

![Popup with the text 'Authorisation Required - A script attached to this document needs your permission to run' and the options 'Continue' and 'Cancel'](images/Authorization%20Required.png)

Click 'Continue' and a popup will appear telling you to pick your account. If you use a personal Google Account afterwards you might get the message that the app is not verified. To be able to use the app, click 'Advanced' and then click 'Go to Gmail/Sheets Mail Merge (unsafe)'

![Popup with the text 'Google hasn’t verified this app - The app is requesting access to sensitive info in your Google Account. Until the developer verifies this app with Google, you shouldn't use it.' with at the bottom the text 'Continue only if you understand the risks and trust the developer.' and a link saying 'Go to Gmail/Sheets Mail Merge (unsafe)'.](images/Verify%20app.png)

After this, you'll get a list of authorizations the app needs:

* Read, compose, send and permanently delete all your email from Gmail
* View and manage spreadsheets that this application has been installed in
* Send email as you
* Display and run third-party web content in prompts and sidebars inside Google applications

These are required for the app to work. To use the app, click 'Allow'. Now you're able to use the app. Be aware, the action you tried to perform before authorizing will not be performed, you need to do it again.

## Create Spreadsheet
The first sheet in your spreadsheet will be used as the data for your Mail Merge. Every row in this spreadsheet is a seperate email and every column is a variable. The first row of the sheet will be used as variable names. The variable names need to follow [JS Variable Naming Rules](https://developer.mozilla.org/en-US/docs/Learn/JavaScript/First_steps/Variables#an_aside_on_variable_naming_rules), with means you cannot use spaces in your variable name. Keep in mind that variable names are case sensitive.

Certain variables names have special properties, since they won't only be used in your email, but also will be used to set for example the receipients of your email. The following variables are special variables:
* **recipient** The addresses of the recipients, separated by commas
* **cc** A comma-separated list of email addresses to CC
* **bcc** A comma-separated list of email addresses to BCC
* **name** The name of the sender of the email (default: the user's name)
* **noReply** TRUE if the email should be sent from a generic no-reply email address to discourage recipients from responding to emails; this option is only possible for Google Workspace accounts, not Gmail users
* **replyTo** An email address to use as the default reply-to address (default: the user's email address)

At least one of the variables *recipient*, *cc* or *bcc* needs to be defined for the mail merge to work.

An example of a spreadsheet could be like the one below:
| recipient         | name | favoriteColor | birthday   |
|-------------------|------|---------------|------------|
| joe@example.com   | Joe  | Purple        | 2000-01-01 |
| anna@example.com  | Anna | Yellow        | 1990-12-31 |

## Write email
To write the email you can simply create an email draft in GMail. To insert variables we use [Google App Script Scriptlets](https://developers.google.com/apps-script/guides/html/templates). For a variable print, you can use the syntax `<?= variableName ?>`. These scriptlets can be used in your email body and in the subject line. In your email you can use any text formatting GMail offers, however, the full scriptlet needs to fully have the same formatting, otherwise you might get errors.
```
Dear <?= name ?>,

Hereby I'd like to inform you that your favorite colour is <?= Purple ?>.

Kind regards,
Me
```

### Text transformation
You can also use Google App Script functions (which mostly correspond to JavaScript functions) to transform your text. All variables are saved as strings, the way they are displayed in the spreadsheet. If your data has a different data type, for example a number, date or boolean, you can access the raw variable by putting an underscore (`_`) at the beginning of the variable name.
```
Dear <?= name.toUpperCase() ?>,

Your name starts with the letter <?= name.slice(0,1) ?>, and you were born in <?= _birthday.getFullYear() ?>.

Kind regards,
Me
```
