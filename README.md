# Mail Merge
Create mass emails using GMail, Google Sheets and Google App Script, with variables and advanced scripting options

## Use case

## Setup
The easiest way to setup the system is by copying [this spreadsheet](https://docs.google.com/spreadsheets/d/1boihzJ4OGOHytMMEi9k6nDC4EFIOD96UzFYxVGtGuPc/copy). If you do not want to do that, you can also create an empty spreadsheet and click 'Extensions > App Script'. Now you can copy the contents of the file [Code.gs](Code.gs) into the script, and you're ready to go.

## Authorization
The first time you use the script in a new document you need to authorize the script. This gives it permission to do the things in your Google Account it needs to do. The first time you try to use any of the options, you'll get a popup, telling you to authorize it.

![Popup with the text 'Authorisation Required - A script attached to this document needs your permission to run' and the options 'Continue' and 'Cancel'](images/Authorization%20Required.png){width=250}

Click 'Continue' and a popup will appear telling you to pick your account. If you use a personal Google Account afterwards you might get the message that the app is not verified. To be able to use the app, click 'Advanced' and then click 'Go to Gmail/Sheets Mail Merge (unsafe)'

![Popup with the text 'Google hasn’t verified this app - The app is requesting access to sensitive info in your Google Account. Until the developer verifies this app with Google, you shouldn't use it.' with at the bottom the text 'Continue only if you understand the risks and trust the developer.' and a link saying 'Go to Gmail/Sheets Mail Merge (unsafe)'.](images/Verify%20app.png){width=250}

After this, you'll get a list of authorizations the app needs:

* Read, compose, send and permanently delete all your email from Gmail
* View and manage spreadsheets that this application has been installed in
* Send email as you
* Display and run third-party web content in prompts and sidebars inside Google applications

These are required for the app to work. To use the app, click 'Allow'. Now you're able to use the app. Be aware, the action you tried to perform before authorizing will not be performed, you need to do it again.
