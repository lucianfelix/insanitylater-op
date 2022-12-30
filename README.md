# insanitylater-op

Get an API key from [OpenAPI](https://beta.openai.com/account/api-keys).
Copy `src/taskpane/components/apiKey.sample.js` to `src/taskpane/components/apiKey.js` and put the key inside. 

To run it:
```
npm run build:dev
npm run dev-server
```

Then go to Outlook, Get Addins (in the main toolbar, click ... if not visible).
Select My Addins, Custom Addins (scroll to the end), Add, From file.., select `manifest.xml` in the project.

New Message, customise toolbar, add 'Serenity Now' button.


To use, write an draft email, then press 'Serenity Now' to show the Taks Pane..
Chose a Persona and a Tone then press 'Insanity Later'.
