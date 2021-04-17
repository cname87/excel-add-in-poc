# Excel Add-In Proof Of Concept

## Notes

1. You need a web server accessible to the users. I use <https://localhost:4200> for local development and test but you need a public site for general users.
   1. The server url is set in the manifest file.
   2. I have set up an Angular-based server.

2. The users would need to use a modern version of Excel.
   1. I use the Office 365 web version (which I have access to via work) as it is easy to sideload the manifest.xml file and it has full functionality.

3. The users need access to a manifest.xml file that is loaded in their local Excel instance.
   1. There are different ways of distributing this.
   2. You 'sideload' for development and test (and limited production) - you can sideload the file into Excel on the Web as follows:
      1. Open Insert -> Add-ins -> Manage My Add-ins -> Upload my Add-in, and then provide a link to the manifest file.
      2. If you change the manifest file then click 'Refresh' and then 'Upload my Add-in', and then provide a link to the manifest file.

4. There a few key available functionalities:

   1. Show a taskpane which is your web site in a pane. You can run commands via UI (e.g. buttons) that exercise the Excel API.
      1. I use <https://localhost:4200/index.html> for the taskpane url. An Angular component can show html and also run commands via buttons and component typescript.

   2. Embed the web server as content in a page - similar to the taskpane but appears as fixed content on the worksheet. I have not implemented this and prefer to use a taskpane.

   3. Add commands via a menu with items that run UI-less commands that exercise the Excel API.  A menu command can also open the taskpane.
      1. The UI-less commands must link to functions declared as global variables in the command url (which is the same url).  I list these as Window object properties in main.ts imported from a separate functions.ts file.

   4. Create custom functions (only on Office on the Web). I have not exercised this yet.

5. Notes:
   1. I have set up the manifest file so a shared runtime is used i.e. the taskpane and command menu run from the same JS instance in the browser.  Without this, even if you pointed each at the same index.html file, they would run as two separate instances.
