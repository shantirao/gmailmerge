This is a Google Sheets script to add a mailmerge user interface

1. Attach a script to your spreadsheet (Extensions -> Apps Script) 
2. Paste in the content of email.gs
3. Add a new file, type HTML, called *mailmerge*
4. Paste the contet of mailmerge.html
5. Add an execution trigger (the clock thingy on the left) and attach onOpen to when the spreadsheet opens
6. Run the onOpen function and grant the permissions
7. Run onOpen again

Then close the script, and you should be ready. 

1. From the new Email menu, do *Setup Mailmerge*
2. Choose the column to use as a filter (TRUE/FALSE or 1/0)
3. Choose the "to" column 
4. To see a preview, click < or > and look at the next row
5. Subject may contain ${field} as well as the message body
6. To include images, read the source code
