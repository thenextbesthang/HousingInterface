This following project built in Excel VBA was designed to allow for the easy processing of housing clients at my company. It would take client information either entered into a spreadsheet or a database, 
and on the basis of a housing client who comes in to my worksite, would look up information currently in "the system" of client service. Then, it would put this information into a a reportable form, and
then send a command to the local printer in order to print out a hard copy of the form. Files include:

- ClientEntryForm.frm - the code behind a custom built form to allow for user input of data
- ClientEntryForm.frx - the form itself for entering in client data
- getUserID.bas - links to a local database of client information, and queries the database for information pertaining to the client
- MouseScroll.bas - I added to a file found on the internet. Designed to allow user to be able to scroll through a list of 130 different deliverables with a mouse. Basic 
                    does not have, as far as I am aware, mouse click or mouse movement events
- PrintIt.frm - code for determining which rows in a spreadsheet ought to be printed. Client data was put into a spreadsheet; needed some way to say "only these 10 rows need to be printed" rather than 
                 600 to 700 pieces of information at a time
- PrintIt.frx - the form itself for printing the client information
- Sheet1.cls - pulls up the printing form
- Sheet3.cls - pulls up the client entry form
- updateActionRows.bas - outputs the specified data, through inserting it at specific bookmarked locations on a word document; and updates the housing database for the new client data entered; and prints 
                         out the word document to a local printer to have a record of serving a housing client.

