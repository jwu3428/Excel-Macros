# Excel Macros
Various Excel macros and macro-enabled workbooks built for a client.

## FileCopier.xlsm
Copies files from a source folder to a destination folder based on the queries in *Queries* sheet.
The matches will be partial (Example: *Hello* will find *HelloWorld.txt*).
1. Select a source folder by clicking *Browse Source...* The macro will parse all the files in the source folders.
2. Select a destination folder.
3. File Search Limit: Warns when the number of files found from a given query exceeds this number.
4. Search Subfolders: Enter "Yes" to search within subfolders or "No" to ignore subfolders
5. In the *Queries* sheet, enter a query in each cell. Click **CLEAR ALL** to clear all queries. 
6. Click **COPY FILES NOW** in the *Setup* sheet to find and copy files from the source folder to the destination.

After the operation is complete, a dialog box will display how many files were copied.
The *Queries* sheet will now display successful copies in green and failures in red.
Subsequent searches will skip queries that were already searched. Edit their cells to re-enable search for that query or enter new queries.

## Search.xlsm
Searches all open Excel workbooks and output results in the worksheet.

Simply type the query into the search box and press Enter to display the results. Certain cells of the result will be hyperlinked to its location.
Clicking on them will take you to its location.
