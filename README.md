# pdf-generator

## Disclaimer
It’s not possible to share the files and info's that this script was using it, and I scrambled the variable names and other stuff to make it anonymous, might made some mistakes in this step.

## Structure
Using os package to acess the files all at once, so the query update is faster, and then acessing them again with 'win32com.client' tools one by one to generate the PDF's. Using these two tools and pyautogui for the last part: acessing another spreadsheet that would update its queries upon entry, and then copying this data to another place and printing the screen in a certain region.

## Approach
I created this to open a few spreadsheets that would update its queries upon opened, and then generating PDF files from each spreadsheet. Then I would acess another spreadsheet, update its queries, copy the data to a third spreadsheet, paste it there and take a print screen of the result.
