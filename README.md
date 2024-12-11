# pdf-generator

## Disclaimer
It’s not possible to share the files that this script was using it, every spreadsheet and pdf is just an example for how the code works.

## Structure
Using os package to acess the files all at once (queries would update upon opened), so the query update is faster, and then acessing them again with 'win32com.client' tools one by one to generate the PDF's. Using these two tools and pyautogui for the last part: acessing another spreadsheet that would update its queries upon entry, and then copying this data to another place and printing the screen in a certain region.

## Tools
Using 'win32com.client', 'datetime', 'pyautogui' and 'os'.
