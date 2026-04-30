This repository consists codes I wrote while teaching myself VBA. 

To create a VBA module, press Alt+L+V (full lou).
In the new Window that opens, insert a new module.

Code your stuff and run it. 
Alternatively, you can insert a button to evoke a sub procedure without opening the vba sheet, or create a .vbs file to run the procedure without opening the Excel. 

To create a .vbs file, open notepad and paste the below:
/*
Dim xlApp, xlBook
Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.Workbooks.Open("C:\YourPath\YourFile.xlsm")
xlApp.Run "subName"
xlBook.Save
xlApp.Quit
*/

Save this file as a .vbs file. You can directly click on it to run. Else:
Open Task Scheduler (search it in Start menu)
Click Create Basic Task
Set the trigger to Weekly on your preferred day/time
Set the action to Start a Program → point it to wscript.exe
In "Add arguments" put the full path to your .vbs file
Finish
