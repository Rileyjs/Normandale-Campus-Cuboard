# Programming-Club
# Master Version 1.04.02

***Not up-to-date***

To contribute, go to the Cupboard-Template Folder and download "New Year Workbook Template.xltm" and the "Master Sheet Macros.bas". Also go to the Test-Data folder and download "Generated Data.xlsm".

Open the "Master Sheet Macros.bas" with Notepad (or Notepad++, if you have it). Use your Find function (ctrl-f in Windows and Linux), 
search for "SaveAs". This should take you to two lines of code. Highlight the line that says "UPDATE THIS ONE ->" and make a copy of it one line down. Remove everything before ThisWorkbook (including the ') on the copy and change the path to a Test Folder of your choosing (Make sure this is separate from the Test-Data folder and is not pushed to git).

Next, open the "New Year Workbook Template.xltm". If you do not have the Office 365 suite, you can download it through your school mail. A blank Excel sheet will load. If you do not have the Developer tab at the top, right-click on the ribbon and select "Customize the ribbon". On the right, select Developer and click ok. Navigate to the Developer Tab, find "View Code" and click it. In the VBA Editor, on the left side, go to Modules and right click on Module1 and remove it. Do not export it. go to File -> "Import File". Find the "Master eet Macros.bas" and import it.
 
Now you can return to the excel sheet. Left-click on Macros and run the "NewWorkbookONLY" function. Fill in whatever dates you see fit. With this sheet open, go to the Generated Data sheet in the Test-Data folder. Select all of the tabs at the bottom, right-click and select "Move or Copy". Check Copy at the bottom and in the drop down box choose the sheet that says "Master". Click ok.  You now have a working test workbook. 

Let me know if you have any further questions. Please read all issues before continuing.

 ***Important Note***
 
 Before you push anything to git, make sure you have deleted your file path in both "Master Sheet Macros.bas" and "New Year Workbook Template.xltm". If you do not know what I mean, please contact me.
