' Heavily based on Kim Greenlee great work at http://krgreenlee.blogspot.com/2006/04/excel-running-excel-on-windows-task.html
' Create a WshShell to get the current directory
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")

' Create an Excel instance
Dim myExcelWorker
Set myExcelWorker = CreateObject("Excel.Application") 

' Disable Excel UI elements
myExcelWorker.DisplayAlerts = False
myExcelWorker.AskToUpdateLinks = False
myExcelWorker.AlertBeforeOverwriting = False
myExcelWorker.FeatureInstall = msoFeatureInstallNone

' Tell Excel what the current working directory is 
Dim strSaveDefaultPath
Dim strPath
strSaveDefaultPath = myExcelWorker.DefaultFilePath
strPath = "PATH - RUTA LOCAL"
myExcelWorker.DefaultFilePath = strPath

' Open the Workbook specified on the command-line 
Dim oWorkBook
Dim strWorkerWB
strWorkerWB = strPath & "\FILENAME.xlsm"

Set oWorkBook = myExcelWorker.Workbooks.Open(strWorkerWB)

' Build the macro name with the full path to the workbook - here, FILENAME is without the xlsm extension. 
' Check MODULES name/number. This example directs to 2 diff modules.
Dim strMacroName1
strMacroName1 = "'" & strPath & "\FILENAME'" & "!Módulo1.NOMBRE-SUB"
Dim strMacroName2
strMacroName2 = "'" & strPath & "\FILENAME'" & "!Módulo2.NOMBRE-SUB"


on error resume next 
   ' Run the calculation macro in order
   myExcelWorker.Run strMacroName1
   myExcelWorker.Run strMacroName2

   if err.number <> 0 Then
      ' Error occurred - just close it down.
   End If

   err.clear
on error goto 0 

'If needed, save a copy of the excel file with today's date
' Delete next 4 lines if not needed
Dim sDate 
sDate = FormatDateTime(Now(), 1)


oWorkBook.SaveAs "PATH\NEWFILENAME"&sDate&".xlsm", 52
oWorkBook.Close (False)

' myExcelWorker.DefaultFilePath = strSaveDefaultPath

' Clean up and shut down
Set oWorkBook = Nothing

' Don’t Quit() Excel if there are other Excel instances 
' running, Quit() will shut those down also
if myExcelWorker.Workbooks.Count = 0 Then
   myExcelWorker.Quit
End If

Set myExcelWorker = Nothing
Set WshShell = Nothing