Option Explicit

ExcelMacroExample
KillProcess

Sub ExcelMacroExample() 

  Dim xlApp, xlBook, strDirPath, strMacroFilePath, exists, fso

  strDirPath = "C:\vba_code" ' CHANGE the path to folder where places .csv files
  strMacroFilePath = "C:\vba_code\Macro-Workbook.xlsm"

  Set fso = CreateObject("Scripting.FileSystemObject")
  exists = fso.FolderExists(strDirPath)

  if (exists = FALSE) then 
    MsgBox "Directory path is empty, aborting...", vbCritical + vbOKOnly, "Directory not found"
    Exit Sub
  end if

  Set xlApp = CreateObject("Excel.Application") 
  Set xlBook = xlApp.Workbooks.Open(strMacroFilePath, 0, True)  
  
  XLApp.Run "'Macro-Workbook.xlsm'!Cycle", CStr(strDirPath)
  xlApp.Quit 

  Set fso = Nothing
  Set xlBook = Nothing 
  Set xlApp = Nothing
End Sub 

' Terminate Excel process 
Sub KillProcess()
  Dim Process 
  For Each Process In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = 'EXCEL.EXE'") 
    Process.Terminate 
  Next 
End Sub