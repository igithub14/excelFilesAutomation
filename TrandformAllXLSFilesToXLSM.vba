' This script creates a copy of all the .xls files inside a folder, and save this new file with .xlsm extension.
' TO BE DONE: make this script erase .xls files.
' TO BE DONE: make this script recursivley run inside a subfolder system.
Sub TrandformAllXLSFilesToXLSM()
Dim myPath As String

myPath = "C:\<PATH_TO_THE_FOLDER>\"
WorkFile = Dir(myPath & "*.xls")

Do While WorkFile <> ""
    If Right(WorkFile, 4) <> "xlsm" Then
        Workbooks.Open Filename:=myPath & WorkFile
        ActiveWorkbook.SaveAs Filename:= _
        myPath & WorkFile & "m", FileFormat:= _
        xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
        ActiveWorkbook.Close
     End If
       
     WorkFile = Dir()
Loop
End Sub
