' TO BE DONE: make this scipt recursively run inside a subfodler system.
Sub ManageFiles()
'
' ManageFiles Macro
' twttw
'
' Scelta rapida da tastiera: CTRL+i
'
   

    Dim MyFile As String, Str As String, MyDir As String, Wb As Workbook
    Dim Rws As Long, Rng As Range
    Set Wb = ThisWorkbook
    'change the address to suite
    MyDir = "C:\PATH_TO_THE_FOLDER\"
    MyFile = Dir(MyDir & "*.xlsm")    'change file extension
    ChDir MyDir
    Application.ScreenUpdating = 0
    Application.DisplayAlerts = 0

    Do While MyFile <> ""
        Workbooks.Open (MyFile)
        With Worksheets("Report")
        
' cancella righe inutili
        Rows("1:8").Select
        Selection.Delete Shift:=xlUp
        
' inserisce colonna E
        Columns("E:E").Insert
 
' conta quante righe ci sono e poi inserisce formula in colonna E
        Dim rowCount As Integer
        rowCount = 1
        rowCount = Range("I:I").Cells.SpecialCells(xlCellTypeConstants).Count
        MsgBox "value is " & rowCount
        Dim Copyrange As String
        Startrow = 3
        Lastrow = rowCount
        Let Copyrange = "E" & Startrow & ":" & "E" & Lastrow
        Range("E2").Formula = "=(D2-Q2)*1440"
        Range("E2").Copy
        Range(Copyrange).PasteSpecial (xlPasteAll)
        
        
        ActiveWorkbook.Close True
        End With
        MyFile = Dir()
    Loop

End Sub
