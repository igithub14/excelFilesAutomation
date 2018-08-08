' TO BE DONE: make this scipt recursively run inside a subfodler system.
Sub ManageFiles()
'
' 
' twttw
'
' Scelta rapida da tastiera: CTRL+i
'
   

    Dim MyFile As String, Str As String, MyDir As String, Wb As Workbook
    Dim Rws As Long, Rng As Range
    Set Wb = ThisWorkbook
    'change the address to suite
      MyDir = "C:\<PATH>"
    MyFile = Dir(MyDir & "*.xlsm")    'change file extension
    ChDir MyDir
    Application.ScreenUpdating = 0
    Application.DisplayAlerts = 0

    Do While MyFile <> ""
        Workbooks.Open (MyFile)
        With Worksheets("Report")
        
        Rows("1:8").Select
        Selection.Delete Shift:=xlUp
        
        Columns("E:E").Insert
 
        Dim rowCount As Integer
        rowCount = 0
        rowCount = Range("D:D").Cells.SpecialCells(xlCellTypeConstants).Count
        rowCount = rowCount - 1
        MsgBox "value is " & rowCount
        Dim Copyrange As String
        Startrow = 2
        Lastrow = rowCount
        Let Copyrange = "E" & Startrow & ":" & "E" & Lastrow
        Range("E2").Formula = "=(D2-Q2)*1440"
        Range("E2").Copy
        Range(Copyrange).PasteSpecial (xlPasteAll)
        
        ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
        
        Sheets("Foglio1").Activate
        Cells(1, 1) = "MEDIA"
        Cells(2, 1) = "MAX"
        Cells(3, 1) = "MIN"
        Cells(4, 1) = "N OCCORRENZE"
        Cells(5, 1) = "N < 20"
        
        Cells(1, 2).FormulaLocal = "=MEDIA(Report!E2:E" & rowCount & ")"
        Cells(2, 2).Formula = "=MAX(Report!E2:E" & rowCount & ")"
        Cells(3, 2).Formula = "=MIN(Report!E2:E" & rowCount & ")"
        Cells(4, 2).FormulaLocal = "=CONTA.NUMERI(Report!E2:E" & rowCount & ")"
        Cells(5, 2).FormulaLocal = "=CONTA.SE(Report!E2:E" & rowCount & ";" & Chr(34) & "<20" & Chr(34) & ")"

        
        ActiveWorkbook.Close True
        End With
        MyFile = Dir()
    Loop
   MsgBox "Terminato con successo."
End Sub
