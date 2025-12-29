Attribute VB_Name = "CopyTo_DB"
' Sub function that copy the data from Sheet(Dafl) to database sheet

Sub copy_pasteDB(DB_Table As String)


If DB_Table = "Dafls" Then 'move data from sheet(Dafl) to "Dafls" data table

    Sheets(2).Select '"Dafl"
    Range("B1").Copy

    Worksheets("DB").Select
    Range("A10000").End(xlUp).Offset(1, 9).PasteSpecial (xlPasteValues) 'Copy to Dafl# Column
    'Sheets(2).Range("A3:I3").Copy (Sheets(3).Range("A1").End(xlUp))


    Sheets(2).Select
    Range("A3:I3").Copy

    Worksheets("DB").Select
    Range("A10000").End(xlUp).Offset(1, 0).PasteSpecial
    'Sheets(2).Range("A3:I3").Copy (Sheets(3).Range("A1").End(xlUp))
End If





If DB_Table = "Bath_Repl" Then 'move all table from sheet("Bath_Repl") to "Bath_Repl" data table



Dim lastRow As Long


Worksheets("Bath_Repl").Select             'Copy to Bath_Repl Column
lastRow = Range("B25").End(xlUp).Row
Range(Range("B4"), Cells(lastRow, 8)).Copy


Worksheets("DB").Select
Range("S10000").End(xlUp).Offset(1, 0).PasteSpecial (xlPasteValues)

End If









End Sub
