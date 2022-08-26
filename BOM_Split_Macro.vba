Sub SplitBOM()
Dim row, lastRow As Long
Dim splitValues As Variant
Dim delimiter As String
delimiter = Application.InputBox("Enter character to split by", "Enter delimiter", ",")
Application.ScreenUpdating = False
lastRow = Cells(Rows.Count, "A").End(xlUp).row
For row = lastRow To 2 Step -1
    If InStr(Cells(row, "C"), ",") Then
        splitValues = Split(Trim(Cells(row, "C")), delimiter)
        Rows(row + 1).Resize(UBound(splitValues)).Insert
        Cells(row, "A").Resize(UBound(splitValues) + 1, 1).Value = Cells(row, "A").Resize(, 1).Value
        Cells(row, "B").Resize(UBound(splitValues) + 1, 1).Value = Cells(row, "B").Resize(, 1).Value
        Cells(row, "C").Resize(UBound(splitValues) + 1) = Application.Transpose(splitValues)
        Cells(row, "D").Resize(UBound(splitValues) + 1) = Cells(row, "D") / (UBound(splitValues) + 1)
        
        ReDim seqArr(0 To UBound(splitValues)) As Long
        For i = 0 To UBound(seqArr)
            seqArr(i) = i + 1
        Next i
        
        Cells(row, "E").Resize(UBound(seqArr) + 1) = Application.Transpose(seqArr)
    Else
        Cells(row, "E").Value = 1
    End If
Next row

Cells(1, "E").Value = "Guideline Seq"
temp = Columns("E").Value
Columns("E").Delete
Columns("B").Insert
Columns("B").Value = temp

Application.ScreenUpdating = True
End Sub

