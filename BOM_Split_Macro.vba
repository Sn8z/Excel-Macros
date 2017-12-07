Sub SplitBOM()
Dim r, lr As Long
Dim Sp As Variant
Application.ScreenUpdating = False
lr = Cells(Rows.Count, 3).End(xlUp).Row
For r = lr To 2 Step -1
    If InStr(Cells(r, 20), ",") Then
        Sp = Split(Trim(Cells(r, 20)), ",")
        Rows(r + 1).Resize(UBound(Sp)).Insert
        Cells(r, 1).Resize(UBound(Sp) + 1, 16).Value = Cells(r, 1).Resize(, 16).Value
        Cells(r, 18).Resize(UBound(Sp) + 1, 19).Value = Cells(r, 18).Resize(, 19).Value
        Cells(r, 21).Resize(UBound(Sp) + 1, 21).Value = Cells(r, 21).Resize(, 21).Value
        Cells(r, 17).Resize(UBound(Sp) + 1) = Cells(r, 17) / (UBound(Sp) + 1)
        Cells(r, 20).Resize(UBound(Sp) + 1) = Application.Transpose(Sp)
    End If
Next r
Application.ScreenUpdating = True
End Sub
