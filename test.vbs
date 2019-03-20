Sub Test()


    LastRow = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row

    for r = 9 to LastRow
        if cells(r,2).value <> "" then
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        end if
    next r

End Sub


Sub foo()
  Dim r As Range, rows As Long, i As Long
  Set r = ActiveSheet.Range("A1:Z50")
  rows = r.rows.Count
  For i = rows To 1 Step (-1)
    If WorksheetFunction.CountA(r.rows(i)) = 0 Then r.rows(i).Add BeforeRow:=Selection.rows(i)
  Next
End Sub
