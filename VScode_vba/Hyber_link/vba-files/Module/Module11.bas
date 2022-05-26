Attribute VB_Name = "Module11"
Sub sample1()

Dim end_row1 As Variant
Dim end_row2 As Variant
Dim end_row3 As Variant
Dim xx, MAXx As Long
Dim yy, MAXy As Long
Dim CellMer As Range
Dim CellStr As Variant
Dim CellAdr As String

MAXx = 47
MAXy = 17

For yy = 8 To MAXy

For xx = 2 To MAXx

If Cells(yy, xx).MergeCells = True Then

CellAdr = Cells(yy, xx).MergeArea.Address
CellStr = Cells(yy, xx).Value
Range(CellAdr).UnMerge

For Each CellMer In Range(CellAdr)
CellMer.Value = CellStr
Next CellMer
End If

Next

Next


end_row1 = Range("B8").End(xlDown).Row

If end_row1 = 1048576 Then
end_row1 = 8
End If

Range(Cells(8, 18), Cells(end_row1, 18)).Borders(xlInsideHorizontal).LineStyle = True
Range(Cells(8, 18), Cells(end_row1, 18)).Copy
Range(Cells(8, 18), Cells(end_row1, 18)).PasteSpecial Paste:=xlPasteValues

Range(Cells(8, 2), Cells(end_row1, 47)).Borders(xlInsideHorizontal).LineStyle = True
Range(Cells(8, 2), Cells(end_row1, 47)).Copy

Workbooks.Open "\\macoho_cad\B-Tech\»ÝÌßÙÃ½Ä‹Æ–±\ŽÐ“à—pÃÞ°À\»ÝÌßÙÃ½ÄŽÐ“à—pÃÞ°À‚Ü‚Æ‚ß(Ver.21.10.25) .xlsx"
ActiveWorkbook.Worksheets("ˆê——").Activate

  Dim sh As Worksheet

  For Each sh In Worksheets
    sh.AutoFilterMode = False
  Next sh


end_row2 = Range("B6").End(xlDown).Row
Cells(end_row2 + 1, 2).PasteSpecial

end_row3 = Range("B8").End(xlDown).Row
Range(Cells(6, 2), Cells(end_row3, 21)).Borders(xlInsideHorizontal).LineStyle = True



ActiveWorkbook.Save
ActiveWorkbook.Close


End Sub

