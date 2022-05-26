VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub sample1()

Dim end_row1 As Variant
Dim end_row2 As Variant


end_row1 = Range("M8").End(xlDown).Row
Range(Cells(8, 2), Cells(end_row1, 21)).Copy

Workbooks.Open "C:\Users\t_hayakawa\Desktop\test2.xlsx"
ActiveWorkbook.Worksheets("ˆê——").Activate

end_row2 = Range("M5").End(xlDown).Row
Cells(end_row2 + 1, 2).PasteSpecial

ActiveWorkbook.Save
ActiveWorkbook.Close


End Sub

