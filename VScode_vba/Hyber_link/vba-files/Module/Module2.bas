Attribute VB_Name = "Module2"
Option Explicit


Sub JIS1994_2()
    Dim varFileName As Variant
    Dim i As Variant
    Dim j As Variant


For i = 6 To 34 Step 7
    varFileName = Application.GetOpenFilename(FileFilter:="CSVファイル(*.csv),*.csv", _
                                        Title:="CSVファイルの選択")
    If varFileName = False Then
        Exit Sub
    End If
    
Workbooks.Open Filename:=varFileName

Range("B2").Select
    Selection.Copy
    ThisWorkbook.Activate
   Cells(i, 9).Select
    ActiveSheet.Paste

Workbooks.Open Filename:=varFileName

   Range("D46,F46,H46,D60,F60,H60,D74,F74,H74").Select
    Selection.Copy
       ThisWorkbook.Activate
   Cells(i, 12).PasteSpecial Paste:=xlPasteValues


    Workbooks.Open Filename:=varFileName
    ActiveWorkbook.Close
    
    Next i

    End Sub

