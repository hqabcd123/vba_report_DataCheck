Attribute VB_Name = "Module7"
Option Explicit

Sub JIS2013_2()
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
   Cells(i, 8).Select
    ActiveSheet.Paste

Workbooks.Open Filename:=varFileName

   Range("L50, L68, L86").Select
    Selection.Copy
       ThisWorkbook.Activate
   Cells(i, 11).PasteSpecial Paste:=xlPasteValues

Workbooks.Open Filename:=varFileName

   Range("H50, H68, H86").Select
    Selection.Copy
       ThisWorkbook.Activate
   Cells(i, 12).PasteSpecial Paste:=xlPasteValues



    Workbooks.Open Filename:=varFileName
    ActiveWorkbook.Close
    
    Next i


    End Sub

