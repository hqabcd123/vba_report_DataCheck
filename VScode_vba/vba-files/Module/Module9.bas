Attribute VB_Name = "Module9"
Option Explicit


Sub JIS2013_4()


        Dim varFileName As Variant
    Dim i As Variant
    Dim j As Variant


For i = 6 To 34 Step 7
    varFileName = Application.GetOpenFilename(FileFilter:="CSV�t�@�C��(*.csv),*.csv", _
                                        Title:="CSV�t�@�C���̑I��")
    If varFileName = False Then
        Exit Sub
    End If
    
Workbooks.Open Filename:=varFileName

Range("B2").Select
    Selection.Copy
    ThisWorkbook.Activate
   Cells(i, 20).Select
    ActiveSheet.Paste

Workbooks.Open Filename:=varFileName

   Range("L50, L68, L86").Select
    Selection.Copy
       ThisWorkbook.Activate
   Cells(i, 23).PasteSpecial Paste:=xlPasteValues

Workbooks.Open Filename:=varFileName

   Range("H50, H68, H86").Select
    Selection.Copy
       ThisWorkbook.Activate
   Cells(i, 24).PasteSpecial Paste:=xlPasteValues



    Workbooks.Open Filename:=varFileName
    ActiveWorkbook.Close
    
    Next i

End Sub