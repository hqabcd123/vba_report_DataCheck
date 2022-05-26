VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

Dim hlc As Variant
Dim Ehlc() As Variant
Dim sht As Worksheet
Dim temp() As Variant
Dim keyValue As Object
Dim reff_len As Integer
Dim reff_start As Integer
Dim counter As Integer
Dim rng_len As Integer

Set sht = Worksheets("【変更不可】")
reff_len = 41 'もし変更不可のところが変えた場合はこっちの数字（長さ）を修正する必要がある
reff_start = 10
counter = 0
rng_len = 0
ReDim Ehlc(reff_len - reff_start)

For i = 0 To 99
  cell = Cells(16 + i, 2)
  If StrComp(cell, "▼処理条件表") Or StrComp(cell, "▼固定条件") Then
    rng_len = i - 2
    Exit For
  End If
Next i

If Target.Row < 16 And Target.Row > rng_len Then
  Exit Sub
End If

For i = reff_start To reff_len
  ' Ehlc(i - reff_start) = sht.Cells(i, 35)
  Ehlc(i - reff_start) = sht.Range(Cells(i, 35), Cells(i, 35))
  te = sht.Cells(i, 35)
  keyValue.Add sht.Cells(i, 35), i
Next i

counter = 0
For Each hlc In Ehlc
  If hlc.Hyperlinks.Count > 0 And hlc = Target Then
    With hlc
      ActiveSheet.Hyperlinks.Add Target, hlc.Hyperlinks.Item(1).Address
    End With
  End If
Next
End Sub
