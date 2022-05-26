Attribute VB_Name = "Module11"
Sub sample1()

Dim end_row1 As Variant
Dim end_row2 As Variant
Dim end_row3 As Variant
Dim xx, MAXx As Long
Dim yy, MAXy As Long
Dim CellMer As Range
Dim CellStr As Variant
Dim deviation As Variant
Dim CellAdr As String
Dim data_base As String
Dim insert_switch As Boolean

'file_name = Cells()

Dim dev As Object
Set dev = CreateObject("Scripting.Dictionary")
dev.Add "A#2000", 0.1
dev.Add "A#800", 0.1
dev.Add "A#320", 0.1

dev_value = 0.1
deviation = Array(0.1)

data_base = "基準片"
rn_col = 46
MAXx = 47
MAXy = 17
insert_switch = True

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

rn = Range(Cells(8, 2), Cells(end_row1, 47))
rn_lenght = UBound(rn) - LBound(rn) + 1

Workbooks.Open "C:\Users\c_ip\Desktop\ｻﾝﾌﾟﾙﾃｽﾄ社内用ﾃﾞｰﾀまとめ(Ver.21.06.29) .xlsx"
ActiveWorkbook.Worksheets("一覧").Activate

  Dim sh As Worksheet

  For Each sh In Worksheets
    sh.AutoFilterMode = False
  Next sh
  
  
  
  
With Worksheets("装置データ")
    Set Rng = .ListObjects("装置データ平均値").Range
    rn_row = 1
    For col = 1 To rn_lenght
        Rng.AutoFilter
        Rng.AutoFilter Field:=1, Criteria1:=rn(col, 4)
        Rng.AutoFilter Field:=2, Criteria1:=rn(col, 9)
        Rng.AutoFilter Field:=3, Criteria1:=rn(col, 10)
        Rng.AutoFilter Field:=4, Criteria1:=rn(col, 15)
        With Rng.CurrentRegion.Offset(1, 0)
            counter = 0
            i = 0
            For Each address_filter In .SpecialCells(xlCellTypeVisible)
                
continue:
                For Each cell In address_filter
                    temp = cell.Value
                    counter = counter + 1
                    If counter > 4 And counter < 6 Then
                    'counter mean only count to Ra, if want to check the further element _
                    just rise the number of counter
                        'MsgBox (temp * (1 + deviation))
                        
                        If Not (IsEmpty(dev(rn(rn_row, 4)))) Then
                            dev_value = dev(rn(rn_row, 4))
                        End If
                        If IsEmpty(rn(rn_row, 24 + i)) Then
                            GoTo continue
                        ElseIf (((temp * (1 + dev_value)) >= rn(rn_row, 24 + i) And _
                            (temp * (1 - dev_value)) <= rn(rn_row, 24 + i))) Then
                            i = i + 1
                        Else
                            temp_str = ""
                            Select Case i
                                Case 0
                                    temp_str = "エア消費量"
                                Case 1
                                    temp_str = "ポンプ圧"
                                Case 2
                                    temp_str = "スラリー流量"
                            End Select
                            MsgBox (rn(rn_row, 1) & " " & rn(rn_row, 5) & " の " & temp_str & _
                            "データが異常です、修正してください")
                            'MsgBox (dev(rn(rn_row, 5)))
                            insert_switch = False
                            i = i + 1
                        End If
                    End If
                Next
            Next
            rn_row = rn_row + 1
        End With
        Rng.AutoFilter
    Next col
End With
  
  
  

With Worksheets("基準片")
    Set Rng = .ListObjects("基準片データ平均値").Range
    rn_row = 1
    For col = 1 To rn_lenght
        Rng.AutoFilter
        Rng.AutoFilter Field:=1, Criteria1:=rn(col, 5)
        Rng.AutoFilter Field:=2, Criteria1:=rn(col, 9)
        Rng.AutoFilter Field:=3, Criteria1:=rn(col, 10)
        Rng.AutoFilter Field:=4, Criteria1:=rn(col, 15)
        With Rng.CurrentRegion.Offset(1, 0)
            counter = 0
            i = 0
            
            current_autofilter_count = WorksheetFunction.Count(Rng.Cells.SpecialCells(xlCellTypeVisible))
            If current_autofilter_count = 0 Then
                GoTo filter_end
            End If
            
            
            For Each address_filter In .SpecialCells(xlCellTypeVisible)
continue1:
                For Each cell In address_filter
                    temp = cell.Value
                    counter = counter + 1
                    If counter > 4 And counter < 7 Then
                    'counter mean only count to Ra, if want to check the further element _
                    just rise the number of counter
                        'MsgBox (temp * (1 + deviation))
                        If Not (IsEmpty(dev(rn(rn_row, 5)))) Then
                            dev_value = dev(rn(rn_row, 5))
                        End If
                        If IsEmpty(rn(rn_row, 24 + i)) Then
                            GoTo continue1
                        ElseIf (((temp * (1 + dev_value)) >= rn(rn_row, 24 + i) And _
                            (temp * (1 - dev_value)) <= rn(rn_row, 24 + i))) Then
                            i = i + 1
                        Else
                            temp_str = ""
                            Select Case i
                                Case 0
                                    temp_str = "削れ量"
                                Case 1
                                    temp_str = "Ra"
                                Case 2
                                    temp_str = "Rz"
                                Case 3
                                    temp_str = "RzJIS"
                            End Select
                            MsgBox (rn(rn_row, 1) & " " & rn(rn_row, 5) & " の " & temp_str & _
                            "データが異常です、修正してください")
                            'MsgBox (dev(rn(rn_row, 5)))
                            insert_switch = False
                            i = i + 1
                        End If
                    End If
                Next
            Next
            rn_row = rn_row + 1
filter_end:
        End With
        Rng.AutoFilter
    Next col
End With
'******************************************************************************************


'******************************************************************************************

ActiveWorkbook.Worksheets("一覧").Activate
end_row2 = Range("B6").End(xlDown).Row
insert_switch = False
If insert_switch Then
    For i = 0 To rn_lenght - 1
        For j = 0 To rn_col - 1
            Cells(end_row2 + 1 + i, 2 + j).Value = rn(i + 1, j + 1)
            Cells(end_row2 + 1 + i, 2 + j).HorizontalAlignment = xlCenter
            'Cells(end_row2 + 1 + i, 2 + j).HorizontalAlignment = ylCenter
        Next j
    Next i
End If

end_row3 = Range("B8").End(xlDown).Row
Range(Cells(6, 2), Cells(end_row3, 21)).Borders(xlInsideHorizontal).LineStyle = True

'ActiveWorkbook.Save
ActiveWorkbook.Close


End Sub

