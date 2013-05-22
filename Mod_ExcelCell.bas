Attribute VB_Name = "Mod_ExcelCell"
''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 指定セルが参照可能かをチェックする。
' 引数１　： セルのアドレス。
' 返り値　： 参照可能な場合、真を返す。
' 使用方法： If CheckRange("K302") Then
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CheckRange(strRange As String) As Boolean
    Dim rng As Range

    On Error GoTo RangeError
    Set rng = Range(strRange)
    Set rng = Nothing
    
    CheckRange = True
    Exit Function

RangeError:
    CheckRange = False
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 指定セルが参照可能かをチェックする。
' 引数１　： セルのアドレス。
' 返り値　： 参照可能な場合、真を返す。
' 使用方法： Dim new_range As Range
'            new_range = getRangeCell(cell_range, "top",    2)  ' セル範囲の上2行を再選択する。
'            new_range = getRangeCell(cell_range, "left",   2)  ' セル範囲の左2列を再選択する。
'            new_range = getRangeCell(cell_range, "right",  2)  ' セル範囲の右2列を再選択する。
'            new_range = getRangeCell(cell_range, "bottom", 2)  ' セル範囲の下2行を再選択する。
'            new_range = getRangeCell(cell_range, "topLimit",    10)  ' セル範囲が上10行を超えるなら10行以内で再選択する。
'            new_range = getRangeCell(cell_range, "leftLimit",   10)  ' セル範囲が左10列を超えるなら10列以内で再選択する。
'            new_range = getRangeCell(cell_range, "rightLimit",  10)  ' セル範囲が右10列を超えるなら10列以内で再選択する。
'            new_range = getRangeCell(cell_range, "bottomLimit", 10)  ' セル範囲が下10行を超えるなら10行以内で再選択する。
'            new_range = getRangeCell(cell_range, "topResize",     5)  ' セル範囲を上へ5行引き延ばした範囲を再選択する。
'            new_range = getRangeCell(cell_range, "leftResize",    5)  ' セル範囲を左へ5列引き延ばした範囲を再選択する。
'            new_range = getRangeCell(cell_range, "rightResize",  -1)  ' セル範囲を右へ1列狭めた範囲を再選択する。
'            new_range = getRangeCell(cell_range, "bottomResize", -1)  ' セル範囲を下へ1行狭めた範囲を再選択する。
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function getRangeCell(cell As Range, command As String, arg As Integer) As Range
    Set sheet = cell.Worksheet
    Dim tmpCell As Range
    Set tmpCell = sheet.Range(cell.Address)
    
    
    Dim startRow As Long
    Dim startColumn As Long
    Dim endRow As Long
    Dim endColumn As Long
    
    startRow = tmpCell.Row
    startColumn = tmpCell.Column
    endRow = tmpCell(tmpCell.Count).Row
    endColumn = tmpCell(tmpCell.Count).Column
    Select Case Trim(LCase(command))
        Case "top"
            endRow = startRow + arg - 1
        Case "left"
            endColumn = startColumn + arg - 1
        Case "right"
            startColumn = endColumn - arg + 1
        Case "bottom"
            startRow = endRow - arg + 1
            
            
        Case "toplimit"
            If endRow - startRow + 1 > arg Then
                endRow = startRow + arg - 1
            End If
        Case "leftlimit"
            If endColumn - startColumn + 1 > arg Then
                endColumn = startColumn + arg - 1
            End If
        Case "rightmlimit"
            If endColumn - startColumn + 1 > arg Then
                startColumn = endColumn - arg + 1
            End If
        Case "bottomlimit"
            If endRow - startRow + 1 > arg Then
                startRow = endRow - arg + 1
            End If
            
            
        Case "topresize"
            startRow = startRow - arg
        Case "leftresize"
            startColumn = startColumn - arg
        Case "rightresize"
            endColumn = endColumn + arg
        Case "bottomresize"
            endRow = endRow + arg
        Case Else
            ' 処理成功
            getRangeCell = Nothing
            Exit Function
    End Select
    
    Set tmpCell = sheet.Cells.Range(sheet.Cells(startRow, startColumn), sheet.Cells(endRow, endColumn))
    Set getRangeCell = tmpCell
End Function

