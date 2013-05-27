Attribute VB_Name = "Mod_ExcelCell"
''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 指定セルが参照可能かをチェックする。
' 引数１　： セルのアドレス。
' 返り値　： 参照可能な場合、真を返す。
' 使用方法： If CheckRange("K302") Then
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CheckRange(strRange As String) As Boolean
    Dim rng As Range
    
    ' 引数セルのアドレスを参照する。
    On Error GoTo RangeError
    Set rng = Range(strRange)
    Set rng = Nothing
    
    ' 例外が発生しなかった為、成功。
    CheckRange = True
    Exit Function

RangeError:
    ' 例外発生。
    CheckRange = False
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 指定セル範囲から新たなセル範囲を返します。
' 引数１　： セル範囲のオブジェクト。
' 引数２　： 演算子。※大文字小文字は問いません。
' 引数３　： 演算子に対する引数数値。
' 返り値　： 新たなセル範囲を返す。
' 使用方法： Dim new_range As Range
'            new_range = getRangeCell(cell_range, "top",    2)  ' セル範囲の上2行を新たなセル範囲として返す。
'            new_range = getRangeCell(cell_range, "left",   2)  ' セル範囲の左2列を新たなセル範囲として返す。
'            new_range = getRangeCell(cell_range, "right",  2)  ' セル範囲の右2列を新たなセル範囲として返す。
'            new_range = getRangeCell(cell_range, "bottom", 2)  ' セル範囲の下2行を新たなセル範囲として返す。
'            new_range = getRangeCell(cell_range, "topLimit",    10)  ' セル範囲が上10行を超えている場合、上10行を新たなセル範囲として返す。
'            new_range = getRangeCell(cell_range, "leftLimit",   10)  ' セル範囲が左10列を超えている場合、左10列を新たなセル範囲として返す。
'            new_range = getRangeCell(cell_range, "rightLimit",  10)  ' セル範囲が右10列を超えている場合、右10列を新たなセル範囲として返す。
'            new_range = getRangeCell(cell_range, "bottomLimit", 10)  ' セル範囲が下10行を超えている場合、下10行を新たなセル範囲として返す。
'            new_range = getRangeCell(cell_range, "topResize",     5)  ' セル範囲を上へ5行引き延ばした範囲を新たなセル範囲として返す。
'            new_range = getRangeCell(cell_range, "leftResize",    5)  ' セル範囲を左へ5列引き延ばした範囲を新たなセル範囲として返す。
'            new_range = getRangeCell(cell_range, "rightResize",  -1)  ' セル範囲を右へ1列狭めた範囲新たなセル範囲として返す。
'            new_range = getRangeCell(cell_range, "bottomResize", -1)  ' セル範囲を下へ1行狭めた範囲新たなセル範囲として返す。
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
    endRow = tmpCell(tmpCell.count).Row
    endColumn = tmpCell(tmpCell.count).Column
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

