Attribute VB_Name = "Mod_ExcelCell"
''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �w��Z�����Q�Ɖ\�����`�F�b�N����B
' �����P�@�F �Z���̃A�h���X�B
' �Ԃ�l�@�F �Q�Ɖ\�ȏꍇ�A�^��Ԃ��B
' �g�p���@�F If CheckRange("K302") Then
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CheckRange(strRange As String) As Boolean
    Dim rng As Range
    
    ' �����Z���̃A�h���X���Q�Ƃ���B
    On Error GoTo RangeError
    Set rng = Range(strRange)
    Set rng = Nothing
    
    ' ��O���������Ȃ������ׁA�����B
    CheckRange = True
    Exit Function

RangeError:
    ' ��O�����B
    CheckRange = False
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �w��Z���͈͂���V���ȃZ���͈͂�Ԃ��܂��B
' �����P�@�F �Z���͈͂̃I�u�W�F�N�g�B
' �����Q�@�F ���Z�q�B���啶���������͖₢�܂���B
' �����R�@�F ���Z�q�ɑ΂���������l�B
' �Ԃ�l�@�F �V���ȃZ���͈͂�Ԃ��B
' �g�p���@�F Dim new_range As Range
'            new_range = getRangeCell(cell_range, "top",    2)  ' �Z���͈͂̏�2�s��V���ȃZ���͈͂Ƃ��ĕԂ��B
'            new_range = getRangeCell(cell_range, "left",   2)  ' �Z���͈͂̍�2���V���ȃZ���͈͂Ƃ��ĕԂ��B
'            new_range = getRangeCell(cell_range, "right",  2)  ' �Z���͈͂̉E2���V���ȃZ���͈͂Ƃ��ĕԂ��B
'            new_range = getRangeCell(cell_range, "bottom", 2)  ' �Z���͈͂̉�2�s��V���ȃZ���͈͂Ƃ��ĕԂ��B
'            new_range = getRangeCell(cell_range, "topLimit",    10)  ' �Z���͈͂���10�s�𒴂��Ă���ꍇ�A��10�s��V���ȃZ���͈͂Ƃ��ĕԂ��B
'            new_range = getRangeCell(cell_range, "leftLimit",   10)  ' �Z���͈͂���10��𒴂��Ă���ꍇ�A��10���V���ȃZ���͈͂Ƃ��ĕԂ��B
'            new_range = getRangeCell(cell_range, "rightLimit",  10)  ' �Z���͈͂��E10��𒴂��Ă���ꍇ�A�E10���V���ȃZ���͈͂Ƃ��ĕԂ��B
'            new_range = getRangeCell(cell_range, "bottomLimit", 10)  ' �Z���͈͂���10�s�𒴂��Ă���ꍇ�A��10�s��V���ȃZ���͈͂Ƃ��ĕԂ��B
'            new_range = getRangeCell(cell_range, "topResize",     5)  ' �Z���͈͂����5�s�������΂����͈͂�V���ȃZ���͈͂Ƃ��ĕԂ��B
'            new_range = getRangeCell(cell_range, "leftResize",    5)  ' �Z���͈͂�����5��������΂����͈͂�V���ȃZ���͈͂Ƃ��ĕԂ��B
'            new_range = getRangeCell(cell_range, "rightResize",  -1)  ' �Z���͈͂��E��1�񋷂߂��͈͐V���ȃZ���͈͂Ƃ��ĕԂ��B
'            new_range = getRangeCell(cell_range, "bottomResize", -1)  ' �Z���͈͂�����1�s���߂��͈͐V���ȃZ���͈͂Ƃ��ĕԂ��B
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
            ' ��������
            getRangeCell = Nothing
            Exit Function
    End Select
    
    Set tmpCell = sheet.Cells.Range(sheet.Cells(startRow, startColumn), sheet.Cells(endRow, endColumn))
    Set getRangeCell = tmpCell
End Function

