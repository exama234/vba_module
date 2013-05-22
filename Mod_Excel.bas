Attribute VB_Name = "Mod_Excel"
''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �w��V�[�g���Q�Ɖ\�����`�F�b�N����B
' �����P�@�F Excel�u�b�N�̐�΃p�X�B
' �����Q�@�F �V�[�g���B
' �Ԃ�l�@�F �Q�Ɖ\�ȏꍇ�A�^��Ԃ��B
' �g�p���@�F If CheckSheet("C:\test.xls", "Sheet1") Then
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CheckSheet(ByRef strBook As String, ByRef strSheet As String) As Boolean
    Dim sheetObj As Worksheet
    fileonly = Dir(strBook)

    On Error GoTo SheetError
    Set sheetObj = Workbooks(fileonly).Worksheets(strSheet)
    Set sheetObj = Nothing
    CheckSheet = True
    Exit Function

SheetError:
    CheckSheet = False
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �w��u�b�N���J���Ă��邩���`�F�b�N����B
' �����P�@�F Excel�u�b�N�̐�΃p�X�B
' �Ԃ�l�@�F �J���Ă���ꍇ�A�^��Ԃ��B
' �g�p���@�F If openCheck("C:\test.xls") Then
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function openCheck(ByRef filefullpath As String) As Boolean
    Dim book As Workbook
    For Each book In Workbooks
        If book.Name = Dir(filefullpath) Then
            Set book = Nothing
            openCheck = True
            Exit Function
        End If
    Next book
    
    Set book = Nothing
    ' �J���ĂȂ�
    openCheck = False
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �Z����I�𒆂��𔻒肷��B
' �����P�@�F �G���[���b�Z�[�W�\���t���O�B�i�f�t�H���g�l�FTrue�j
' �Ԃ�l�@�F �Z����I�𒆂���\���^�U�l�B�iTrue�F�Z���I�𒆁AFalse�F�Z���I�𒆂łȂ��j
' �g�p���@�F If isSelectionRange() Then
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function isSelectionRange(Optional msg_flg As Boolean = True) As Boolean
    If TypeName(Selection) = "Range" Then
        ' �Z���I�𒆁B�^��Ԃ��B
        isSelectionRange = True
        Exit Function
    End If
    
    ' �Z���I�𒆂łȂ��B�K�v�Ȃ�G���[�\������B
    If msg_flg Then
        MsgBox ("�Z����I�����Ă��������B")
    End If
End Function




