Attribute VB_Name = "Mod_Date"
''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �w��N���[�N�����`�F�b�N����B
' �����P�@�F ����N�B
' �Ԃ�l�@�F �[�N�̏ꍇ�A�^��Ԃ��B
' �g�p���@�F If isIntercalaryYear(2020) Then
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function isIntercalaryYear(arg_year As Integer) As Boolean
    If arg_year Mod 400 = 0 Then
        ' 400�Ŋ����ꍇ�A�[�N�ƂȂ�B
        isIntercalaryYear = True
        Exit Function
    End If
    If arg_year Mod 100 = 0 Then
        ' �i400�Ŋ���Ȃ� ���j100�Ŋ����ꍇ�A�[�N�ƂȂ�Ȃ��B
        isIntercalaryYear = False
        Exit Function
    End If
    If arg_year Mod 4 = 0 Then
        ' �i100�Ŋ���Ȃ� ���j4�Ŋ����ꍇ�A�[�N�ƂȂ�B
        isIntercalaryYear = True
        Exit Function
    End If
    
    ' �i100�Ŋ���Ȃ� ���j4�Ŋ���Ȃ��ꍇ�A�[�N�ƂȂ�Ȃ��B
    isIntercalaryYear = False
End Function

