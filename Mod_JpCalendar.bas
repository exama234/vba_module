Public Enum era
    ' �g�p�\��Ȃ��B
    ���� = 0
    �吳 = 1
    ���a = 2
    ���� = 3
End Enum



''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �����w��̘a��N���������݂��邩���`�F�b�N����B
' �@�@�@�@�F ����������6�N�ȍ~��ΏۂƂ��Ă��܂��B
' �@�@�@�@�F ���邤�N�l���B
' �@�@�@�@�F �����̉���ɂ�邸��͔�l���B�܂薾��1~5�N�͐����������ł��Ă܂���B
' �����P�@�F �����B
' �����Q�@�F �N�B
' �����R�@�F ���B
' �����S�@�F ���B
' �Ԃ�l�@�F �a��Ƃ��đ��݂���ꍇ�A�^��Ԃ��B
'
' �ˑ�Mod �F �v�uMod_Date�v���W���[��
' �g�p���@�F If isExistJpCalendar('���a', 22, 3, 4) Then
'         �F     MsgBox ("�a��Ƃ��đ��݂��Ȃ��B")
'         �F End If
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function isExistJpCalendar(arg_era As String, arg_year As Integer, arg_month As Integer, arg_day As Integer) As Boolean
    Dim date1 As Date
    date1 = getDominicalDate(arg_era, arg_year, arg_month, arg_day)
    
    If date1 = 0 Then
        ' �a��Ƃ��đ��݂��Ȃ��B
        Exit Function
    End If
    
    ' �a��Ƃ��đ��݂���B
    isExistJpCalendar = True
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �����w��̘a��N��������t�^�ɕϊ�����B
' �@�@�@�@�F ����������6�N�ȍ~��ΏۂƂ��Ă��܂��B
' �@�@�@�@�F ���邤�N�l���B
' �@�@�@�@�F �����̉���ɂ�邸��͔�l���B�܂薾��1~5�N�͐����������ł��Ă܂���B
' �����P�@�F �����B
' �����Q�@�F �N�B
' �����R�@�F ���B
' �����S�@�F ���B
' �Ԃ�l�@�F �a��Ƃ��đ��݂���ꍇ�A���t�^��Ԃ��B
' �@�@�@�@�F �ϊ��Ɏ��s�����ꍇ�A�u0:00:00�i���l�Ƃ��Ă̓[���j�v��Ԃ��B
'
' �ˑ�Mod �F �v�uMod_Date�v���W���[��
' �g�p���@�F date1 = Mod_JpCalendar.getDominicalDate("���a", 22, 3, 4)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function getDominicalDate(arg_era As String, arg_year As Integer, arg_month As Integer, arg_day As Integer) As Date
    If arg_year < 1 Then
        ' ���݂��Ȃ��N�B
        Exit Function
    End If
    If arg_month < 1 Or 12 < arg_month Then
        ' ���݂��Ȃ����B
        Exit Function
    End If
    If arg_day < 1 Or 31 < arg_day Then
        ' ���݂��Ȃ����B
        Exit Function
    End If
    
    Select Case Strings.Trim(arg_era)
        ' �y���m���z
        ' �c��4�N9��8���B
        ' �u�������疾�����g����[�B
        ' �ł����肪�������獡�܂ł̌c��4�N1��1���`9��8���͖������N1��1���`9��8���ł������[�B
        ' ���Ԃ邯�ǋC�ɂ��Ȃ��łˁB�v
        ' �y���m���z
        ' ����45�N7��30���B
        ' �u��������吳���g����[�B���������͖���45�N7��30���ł��吳���N7��30���ł������[�B
        ' ���Ԃ邯�ǋC�ɂ��Ȃ��łˁB�v
        ' �y���m���z
        ' �吳15�N12��25���B
        ' �u�������珺�a���g����[�B���������͑吳15�N12��25���ł����a���N12��25���ł������[�B
        ' ���Ԃ邯�ǋC�ɂ��Ȃ��łˁB�v
        ' �y���m���z
        ' ���a64�N1��7���B
        ' �u�����͏��a64�N1��7���A�����͕������N1��8������[�B����͂��Ԃ�Ȃ�������S���ĂˁB�v
        Case "����"
            ' 1868/01/01 - 1912/07/30�i����45�N�܂Łj
            start_date = #1/1/1868#
            end_date = #7/30/1912#
        Case "�吳"
            ' 1912/07/30 - 1926/12/25�i�吳15�N�܂Łj
            start_date = #7/30/1912#
            end_date = #12/25/1926#
        Case "���a"
            ' 1926/12/25 - 1989/01/07�i���a64�N�܂Łj
            start_date = #12/25/1926#
            end_date = #1/7/1989#
        Case "����"
            ' 1989/01/08 - 20XX/XX/XX�i����� �₿��Ɂj
            start_date = #1/8/1989#
            end_date = #12/31/9999#
        Case Else
            ' ���݂��Ȃ������B
            Exit Function
    End Select
    
    ' ����i������j�ɕϊ�����B
    Dim base_year, dominical_year As Integer
    base_year = Year(start_date)
    dominical_year = base_year + arg_year - 1
    str_yyyymmdd = dominical_year & "/" & CStr(arg_month) & "/" & CStr(arg_day)
    If dominical_year < 1873 Then
        ' ����6�N���O�͑ΏۊO�Ƃ���B
        Exit Function
    End If
    
    
    Dim last_day As Variant
    last_day = Array(-1, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    ' ���W���[���Ăяo���i���邤�N����j
    If Mod_Date.isIntercalaryYear(dominical_year) Then
        ' ���邤�N2����29���܂ŁB
        last_day(2) = 29
    End If
    If last_day(arg_month) < arg_day Then
        ' ���݂��Ȃ����B
        Exit Function
    End If
    
    ' ����i���t�^�j�ɕϊ�����B
    Dim dominical_date As Date
    dominical_date = DateValue(str_yyyymmdd)
    
    
    If dominical_date < start_date Or end_date < dominical_date Then
        ' ���݂��Ȃ����t�B�N���̍ŏI���𒴂��Ă���B
        Exit Function
    End If
    
    
    ' �a��Ƃ��đ��݂���B
    getDominicalDate = dominical_date
End Function

