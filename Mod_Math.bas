Attribute VB_Name = "Mod_Math"
''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �Q�̐��l�̂����A�l�̑傫�Ȃ��̂��擾����B
' �����P�@�F ���l������킷�I�u�W�F�N�g�BVariant�^�B
' �����Q�@�F ���l������킷�I�u�W�F�N�g�BVariant�^�B
' �Ԃ�l�@�F �n���ꂽ�����ő傫�Ȓl�B
' �g�p���@�F val = max(300, 299.99)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function max(arg1 As Variant, arg2 As Variant) As Variant
	' �e�X�g�ǉ��B
    If (arg1 > arg2) Then
        max = arg1
    Else
        max = arg2
    End If
End Function

