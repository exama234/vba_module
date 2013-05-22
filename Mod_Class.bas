Attribute VB_Name = "Mod_Class"
''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �I�u�W�F�N�g�̕�����\�����擾����B
' �����P�@�F �ΏۂƂȂ�I�u�W�F�N�g�BVariant�^�B
' �Ԃ�l�@�F �n���ꂽ�I�u�W�F�N�g�̕�����\���B
' �g�p���@�F str = toString(obj)
'            toString(18)    <= ���l�^�̏ꍇ�A������u18�v���Ԃ�B
'            toString("21")  <= ������̏ꍇ�A������u"21"�v���Ԃ�B
'            toString(True)  <= �^�U�l�̏ꍇ�A������uTrue�v���Ԃ�B
'            toString(today) <= ���t�^�̏ꍇ�A������u#2012/01/04#�v���Ԃ�B
'            toString(array) <= �z��^�̏ꍇ�A������u[18, "21", #2012/01/04#]�v���Ԃ�B
'            toString(obj)   <= "Nothing"�̏ꍇ�A������u(Nothing)�v���Ԃ�B
'            toString(obj)   <= "Empty"�̏ꍇ�A������u(Empty)�v���Ԃ�B
'            toString(obj)   <= "Null"�̏ꍇ�A������u(Null)�v���Ԃ�B
'            toString(obj)   <= "Error"�̏ꍇ�A������u(Error)�v���Ԃ�B
'            toString(obj)   <= �I�u�W�F�N�g�̏ꍇ�A������u�N���X��[...]�v���Ԃ�B
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function toString(item As Variant) As String
    '
    Dim key, buff, tmp, delim As String
    ' select��
    Select Case TypeName(item)
        Case "Byte", "Integer", "Long", "Boolean", "Single", "Double", "Currency"
            tmp = item
        Case "String"
            tmp = """" & item & """"
        Case "Date"
            tmp = "#" & item & "#"
        Case "Nothing", "Empty", "Null", "Error"
            tmp = "(" & TypeName(item) & ")"
        Case Else
            If IsArray(item) Then
                Dim v As Variant
                For idx = LBound(item) To UBound(item)
                    If IsObject(item(idx)) Then
                        Set v = item(idx)
                    Else
                        v = item(idx)
                    End If
                    
                    tmp = tmp & delim & toString(v)
                    delim = ", "
                Next idx
                tmp = "[" & tmp & "]"
            Else
                If Mod_Class.hasMethod(item, "toString") Then
                    ' toString()������Ȃ���s�������B
                    tmp = CallByName(item, "toString", VbMethod)
                    tmp = TypeName(item) & "[" & tmp & "]"
                Else
                    tmp = TypeName(item) & "[...]"
                End If
            End If
    End Select
    
    buff = tmp
    toString = buff
End Function




''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �I�u�W�F�N�g�����\�b�h�������Ă��邩���m�F����B
' �����P�@�F �m�F�Ώۂ̃I�u�W�F�N�g�BVariant�^�B
' �����Q�@�F �m�F���郁�\�b�h�̖��́B
' �Ԃ�l�@�F ���\�b�h�����ꍇ�A�^��Ԃ��B
' �g�p���@�F If hasMethod(obj, "toString") Then
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function hasMethod(obj As Variant, method As String) As Boolean
On Error GoTo ErrHandler
    ' �I�u�W�F�N�g�Ɏw��̃��\�b�h������Ȃ���s����B
    Call CallByName(obj, method, VbMethod)
    ' ��O���������Ȃ������ׁA�����B
    hasMethod = True
    Exit Function
    
ErrHandler:
    ' ��O�����B
    hasMethod = False
End Function

