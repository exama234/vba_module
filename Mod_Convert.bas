Attribute VB_Name = "Mod_Convert"
''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �I�u�W�F�N�g�𐔒l�iInteger�j�Ƃ��ĕϊ��A�擾����B
'           �ϊ��Ɏ��s�����ۂ̓f�t�H���g�l���Ԃ�܂��B
' �����P�@�F ���l������킷�I�u�W�F�N�g�BVariant�^�B
' �����Q�@�F �f�t�H���g�l�B�i�f�t�H���g�l�F-9999�j
' �Ԃ�l�@�F �ϊ����ꂽ�����l�B
' �g�p���@�F int_val = toInteger(val)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function toInteger(arg1 As Variant, Optional default As Integer = -9999) As Integer
    If IsNumeric(arg1) Then
        toInteger = CInt(arg1)
        Exit Function
    End If
    
    toInteger = default
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �I�u�W�F�N�g�𐔒l�iDouble�j�Ƃ��ĕϊ��A�擾����B
'           �ϊ��Ɏ��s�����ۂ̓f�t�H���g�l���Ԃ�܂��B
' �����P�@�F ���l������킷�I�u�W�F�N�g�BVariant�^�B
' �����Q�@�F �f�t�H���g�l�B�i�f�t�H���g�l�F-9999�j
' �Ԃ�l�@�F �ϊ����ꂽ�����l�B
' �g�p���@�F double_val = toDouble(val)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function toDouble(arg1 As Variant, Optional default As Double = -9999) As Double
    If IsNumeric(arg1) Then
        toDouble = CDbl(arg1)
        Exit Function
    End If
    
    toDouble = default
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �I�u�W�F�N�g�𐔒l�iDecimal�j�Ƃ��ĕϊ��A�擾����B
'           �ϊ��Ɏ��s�����ۂ̓f�t�H���g�l���Ԃ�܂��B
' �����P�@�F ���l������킷�I�u�W�F�N�g�BVariant�^�B
' �����Q�@�F �f�t�H���g�l�B�i�f�t�H���g�l�F-9999�j
' �Ԃ�l�@�F �ϊ����ꂽ�����l�B
' �g�p���@�F decimal_val = toDecimal(val)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function toDecimal(arg1 As Variant, Optional default As Variant = -9999) As Variant
    If IsNumeric(arg1) Then
        toDecimal = CDec(arg1)
        Exit Function
    End If
    
    toDecimal = default
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F ��������R���N�V�����ɕϊ�����B
' �����P�@�F ���X�g������B
' �����Q�@�F ��؂蕶���B�i�f�t�H���g�l�F","�j
' �Ԃ�l�@�F �ϊ����ꂽ�R���N�V�����B
' �g�p���@�F Dim col As Collection
'            Set col = String2Collection("data1 data2 data3", " ")
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function String2Collection(ByRef list_string As String, Optional delim As String = ",") As Collection
    Dim col As Collection
    Set col = New Collection
    
    ' ���X�g���������؂蕶���ŕ�������B
    Dim tmp As Variant
    tmp = Strings.Split(list_string, delim)
    ' ���������z����R���N�V�����ɕϊ�����B
    Set col = Array2Collection(tmp)
    
    Set String2Collection = col
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F Variant�^�̔z����R���N�V�����ɕϊ�����B
' �����P�@�F Variant�^�z��B
' �Ԃ�l�@�F �ϊ����ꂽ�R���N�V�����B
' �g�p���@�F Dim col As Collection
'            Set col = Array2Collection(variant_list)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Array2Collection(ByRef list As Variant) As Collection
    Dim col As Collection
    Set col = New Collection
    
    index = 0
    For index = LBound(list) To UBound(list)
        ' �z��̗v�f���R���N�V�����ɒǉ�����B
        col.add list(index)
    Next index
    
    Set Array2Collection = col
End Function

