Attribute VB_Name = "Mod_ConvertEtc"
''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �R���N�V������MyList�^�̔z��ɕϊ�����B
' �����P�@�F �R���N�V�����B
' �Ԃ�l�@�F �ϊ����ꂽMyList�^�̔z��B
' �g�p���@�F Dim list As MyList
'            Set list = Collection2MyList(col)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Collection2MyList(col As Collection) As MyList
    Dim list As MyList
    Set list = New MyList
    
    Dim v As Variant
    For Each v In col
        ' �R���N�V�����̗v�f��MyList�^�ɒǉ�����B
        Call list.add(v)
    Next
    
    Set Collection2MyList = list
End Function


