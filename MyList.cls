VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private col As Collection

' �R���X�g���N�^
Private Sub Class_Initialize()
    Set col = New Collection
End Sub

Public Function isObj(index As Integer) As Boolean
    ' �I�u�W�F�N�g�^���m�F����B
    isObj = IsObject(col.item(index))
End Function


'''''''''''''''''''''''''''''''''''''
' Collection �v���p�e�B
'''''''''''''''''''''''''''''''''''''
Public Function count() As Integer
    ' �I�u�W�F�N�g�^���m�F����B
    count = col.count
End Function

Public Function item(index As Integer) As Variant
    ' �L�[�ɑΉ�����l���擾����B
    If IsObject(col.item(index)) Then
        Set item = col.item(index)
    Else
        item = col.item(index)
    End If
End Function


'''''''''''''''''''''''''''''''''''''
' Collection ���\�b�h
'''''''''''''''''''''''''''''''''''''
Public Function add(val As Variant)
    ' �l��ǉ�����B
    ' ��2�Ԗڂ̈����ikey�j���Ȃ��ƃI�u�W�F�N�g���ǉ��ł��Ȃ��B
    ' �d���Ȃ��ׁA�Ƃ肠���������^�̃C���f�b�N�X���w�肷��B
    Call col.add(val, CStr(col.count + 1))
End Function



Public Function exists(obj As Variant) As Boolean
    Dim item As Variant
    ' �l�̑��݂��m�F����B
    For Each item In col
        If TypeName(item) = TypeName(obj) Then
            If IsObject(item) Then
                If item Is obj Then
                    exists = True
                    Exit Function
                End If
            Else
                If item = obj Then
                    exists = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function



Public Function items() As Collection
    ' �I�u�W�F�N�g�^���m�F����B
    Set items = col
End Function



Public Function remove(index As Integer) As Boolean
    If index < 1 Or col.count < index Then
        Exit Function
    End If
    
    ' �w��C���f�b�N�X�̍��ڂ��폜���܂��B
    Call col.remove(index)
    remove = True
End Function



Public Sub RemoveAll()
    ' ���ׂĂ̍��ڂ��폜���܂��B
    Set col = Nothing
    Set col = New Collection
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �I�u�W�F�N�g�̕�����\�����擾����B
' �Ԃ�l�@�F �I�u�W�F�N�g�̕�����\���B
' �g�p���@�F MsgBox list.toString()
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function toString() As String
    '
    Dim key, buff, tmp, delim As String
    Dim item As Variant
    For Each item In col
        tmp = ""
        tmp = Mod_Class.toString(item)
        Debug.Print (tmp)
        
        buff = buff & delim & tmp & ""
        delim = ", "
    Next

    toString = buff
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F ���X�g�����������ɍ��������X�g�Ƀt�B���^�����O����B
' �����P�@�F ���K�\���B
' �Ԃ�l�@�F �t�B���^�����O���ꂽ�R���N�V�����B
' �g�p���@�F Set list = list.filltering_regex("^.*\.txt$")
''''''''''''''''''''''''''''''''''''''''''''''''''
Function filltering_regex(regex_str As String) As MyList
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = regex_str
    
    ' �V�K���X�g���쐬���A�����ɍ����f�[�^�݂̂̃��X�g��Ԃ��B
    Dim new_list As MyList
    Set new_list = New MyList
    
    For Each tmp In col
        If regex.test(tmp) Then
            ' �����w��̏����ɍ����ׁA�ǉ�����B
            new_list.add tmp
        End If
    Next
    
    Set filltering_regex = new_list
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F ���X�g�����������ɍ��������X�g�Ƀt�B���^�����O����B
' �����P�@�F ���Z�q�B�ȉ����g�p�\�ł��B
'         �F =    <>    <    <=   >    >=
' �����Q�@�F ��r���鐔�l�B
' �Ԃ�l�@�F �t�B���^�����O���ꂽ�R���N�V�����B
' �g�p���@�F Set list = list.filltering_numeric("=", 100)
'         �F Set list = list.filltering_numeric("<=", 100)
''''''''''''''''''''''''''''''''''''''''''''''''''
Function filltering_numeric(operand As String, number As Variant) As MyList
    If IsNumeric(nunber) = False Then
        ' ���������l�łȂ��B
        Exit Function
    End If
    
    ' �V�K���X�g���쐬���A�����ɍ����f�[�^�݂̂̃��X�g��Ԃ��B
    Dim new_list As MyList
    Set new_list = New MyList
    
    Dim add_flg As Boolean
    For Each tmp In col
        add_flg = False
        ' ���Z�q���m�F���āA�����ɍ����ꍇ�͒ǉ��t���O��^�ɂ���B
        Select Case Strings.Trim(operand)
            Case "="
                If tmp = number Then
                    add_flg = True
                End If
                
            Case "<>"
                If tmp <> number Then
                    add_flg = True
                End If
                
            Case "<"
                If tmp < number Then
                    add_flg = True
                End If
                
            Case "<="
                If tmp <= number Then
                    add_flg = True
                End If
                
            Case ">="
                If tmp >= number Then
                    add_flg = True
                End If
                
            Case ">"
                If tmp > number Then
                    add_flg = True
                End If
                
            Case Else
                
        End Select

        If add_flg Then
            ' �����w��̏����ɍ����ׁA�ǉ�����B
            new_list.add tmp
        End If
    Next
    
    Set filltering_numeric = new_list
End Function