Attribute VB_Name = "Mod_Dialog"
''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �u�t�@�C�����J���v�_�C�A���O��\������B
' �����P�@�F �t�@�C���g���q�̍i�荞�ݎw��B�i�f�t�H���g�l�F""�j
' �����Q�@�F �i�荞�݂̃C���f�b�N�X�B�i�f�t�H���g�l�F0�j
' �Ԃ�l�@�F �I�����ꂽ�t�@�C���̐�΃p�X�B
'            �_�C�A���O���L�����Z�����ꂽ�ۂ͋󕶎����Ԃ�܂��B
' �g�p���@�F file_fullpath = OpenFileDialog()
'            file_fullpath = OpenFileDialog("CSV �t�@�C�� (*.csv),*.csv,�e�L�X�g�t�@�C�� (*.txt),*.txt", 1)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OpenFileDialog(Optional filter As String = "", Optional idx As Integer = 0) As String
    Dim file_fullpath As Variant

    ' �u�t�@�C�����J���v�_�C�A���O���J���܂��B
    file_fullpath = Application.GetOpenFilename(filter, idx)

    ' �t�@�C�����I�����ꂽ�ۂ͕�����^�i�t�@�C����΃p�X�j�A
    ' �L�����Z�����ꂽ�ۂ�Boolean�^�iFalse�j���Ԃ�܂��
    If VarType(file_fullpath) <> vbBoolean Then
        ' �I���t�@�C���̐�΃p�X��Ԃ��B
        OpenFileDialog = file_fullpath
    End If
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �u�t�@�C�����J���v�_�C�A���O��\������B�i�����I���\�j
' �����P�@�F �t�@�C���g���q�̍i�荞�ݎw��B�i�f�t�H���g�l�F""�j
' �����Q�@�F �i�荞�݂̃C���f�b�N�X�B�i�f�t�H���g�l�F0�j
' �Ԃ�l�@�F �I�����ꂽ�t�@�C���̃��X�g�B
'            �_�C�A���O���L�����Z�����ꂽ�ۂ�Empty�l���Ԃ�܂��B
' �g�p���@�F list = OpenFileDialogMultiSelect()
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OpenFileDialogMultiSelect(Optional filter As String = "", Optional idx As Integer = 0) As Variant
    Dim file_list As Variant

    ' �u�t�@�C�����J���v�_�C�A���O���J���܂��B�i�����I���\�j
    file_list = Application.GetOpenFilename(filter, idx, , , True)

    ' �t�@�C�����I�����ꂽ�ۂ͕�����^�i�t�@�C�����j�A
    ' �L�����Z�����ꂽ�ۂ�Boolean�^�iFalse�j���Ԃ�܂��
    If VarType(file_list) <> vbBoolean Then
        ' ���X�g�i�I���t�@�C���̐�΃p�X�j��Ԃ��B
        OpenFileDialogMultiSelect = file_list
    End If
End Function



Public Function SaveFileDialog(Optional init_filename As String = "", Optional filter As String = "", Optional idx As Integer = 0) As Variant
    Dim file_fullpath As Variant

    ' �u�t�@�C�����J���v�_�C�A���O���J���܂��B
    file_fullpath = Application.GetSaveAsFilename(init_filename, filter, idx)

    ' �t�@�C�����I�����ꂽ�ۂ͕�����^�i�t�@�C����΃p�X�j�A
    ' �L�����Z�����ꂽ�ۂ�Boolean�^�iFalse�j���Ԃ�܂��
    If VarType(file_fullpath) <> vbBoolean Then
        ' �I���t�@�C���̐�΃p�X��Ԃ��B
        SaveFileDialog = file_fullpath
    End If
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �u�Q�Ɓv�_�C�A���O��\������B
' �Ԃ�l�@�F �I�����ꂽ�t�H���_�̐�΃p�X�B
'            �_�C�A���O���L�����Z�����ꂽ�ۂ͋󕶎����Ԃ�܂��B
' �g�p���@�F folder_fullpath = FolderDialog()
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FolderDialog() As String
    Dim dialog As Object
    
    ' �u�Q�Ɓv�_�C�A���O���J���܂��B
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    If dialog.Show = False Then
        ' �_�C�A���O�̃L�����Z���B
        Exit Function
    End If
    
    ' �I���t�H���_�̐�΃p�X��Ԃ��B
    folderFullpath = dialog.SelectedItems(1)
    FolderDialog = folderFullpath
End Function

