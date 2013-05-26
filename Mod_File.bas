Attribute VB_Name = "Mod_File"
''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F ��΃p�X�̃t�@�C��������p�X�݂̂��擾����B
' �����P�@�F �t�@�C���̐�΃p�X�B
' �Ԃ�l�@�F �p�X��Ԃ��B
' �g�p���@�F path = getPath("C:\folder1\file1.txt")
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function getPath(ByVal filefullpath As String) As String
    Dim FSO
    Dim path As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    path = FSO.GetParentFolderName(filefullpath)
    Set FSO = Nothing
    
    getPath = path
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F ��΃p�X�̃t�@�C��������t�@�C�����݂̂��擾����B
' �����P�@�F �t�@�C���̐�΃p�X�B
' �Ԃ�l�@�F �t�@�C�����݂̂�Ԃ��B
' �g�p���@�F fileonly = getFilename("C:\folder1\file1.txt")
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function getFilename(ByVal filefullpath As String) As String
    Dim FSO
    Dim filename As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    filename = FSO.getFilename(filefullpath)
    Set FSO = Nothing
    
    getFilename = filename
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �w��t�@�C�������݂��邩���`�F�b�N����B
' �����P�@�F �t�@�C���̐�΃p�X�B
' �Ԃ�l�@�F ���݌��ʂ�Ԃ��B�i�^�F���݂���^�U�F���݂��Ȃ��j
' �g�p���@�F If isExistFile("C:\file1.txt") Then
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function isExistFile(ByRef filefullpath As String) As Boolean
    ' FileSystemObject���擾����B
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If FSO.FileExists(filefullpath) Then
        ' �t�@�C�������݂���B
        Set FSO = Nothing
        isExistFile = True
        Exit Function
    End If
    
    ' �t�@�C�������݂��Ȃ��B
    Set FSO = Nothing
    isExistFile = False
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �w��t�H���_�����݂��邩���`�F�b�N����B
' �����P�@�F �t�H���_�̐�΃p�X�B
' �Ԃ�l�@�F ���݌��ʂ�Ԃ��B�i�^�F���݂���t�H���_�U�F���݂��Ȃ��j
' �g�p���@�F If isExistFolder("C:\folder1") Then
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function isExistFolder(ByRef filefullpath As String) As Boolean
    ' FileSystemObject���擾����B
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If FSO.FolderExists(filefullpath) Then
        ' �t�H���_�����݂���B
        Set FSO = Nothing
        isExistFolder = True
        Exit Function
    End If
    
    ' �t�H���_�����݂��Ȃ��B
    Set FSO = Nothing
    isExistFolder = False
End Function



Function getFileList(folder_path As String, Optional regex_str As String = "*") As Collection
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = regex_str
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")

    Dim col As Collection
    Set col = New Collection
    Dim col2 As Collection
    
    Dim file_list As Variant
    Dim Folder_List As Variant
    
    ' �t�@�C�����X�g�擾�B
    Dim FileList As Object
    Set FileList = FSO.GetFolder(folder_path).Files
    For Each tmp In FileList
        If regex.test(tmp.name) Then
            col.add tmp.path
        End If
    Next
    
    ' �T�u�t�H���_���X�g�擾�B
    Dim folderList As Object
    Set folderList = FSO.GetFolder(folder_path).SubFolders
    For Each tmp In folderList
        Set col2 = getFileList(folder_path & "\" & tmp.name, regex_str)
        For Each tmp2 In col2
            col.add tmp2
        Next
    Next
    
    Set getFileList = col
End Function
