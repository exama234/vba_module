Attribute VB_Name = "Mod_File"
''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 絶対パスのファイル名からドライブ名のみを取得する。
' 引数１　： ファイルの絶対パス。
' 返り値　： ドライブ名のみを返す。
' 使用方法： drive_name = getDrive("C:\folder1\file1.txt")
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function getDrive(ByVal filefullpath As String) As String
    Dim FSO
    Dim drive As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    drive = FSO.getDriveName(filefullpath)
    Set FSO = Nothing
    
    getDrive = drive
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 絶対パスのファイル名からパスのみを取得する。
' 引数１　： ファイルの絶対パス。
' 返り値　： パスを返す。
' 使用方法： path = getPath("C:\folder1\file1.txt")
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
' 説明　　： 絶対パスのファイル名からファイル名のみを取得する。
' 引数１　： ファイルの絶対パス。
' 返り値　： ファイル名のみを返す。
' 使用方法： fileonly = getFilename("C:\folder1\file1.txt")
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
' 説明　　： 絶対パスのファイル名からベース名のみを取得する。
' 引数１　： ファイルの絶対パス。
' 返り値　： ベース名のみを返す。
' 使用方法： basename = getBasename("C:\folder1\file1.txt")
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function getBasename(ByVal filefullpath As String) As String
    Dim FSO
    Dim basename As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    basename = FSO.getBasename(filefullpath)
    Set FSO = Nothing
    
    getBasename = basename
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 絶対パスのファイル名から拡張子のみを取得する。
' 引数１　： ファイルの絶対パス。
' 返り値　： 拡張子のみを返す。
' 使用方法： ext = getExt("C:\folder1\file1.txt")
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function getExt(ByVal filefullpath As String) As String
    Dim FSO
    Dim ext As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    ext = FSO.GetExtensionName(filefullpath)
    Set FSO = Nothing
    
    getExt = ext
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 指定ファイルが存在するかをチェックする。
' 引数１　： ファイルの絶対パス。
' 返り値　： 存在結果を返す。（真：存在する／偽：存在しない）
' 使用方法： If isExistFile("C:\file1.txt") Then
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function isExistFile(ByRef filefullpath As String) As Boolean
    ' FileSystemObjectを取得する。
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If FSO.FileExists(filefullpath) Then
        ' ファイルが存在する。
        Set FSO = Nothing
        isExistFile = True
        Exit Function
    End If
    
    ' ファイルが存在しない。
    Set FSO = Nothing
    isExistFile = False
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 指定フォルダが存在するかをチェックする。
' 引数１　： フォルダの絶対パス。
' 返り値　： 存在結果を返す。（真：存在するフォルダ偽：存在しない）
' 使用方法： If isExistFolder("C:\folder1") Then
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function isExistFolder(ByRef filefullpath As String) As Boolean
    ' FileSystemObjectを取得する。
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If FSO.FolderExists(filefullpath) Then
        ' フォルダが存在する。
        Set FSO = Nothing
        isExistFolder = True
        Exit Function
    End If
    
    ' フォルダが存在しない。
    Set FSO = Nothing
    isExistFolder = False
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 指定フォルダからファイルリストを取得する。
' 引数１　： フォルダの絶対パス。
' 引数２　： 取得ファイルの正規表現文字列。
' 返り値　： ファイルのコレクションを返す。
' 使用方法： Set file_col = getFileList(path)
''''''''''''''''''''''''''''''''''''''''''''''''''
Function getFileList(folder_path As String, Optional regex_str As String = ".*") As Collection
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
    
    ' ファイルリスト取得。
    Dim FileList As Object
    Set FileList = FSO.GetFolder(folder_path).Files
    For Each tmp In FileList
        If regex.Test(tmp.Name) Then
            col.Add tmp.path
        End If
    Next
    
    ' サブフォルダリスト取得。
    Dim folderList As Object
    Set folderList = FSO.GetFolder(folder_path).SubFolders
    For Each tmp In folderList
        Set col2 = getFileList(folder_path & "\" & tmp.Name, regex_str)
        For Each tmp2 In col2
            col.Add tmp2
        Next
    Next
    
    Set getFileList = col
End Function
