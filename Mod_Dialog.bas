Attribute VB_Name = "Mod_Dialog"
''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 「ファイルを開く」ダイアログを表示する。
' 引数１　： ファイル拡張子の絞り込み指定。（デフォルト値：""）
' 引数２　： 絞り込みのインデックス。（デフォルト値：0）
' 返り値　： 選択されたファイルの絶対パス。
'            ダイアログがキャンセルされた際は空文字が返ります。
' 使用方法： file_fullpath = OpenFileDialog()
'            file_fullpath = OpenFileDialog("CSV ファイル (*.csv),*.csv,テキストファイル (*.txt),*.txt", 1)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OpenFileDialog(Optional filter As String = "", Optional idx As Integer = 0) As String
    Dim file_fullpath As Variant

    ' 「ファイルを開く」ダイアログを開きます。
    file_fullpath = Application.GetOpenFilename(filter, idx)

    ' ファイルが選択された際は文字列型（ファイル絶対パス）、
    ' キャンセルされた際はBoolean型（False）が返ります｡
    If VarType(file_fullpath) <> vbBoolean Then
        ' 選択ファイルの絶対パスを返す。
        OpenFileDialog = file_fullpath
    End If
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 「ファイルを開く」ダイアログを表示する。（複数選択可能）
' 引数１　： ファイル拡張子の絞り込み指定。（デフォルト値：""）
' 引数２　： 絞り込みのインデックス。（デフォルト値：0）
' 返り値　： 選択されたファイルのリスト。
'            ダイアログがキャンセルされた際はEmpty値が返ります。
' 使用方法： list = OpenFileDialogMultiSelect()
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OpenFileDialogMultiSelect(Optional filter As String = "", Optional idx As Integer = 0) As Variant
    Dim file_list As Variant

    ' 「ファイルを開く」ダイアログを開きます。（複数選択可能）
    file_list = Application.GetOpenFilename(filter, idx, , , True)

    ' ファイルが選択された際は文字列型（ファイル名）、
    ' キャンセルされた際はBoolean型（False）が返ります｡
    If VarType(file_list) <> vbBoolean Then
        ' リスト（選択ファイルの絶対パス）を返す。
        OpenFileDialogMultiSelect = file_list
    End If
End Function



Public Function SaveFileDialog(Optional init_filename As String = "", Optional filter As String = "", Optional idx As Integer = 0) As Variant
    Dim file_fullpath As Variant

    ' 「ファイルを開く」ダイアログを開きます。
    file_fullpath = Application.GetSaveAsFilename(init_filename, filter, idx)

    ' ファイルが選択された際は文字列型（ファイル絶対パス）、
    ' キャンセルされた際はBoolean型（False）が返ります｡
    If VarType(file_fullpath) <> vbBoolean Then
        ' 選択ファイルの絶対パスを返す。
        SaveFileDialog = file_fullpath
    End If
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 「参照」ダイアログを表示する。
' 返り値　： 選択されたフォルダの絶対パス。
'            ダイアログがキャンセルされた際は空文字が返ります。
' 使用方法： folder_fullpath = FolderDialog()
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FolderDialog() As String
    Dim dialog As Object
    
    ' 「参照」ダイアログを開きます。
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    If dialog.Show = False Then
        ' ダイアログのキャンセル。
        Exit Function
    End If
    
    ' 選択フォルダの絶対パスを返す。
    folderFullpath = dialog.SelectedItems(1)
    FolderDialog = folderFullpath
End Function

