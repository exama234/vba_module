Attribute VB_Name = "Mod_ExcelModule"
Public Enum overwrite
    yes = 0
    warning = 1
    no = 2
End Enum

Sub ModuleImport()
  'プロジェクトの宣言
  Set Project = ActiveWorkbook.VBProject.VBComponents
 
  'Pathを指定
  Dim pathlist As Variant
'  pathlist = OpenFileDialogMultiSelect()
        Dim file_list As Variant
    
        ' 「ファイルを開く」ダイアログを開きます。（複数選択可能）
        file_list = Application.GetOpenFilename(, , , , True)
    
        ' ファイルが選択された際は文字列型（ファイル名）、
        ' キャンセルされた際はBoolean型（False）が返ります｡
        If VarType(file_list) <> vbBoolean Then
            ' リスト（選択ファイルの絶対パス）を返す。
            pathlist = file_list
        End If
  If IsEmpty(pathlist) Then
    Exit Sub
  End If
  
  For Each v In pathlist
    ' Import処理。
    Project.Import v
  Next v
  
End Sub







Private Function getExportFolderPath(Optional lump_flag As String = True) As String
  Dim folder_fullpath As Variant
  If lump_flag Then
    ' 一括
    folder_fullpath = FolderDialog()
    If folder_fullpath = "" Then
        ' 終了
        End
    End If
  End If

    getExportFolderPath = folder_fullpath
End Function

Public Sub ModuleExportAll(Optional lump_flag As String = True, Optional over_write_flg As overwrite = overwrite.warning)
    Dim folder_path As String
    folder_path = getExportFolderPath(lump_flag)
    
    Call ModuleExport(folder_path, 0, over_write_flg)
End Sub
Public Sub ModuleExportModule(Optional lump_flag As String = True, Optional over_write_flg As overwrite = overwrite.warning)
    Dim folder_path As String
    folder_path = getExportFolderPath(lump_flag)
    
    Call ModuleExport(folder_path, 1, over_write_flg)
End Sub
Public Sub ModuleExportClass(Optional lump_flag As String = True, Optional over_write_flg As overwrite = overwrite.warning)
    Dim folder_path As String
    folder_path = getExportFolderPath(lump_flag)
    
    Call ModuleExport(folder_path, 2, over_write_flg)
End Sub
Public Sub ModuleExportUserForm(Optional lump_flag As String = True, Optional over_write_flg As overwrite = overwrite.warning)
    Dim folder_path As String
    folder_path = getExportFolderPath(lump_flag)
    
    Call ModuleExport(folder_path, 3, over_write_flg)
End Sub
Public Sub ModuleExport(folder_path As String, Optional module_type As Integer = 0, Optional over_write_flg As overwrite = overwrite.warning)
  ' モジュールをすべてエクスポート
  Set ComponentList = ActiveWorkbook.VBProject.VBComponents
  Dim component As Object
  For Each component In ComponentList
    If module_type = 0 Or component.Type = module_type Then
        Call ModuleExportUnit(component, CStr(folder_path), over_write_flg)
    End If
  Next
End Sub

Private Function getExportFlg(file_fullpath As String, Optional over_write_flg As overwrite = overwrite.warning) As Boolean
    If Mod_File.isExistFile(file_fullpath) = False Then
        ' ファイルが存在しないため、そのまま出力可能。
        getExportFlg = True
        Exit Function
    End If
    
    
    ' 既にファイルが存在する。
    ' 出力可能かは上書きフラグによって判断する。
    Dim export_flg As Boolean
    export_flg = False
    Select Case over_write_flg
        Case overwrite.yes
              ' 引数にて上書き指定されている。
            export_flg = True
        Case overwrite.warning
              ' 引数にて上書きを確認する。
              answer1 = MsgBox("ファイルが既に存在します。" & vbNewLine & "上書きしますか。" & vbNewLine & vbTab & file_fullpath, vbYesNo)
              If answer1 = vbYes Then
                export_flg = True
              End If
        Case overwrite.no
              ' 引数にて上書きしない。
            export_flg = False
    End Select
    
    getExportFlg = export_flg
End Function


Private Function getExportFilename(component As Object, Optional folder_path As String) As String
    Dim file_fullpath As Variant
    
    ' 保存ファイル名を取得する。
    Dim filename_only As String
    Select Case component.Type
        Case 1
            ' 標準モジュール
            filename_only = component.Name & ".bas"
        Case 2
            ' クラスモジュール
            filename_only = component.Name & ".cls"
        Case 3
            ' ユーザーフォーム
            filename_only = component.Name & ".frm"
        Case Else
            Exit Function
    End Select
    ' 保存ファイル名（絶対パス）を取得する。
    If folder_path = "" Then
        ' 「ファイル保存」ダイアログで保存先を指定
        file_fullpath = Mod_Dialog.SaveFileDialog(filename_only)
        If IsEmpty(file_fullpath) Then
            ' キャンセルされた。
            Exit Function
        End If
    Else
        If Mod_File.isExistFolder(folder_path) = False Then
            ' 引数指定の出力先フォルダが存在しない。
            Exit Function
        End If
        file_fullpath = folder_path & Application.PathSeparator & filename_only
    End If
    
    ' 出力ファイルの絶対パスを返す。
    getExportFilename = file_fullpath
End Function

Public Sub ModuleExportUnit(component As Object, Optional folder_path As String, Optional over_write_flg As overwrite = overwrite.warning)
    ' 保存ファイル名を取得する。
    Dim file_fullpath As String
    file_fullpath = getExportFilename(component, folder_path)
    If file_fullpath = "" Then
        Exit Sub
    End If
    
    Dim flg1 As Boolean
    flg1 = getExportFlg(file_fullpath, over_write_flg)
    If flg1 Then
        ' Export処理。
        component.Export file_fullpath
    End If
End Sub


