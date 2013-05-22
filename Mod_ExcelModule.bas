Attribute VB_Name = "Mod_ExcelModule"
Public Enum overwrite
    yes = 0
    warning = 1
    no = 2
End Enum

Sub ModuleImport()
  '�v���W�F�N�g�̐錾
  Set Project = ActiveWorkbook.VBProject.VBComponents
 
  'Path���w��
  Dim pathlist As Variant
'  pathlist = OpenFileDialogMultiSelect()
        Dim file_list As Variant
    
        ' �u�t�@�C�����J���v�_�C�A���O���J���܂��B�i�����I���\�j
        file_list = Application.GetOpenFilename(, , , , True)
    
        ' �t�@�C�����I�����ꂽ�ۂ͕�����^�i�t�@�C�����j�A
        ' �L�����Z�����ꂽ�ۂ�Boolean�^�iFalse�j���Ԃ�܂��
        If VarType(file_list) <> vbBoolean Then
            ' ���X�g�i�I���t�@�C���̐�΃p�X�j��Ԃ��B
            pathlist = file_list
        End If
  If IsEmpty(pathlist) Then
    Exit Sub
  End If
  
  For Each v In pathlist
    ' Import�����B
    Project.Import v
  Next v
  
End Sub







Private Function getExportFolderPath(Optional lump_flag As String = True) As String
  Dim folder_fullpath As Variant
  If lump_flag Then
    ' �ꊇ
    folder_fullpath = FolderDialog()
    If folder_fullpath = "" Then
        ' �I��
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
  ' ���W���[�������ׂăG�N�X�|�[�g
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
        ' �t�@�C�������݂��Ȃ����߁A���̂܂܏o�͉\�B
        getExportFlg = True
        Exit Function
    End If
    
    
    ' ���Ƀt�@�C�������݂���B
    ' �o�͉\���͏㏑���t���O�ɂ���Ĕ��f����B
    Dim export_flg As Boolean
    export_flg = False
    Select Case over_write_flg
        Case overwrite.yes
              ' �����ɂď㏑���w�肳��Ă���B
            export_flg = True
        Case overwrite.warning
              ' �����ɂď㏑�����m�F����B
              answer1 = MsgBox("�t�@�C�������ɑ��݂��܂��B" & vbNewLine & "�㏑�����܂����B" & vbNewLine & vbTab & file_fullpath, vbYesNo)
              If answer1 = vbYes Then
                export_flg = True
              End If
        Case overwrite.no
              ' �����ɂď㏑�����Ȃ��B
            export_flg = False
    End Select
    
    getExportFlg = export_flg
End Function


Private Function getExportFilename(component As Object, Optional folder_path As String) As String
    Dim file_fullpath As Variant
    
    ' �ۑ��t�@�C�������擾����B
    Dim filename_only As String
    Select Case component.Type
        Case 1
            ' �W�����W���[��
            filename_only = component.Name & ".bas"
        Case 2
            ' �N���X���W���[��
            filename_only = component.Name & ".cls"
        Case 3
            ' ���[�U�[�t�H�[��
            filename_only = component.Name & ".frm"
        Case Else
            Exit Function
    End Select
    ' �ۑ��t�@�C�����i��΃p�X�j���擾����B
    If folder_path = "" Then
        ' �u�t�@�C���ۑ��v�_�C�A���O�ŕۑ�����w��
        file_fullpath = Mod_Dialog.SaveFileDialog(filename_only)
        If IsEmpty(file_fullpath) Then
            ' �L�����Z�����ꂽ�B
            Exit Function
        End If
    Else
        If Mod_File.isExistFolder(folder_path) = False Then
            ' �����w��̏o�͐�t�H���_�����݂��Ȃ��B
            Exit Function
        End If
        file_fullpath = folder_path & Application.PathSeparator & filename_only
    End If
    
    ' �o�̓t�@�C���̐�΃p�X��Ԃ��B
    getExportFilename = file_fullpath
End Function

Public Sub ModuleExportUnit(component As Object, Optional folder_path As String, Optional over_write_flg As overwrite = overwrite.warning)
    ' �ۑ��t�@�C�������擾����B
    Dim file_fullpath As String
    file_fullpath = getExportFilename(component, folder_path)
    If file_fullpath = "" Then
        Exit Sub
    End If
    
    Dim flg1 As Boolean
    flg1 = getExportFlg(file_fullpath, over_write_flg)
    If flg1 Then
        ' Export�����B
        component.Export file_fullpath
    End If
End Sub


