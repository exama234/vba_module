Attribute VB_Name = "Mod_Excel"
''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 指定シートが参照可能かをチェックする。
' 引数１　： Excelブックの絶対パス。
' 引数２　： シート名。
' 返り値　： 参照可能な場合、真を返す。
' 使用方法： If CheckSheet("C:\test.xls", "Sheet1") Then
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CheckSheet(ByRef strBook As String, ByRef strSheet As String) As Boolean
    Dim sheetObj As Worksheet
    fileonly = Dir(strBook)

    On Error GoTo SheetError
    Set sheetObj = Workbooks(fileonly).Worksheets(strSheet)
    Set sheetObj = Nothing
    CheckSheet = True
    Exit Function

SheetError:
    CheckSheet = False
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 指定ブックが開いているかをチェックする。
' 引数１　： Excelブックの絶対パス。
' 返り値　： 開いている場合、真を返す。
' 使用方法： If openCheck("C:\test.xls") Then
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function openCheck(ByRef filefullpath As String) As Boolean
    Dim book As Workbook
    For Each book In Workbooks
        If book.Name = Dir(filefullpath) Then
            Set book = Nothing
            openCheck = True
            Exit Function
        End If
    Next book
    
    Set book = Nothing
    ' 開いてない
    openCheck = False
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： セルを選択中かを判定する。
' 引数１　： エラーメッセージ表示フラグ。（デフォルト値：True）
' 返り値　： セルを選択中かを表す真偽値。（True：セル選択中、False：セル選択中でない）
' 使用方法： If isSelectionRange() Then
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function isSelectionRange(Optional msg_flg As Boolean = True) As Boolean
    If TypeName(Selection) = "Range" Then
        ' セル選択中。真を返す。
        isSelectionRange = True
        Exit Function
    End If
    
    ' セル選択中でない。必要ならエラー表示する。
    If msg_flg Then
        MsgBox ("セルを選択してください。")
    End If
End Function




