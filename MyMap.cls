VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private map As Object

' コンストラクタ
Private Sub Class_Initialize()
    Set map = CreateObject("Scripting.Dictionary")
End Sub


Public Function isObj(key As String) As Boolean
    ' オブジェクト型か確認する。
    isObj = IsObject(map.item(key))
End Function


'''''''''''''''''''''''''''''''''''''
' Scripting.Dictionary プロパティ
'''''''''''''''''''''''''''''''''''''
Public Function count() As Integer
    ' オブジェクト型か確認する。
    count = map.count
End Function

Public Function item(key As String) As Variant
    ' キーに対応する値を取得する。
    If IsObject(map.item(key)) Then
        Set item = map.item(key)
    Else
        item = map.item(key)
    End If
End Function


'''''''''''''''''''''''''''''''''''''
' Scripting.Dictionary メソッド
'''''''''''''''''''''''''''''''''''''
Public Function add(key As String, val As Variant) As Variant
    ' 一度削除する。
    Call remove(key)
    ' 値を追加する。
    add = map.add(key, val)
End Function

Public Function exists(key As String) As Boolean
    ' キーの存在を確認する。
    exists = map.exists(key)
End Function


Public Function items() As Variant()
    ' オブジェクト型か確認する。
    items = map.items
End Function

Public Function keys() As Variant()
    ' オブジェクト型か確認する。
    keys = map.keys
End Function


Public Sub remove(key As String)
    If exists(key) Then
        ' 指定したキーと項目を削除します。
        Call map.remove(key)
    End If
End Sub

Public Sub RemoveAll()
    ' すべてのキーと項目を削除します。
    Call map.RemoveAll
End Sub


Public Function toString() As String
    '
    Dim key, buff, tmp, delim As String
    Dim item As Variant
    For Each key In map.keys
        item = ""
        tmp = ""
        If IsObject(map.item(key)) Then
            Set item = map.item(key)
        Else
            item = map.item(key)
        End If
        tmp = Mod_Class.toString(item)
        Debug.Print (tmp)
        tmp = key & "=" & tmp
        
        buff = buff & delim & tmp & ""
        delim = ", "
    Next

    toString = buff
End Function

