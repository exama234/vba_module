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

' コンストラクタ
Private Sub Class_Initialize()
    Set col = New Collection
End Sub

Public Function isObj(index As Integer) As Boolean
    ' オブジェクト型か確認する。
    isObj = IsObject(col.item(index))
End Function


'''''''''''''''''''''''''''''''''''''
' Collection プロパティ
'''''''''''''''''''''''''''''''''''''
Public Function count() As Integer
    ' オブジェクト型か確認する。
    count = col.count
End Function

Public Function item(index As Integer) As Variant
    ' キーに対応する値を取得する。
    If IsObject(col.item(index)) Then
        Set item = col.item(index)
    Else
        item = col.item(index)
    End If
End Function


'''''''''''''''''''''''''''''''''''''
' Collection メソッド
'''''''''''''''''''''''''''''''''''''
Public Function add(val As Variant)
    ' 値を追加する。
    ' ※2番目の引数（key）がないとオブジェクトが追加できない。
    ' 仕方ない為、とりあえず文字型のインデックスを指定する。
    Call col.add(val, CStr(col.count + 1))
End Function



Public Function exists(obj As Variant) As Boolean
    Dim item As Variant
    ' 値の存在を確認する。
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
    ' オブジェクト型か確認する。
    Set items = col
End Function



Public Function remove(index As Integer) As Boolean
    If index < 1 Or col.count < index Then
        Exit Function
    End If
    
    ' 指定インデックスの項目を削除します。
    Call col.remove(index)
    remove = True
End Function



Public Sub RemoveAll()
    ' すべての項目を削除します。
    Set col = Nothing
    Set col = New Collection
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： オブジェクトの文字列表現を取得する。
' 返り値　： オブジェクトの文字列表現。
' 使用方法： MsgBox list.toString()
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
' 説明　　： リストを引数条件に合ったリストにフィルタリングする。
' 引数１　： 正規表現。
' 返り値　： フィルタリングされたコレクション。
' 使用方法： Set list = list.filltering_regex("^.*\.txt$")
''''''''''''''''''''''''''''''''''''''''''''''''''
Function filltering_regex(regex_str As String) As MyList
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = regex_str
    
    ' 新規リストを作成し、条件に合うデータのみのリストを返す。
    Dim new_list As MyList
    Set new_list = New MyList
    
    For Each tmp In col
        If regex.test(tmp) Then
            ' 引数指定の条件に合う為、追加する。
            new_list.add tmp
        End If
    Next
    
    Set filltering_regex = new_list
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： リストを引数条件に合ったリストにフィルタリングする。
' 引数１　： 演算子。以下が使用可能です。
'         ： =    <>    <    <=   >    >=
' 引数２　： 比較する数値。
' 返り値　： フィルタリングされたコレクション。
' 使用方法： Set list = list.filltering_numeric("=", 100)
'         ： Set list = list.filltering_numeric("<=", 100)
''''''''''''''''''''''''''''''''''''''''''''''''''
Function filltering_numeric(operand As String, number As Variant) As MyList
    If IsNumeric(nunber) = False Then
        ' 引数が数値でない。
        Exit Function
    End If
    
    ' 新規リストを作成し、条件に合うデータのみのリストを返す。
    Dim new_list As MyList
    Set new_list = New MyList
    
    Dim add_flg As Boolean
    For Each tmp In col
        add_flg = False
        ' 演算子を確認して、条件に合う場合は追加フラグを真にする。
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
            ' 引数指定の条件に合う為、追加する。
            new_list.add tmp
        End If
    Next
    
    Set filltering_numeric = new_list
End Function
