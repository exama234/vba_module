Attribute VB_Name = "Mod_Class"
''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： オブジェクトの文字列表現を取得する。
' 引数１　： 対象となるオブジェクト。Variant型。
' 返り値　： 渡されたオブジェクトの文字列表現。
' 使用方法： str = toString(obj)
'            toString(18)    <= 数値型の場合、文字列「18」が返る。
'            toString("21")  <= 文字列の場合、文字列「"21"」が返る。
'            toString(True)  <= 真偽値の場合、文字列「True」が返る。
'            toString(today) <= 日付型の場合、文字列「#2012/01/04#」が返る。
'            toString(array) <= 配列型の場合、文字列「[18, "21", #2012/01/04#]」が返る。
'            toString(obj)   <= "Nothing"の場合、文字列「(Nothing)」が返る。
'            toString(obj)   <= "Empty"の場合、文字列「(Empty)」が返る。
'            toString(obj)   <= "Null"の場合、文字列「(Null)」が返る。
'            toString(obj)   <= "Error"の場合、文字列「(Error)」が返る。
'            toString(obj)   <= オブジェクトの場合、文字列「クラス名[...]」が返る。
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function toString(item As Variant) As String
    '
    Dim key, buff, tmp, delim As String
    ' select文
    Select Case TypeName(item)
        Case "Byte", "Integer", "Long", "Boolean", "Single", "Double", "Currency"
            tmp = item
        Case "String"
            tmp = """" & item & """"
        Case "Date"
            tmp = "#" & item & "#"
        Case "Nothing", "Empty", "Null", "Error"
            tmp = "(" & TypeName(item) & ")"
        Case Else
            If IsArray(item) Then
                Dim v As Variant
                For idx = LBound(item) To UBound(item)
                    If IsObject(item(idx)) Then
                        Set v = item(idx)
                    Else
                        v = item(idx)
                    End If
                    
                    tmp = tmp & delim & toString(v)
                    delim = ", "
                Next idx
                tmp = "[" & tmp & "]"
            Else
                If Mod_Class.hasMethod(item, "toString") Then
                    ' toString()があるなら実行したい。
                    tmp = CallByName(item, "toString", VbMethod)
                    tmp = TypeName(item) & "[" & tmp & "]"
                Else
                    tmp = TypeName(item) & "[...]"
                End If
            End If
    End Select
    
    buff = tmp
    toString = buff
End Function




''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： オブジェクトがメソッドを持っているかを確認する。
' 引数１　： 確認対象のオブジェクト。Variant型。
' 引数２　： 確認するメソッドの名称。
' 返り値　： メソッドを持つ場合、真を返す。
' 使用方法： If hasMethod(obj, "toString") Then
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function hasMethod(obj As Variant, method As String) As Boolean
On Error GoTo ErrHandler
    ' オブジェクトに指定のメソッドがあるなら実行する。
    Call CallByName(obj, method, VbMethod)
    ' 例外が発生しなかった為、成功。
    hasMethod = True
    Exit Function
    
ErrHandler:
    ' 例外発生。
    hasMethod = False
End Function

