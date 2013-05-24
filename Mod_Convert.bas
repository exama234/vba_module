Attribute VB_Name = "Mod_Convert"
''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： オブジェクトを数値（Integer）として変換、取得する。
'           変換に失敗した際はデフォルト値が返ります。
' 引数１　： 数値をあらわすオブジェクト。Variant型。
' 引数２　： デフォルト値。（デフォルト値：-9999）
' 返り値　： 変換された整数値。
' 使用方法： int_val = toInteger(val)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function toInteger(arg1 As Variant, Optional default As Integer = -9999) As Integer
    If IsNumeric(arg1) Then
        toInteger = CInt(arg1)
        Exit Function
    End If
    
    toInteger = default
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： オブジェクトを数値（Double）として変換、取得する。
'           変換に失敗した際はデフォルト値が返ります。
' 引数１　： 数値をあらわすオブジェクト。Variant型。
' 引数２　： デフォルト値。（デフォルト値：-9999）
' 返り値　： 変換された整数値。
' 使用方法： double_val = toDouble(val)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function toDouble(arg1 As Variant, Optional default As Double = -9999) As Double
    If IsNumeric(arg1) Then
        toDouble = CDbl(arg1)
        Exit Function
    End If
    
    toDouble = default
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： オブジェクトを数値（Decimal）として変換、取得する。
'           変換に失敗した際はデフォルト値が返ります。
' 引数１　： 数値をあらわすオブジェクト。Variant型。
' 引数２　： デフォルト値。（デフォルト値：-9999）
' 返り値　： 変換された整数値。
' 使用方法： decimal_val = toDecimal(val)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function toDecimal(arg1 As Variant, Optional default As Variant = -9999) As Variant
    If IsNumeric(arg1) Then
        toDecimal = CDec(arg1)
        Exit Function
    End If
    
    toDecimal = default
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 文字列をコレクションに変換する。
' 引数１　： リスト文字列。
' 引数２　： 区切り文字。（デフォルト値：","）
' 返り値　： 変換されたコレクション。
' 使用方法： Dim col As Collection
'            Set col = String2Collection("data1 data2 data3", " ")
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function String2Collection(ByRef list_string As String, Optional delim As String = ",") As Collection
    Dim col As Collection
    Set col = New Collection
    
    ' リスト文字列を区切り文字で分割する。
    Dim tmp As Variant
    tmp = Strings.Split(list_string, delim)
    ' 分割した配列をコレクションに変換する。
    Set col = Array2Collection(tmp)
    
    Set String2Collection = col
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： Variant型の配列をコレクションに変換する。
' 引数１　： Variant型配列。
' 返り値　： 変換されたコレクション。
' 使用方法： Dim col As Collection
'            Set col = Array2Collection(variant_list)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Array2Collection(ByRef list As Variant) As Collection
    Dim col As Collection
    Set col = New Collection
    
    index = 0
    For index = LBound(list) To UBound(list)
        ' 配列の要素をコレクションに追加する。
        col.add list(index)
    Next index
    
    Set Array2Collection = col
End Function

