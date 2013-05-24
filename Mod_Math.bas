Attribute VB_Name = "Mod_Math"
''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： ２つの数値のうち、値の大きなものを取得する。
' 引数１　： 数値をあらわすオブジェクト。Variant型。
' 引数２　： 数値をあらわすオブジェクト。Variant型。
' 返り値　： 渡された引数で大きな値。
' 使用方法： val = max(300, 299.99)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function max(arg1 As Variant, arg2 As Variant) As Variant
	' テスト追加。
    If (arg1 > arg2) Then
        max = arg1
    Else
        max = arg2
    End If
End Function

