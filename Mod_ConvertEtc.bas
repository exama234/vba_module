Attribute VB_Name = "Mod_ConvertEtc"
''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： コレクションをMyList型の配列に変換する。
' 引数１　： コレクション。
' 返り値　： 変換されたMyList型の配列。
' 使用方法： Dim list As MyList
'            Set list = Collection2MyList(col)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Collection2MyList(col As Collection) As MyList
    Dim list As MyList
    Set list = New MyList
    
    Dim v As Variant
    For Each v In col
        ' コレクションの要素をMyList型に追加する。
        Call list.add(v)
    Next
    
    Set Collection2MyList = list
End Function


