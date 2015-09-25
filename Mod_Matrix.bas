Attribute VB_Name = "Mod_Matrix"
''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 多次元配列の次元数を取得する。
' 　　　　： array2[9][9] -> 2
' 　　　　： array3[10][20][30] -> 3
' 引数１　： 多次元配列。
' 返り値　： 多次元配列の次元数を返す。
' 使用方法： getDimension(xyArray)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function getDimension(matrix As Variant) As Integer
    ' 次元数を求める。
    getDimension = Application.DimSize(matrix)
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 多次元配列の指定次元の要素数を取得する。
' 　　　　： array2[9][9], 1 -> 9
' 　　　　： array3[10][20][30], 1 -> 10
' 引数１　： 多次元配列。
' 引数２　： 次元。
' 返り値　： 指定次元の要素数を返す。
' 使用方法： getSize(xyArray, 1)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function getSize(matrix As Variant, dimension As Integer) As Integer
    ' 指定次元の要素数を求める。
    getSize = Application.RBound(matrix, dimension)
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： ２次元配列の軸を入れ換える。
' 引数１　： ２次元配列。
' 返り値　： 軸を入換済みの配列を返す。
' 使用方法： xyArray = swapXYAxis(xyArray)
' 　　　　：     A1  B1  C1      A1  A2  A3
' 　　　　：     A2  B2  C2  ->  B1  B2  B3
' 　　　　：     A3  B3  C3      C1  C2  C3
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function swapXYAxis(matrix As Variant) As Variant
    ' XY軸を入れ替える。
    swapXYAxis = Application.Transpose(matrix)
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： ２次元配列の行範囲を狭める。
' 引数１　： ２次元配列。
' 引数２　： 開始インデックス。
' 引数３　： 配列長。
' 返り値　： 行範囲を狭めた配列を返す。
' 使用方法： xyArray = narrowXYRow(xyArray, 2, 1)
' 　　　　：     A1  B1  C1      A2  B2  C2
' 　　　　：     A2  B2  C2  ->
' 　　　　：     A3  B3  C3
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function narrowXYRow(matrix As Variant, start_idx As Long, Optional length As Long = -1) As Variant
    ' ループの最終Indexを取得する。
    Dim loopEndIndex As Long
    If length < 0 Then
        ' 長さ指定なし。
        loopEndIndex = UBound(matrix)
    Else
        ' 長さ指定あり。
        loopEndIndex = start_idx + length - 1
        If UBound(matrix) < loopEndIndex Then
            loopEndIndex = UBound(matrix)
        End If
    End If
    
    
    Dim newMatrix As Variant
    ReDim newMatrix(1 To loopEndIndex - start_idx + 1) As Variant
    Dim newIndex As Integer
    newIndex = 1
    For idx = start_idx To loopEndIndex
        newMatrix(newIndex) = Application.Index(matrix, idx)
        newIndex = newIndex + 1
    Next
    
    narrowXYRow = newMatrix
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： ２次元配列の列範囲を狭める。
' 引数１　： ２次元配列。
' 引数２　： 開始インデックス。
' 引数３　： 配列長。
' 返り値　： 列範囲を狭めた配列を返す。
' 使用方法： xyArray = narrowXYColumn(xyArray, 1, 2)
' 　　　　：     A1  B1  C1      A1  B2
' 　　　　：     A2  B2  C2  ->  A2  B2
' 　　　　：     A3  B3  C3      A3  B3
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function narrowXYColumn(ByVal matrix As Variant, start_idx As Long, Optional length As Long = -1) As Variant
    matrix = Application.Transpose(matrix)
    matrix = narrowXYRow(matrix, start_idx, length)
    matrix = Application.Transpose(matrix)
    
    narrowXYColumn = matrix
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： ２次元配列をカテゴリで分割する。
' 引数１　： ２次元配列。
' 引数２　： カテゴリ列インデックス文字列。
' 返り値　： カテゴリ（列の同一データ）で分割したマトリクス配列を返す。
' 使用方法： xyArray = getMatrixDivideRow(xyArray, "1,2")
' 　　　　：     A1  B1  C1      A1  B1 C1  xyArray(1)
' 　　　　：     A1  B1  C2  ->  A1  B1 C2
' 　　　　：     A1  B2  C1      ---------
' 　　　　：                     A1  B2 C1  xyArray(2)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function getMatrixDivideRow(matrix As Variant, categoryIndexString As String) As Variant()
    Dim col As Collection
    Set col = getCategoryIndexList(matrix, categoryIndexString)

    Dim return_value() As Variant
    ReDim return_value(1 To col.Count) As Variant
    Dim newIndex, length, column_size As Long
    For idx1 = 1 To col.Count
        Dim start_idx, end_idx As Long
        start_idx = col.Item(idx1)
        If idx1 = col.Count Then
            end_idx = UBound(matrix, 1)
        Else
            end_idx = col.Item(idx1 + 1) - 1
        End If
        
        newIndex = 1
        length = end_idx - start_idx + 1
        column_size = UBound(matrix, 2)
        Dim newMatrix() As Variant
        ReDim newMatrix(1 To length, 1 To column_size) As Variant
        For idx2 = start_idx To end_idx
            For idx3 = 1 To column_size
                newMatrix(newIndex, idx3) = matrix(idx2, idx3)
            Next
            newIndex = newIndex + 1
        Next
        
        return_value(idx1) = newMatrix
    Next
    
    getMatrixDivideRow = return_value
End Function

Private Function getCategoryIndexList(matrix As Variant, categoryIndexString As String) As Collection
    Dim categoryIndexArray  As Variant
    categoryIndexArray = Split(categoryIndexString, ",")
    
    Dim col As Collection
    Set col = New Collection
    
    ' スタートインデックスはカテゴリ開始データ。
    Call col.Add(LBound(matrix, 1))
    
    ' カテゴリのインデックスを指定して、サイズを求める。
    For idx1 = LBound(matrix, 1) + 1 To UBound(matrix, 1)
        Dim currData As Variant
        Dim previousData As Variant
        currData = Application.Index(matrix, idx1)
        previousData = Application.Index(matrix, idx1 - 1)
        
        flg = True
        ' 指定カテゴリごとにループ。
        For idx2 = LBound(categoryIndexArray) To UBound(categoryIndexArray)
            categoryIndex = categoryIndexArray(idx2)
            ' 【カテゴリについて】
            ' 現在データと１つ前のデータで違いがあれば異なるカテゴリと判断する。
            If currData(categoryIndex) <> "" Then
                If currData(categoryIndex) <> previousData(categoryIndex) Then
                    flg = False
                    Exit For
                End If
            End If
        Next
        
        If flg = False Then
            Call col.Add(idx1)
        End If
    Next
    
    Set getCategoryIndexList = col
End Function


Private Function isSameCategory(data1 As Variant, data2 As Variant, categoryIndexString As String) As Boolean
    Dim flg As Boolean
    flg = True
    
    Dim categoryIndexArray  As Variant
    categoryIndexArray = Split(categoryIndexString, ",")
    
    ' 指定カテゴリごとにループ。
    For idx = LBound(categoryIndexArray) To UBound(categoryIndexArray)
        categoryIndex = categoryIndexArray(idx)
        ' ２つのデータで違いがあれば異なるカテゴリと判断する。
        If data1(categoryIndex) <> "" And data2(categoryIndex) <> "" Then
            If data1(categoryIndex) <> data2(categoryIndex) Then
                flg = False
                Exit For
            End If
        End If
    Next
    
    isSameCategory = flg
End Function


