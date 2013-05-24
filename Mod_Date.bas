Attribute VB_Name = "Mod_Date"
''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 指定年が閏年かをチェックする。
' 引数１　： 西暦年。
' 返り値　： 閏年の場合、真を返す。
' 使用方法： If isIntercalaryYear(2020) Then
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function isIntercalaryYear(arg_year As Integer) As Boolean
    If arg_year Mod 400 = 0 Then
        ' 400で割れる場合、閏年となる。
        isIntercalaryYear = True
        Exit Function
    End If
    If arg_year Mod 100 = 0 Then
        ' （400で割れない かつ）100で割れる場合、閏年とならない。
        isIntercalaryYear = False
        Exit Function
    End If
    If arg_year Mod 4 = 0 Then
        ' （100で割れない かつ）4で割れる場合、閏年となる。
        isIntercalaryYear = True
        Exit Function
    End If
    
    ' （100で割れない かつ）4で割れない場合、閏年とならない。
    isIntercalaryYear = False
End Function

