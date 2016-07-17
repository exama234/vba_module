Public Enum era
    ' 使用予定なし。
    明治 = 0
    大正 = 1
    昭和 = 2
    平成 = 3
End Enum



''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 引数指定の和暦年月日が存在するかをチェックする。
' 　　　　： ただし明治6年以降を対象としています。
' 　　　　： うるう年考慮。
' 　　　　： 旧暦からの改暦によるずれは非考慮。つまり明治1~5年は正しく処理できてません。
' 引数１　： 元号。
' 引数２　： 年。
' 引数３　： 月。
' 引数４　： 日。
' 返り値　： 和暦として存在する場合、真を返す。
'
' 依存Mod ： 要「Mod_Date」モジュール
' 使用方法： If isExistJpCalendar('昭和', 22, 3, 4) Then
'         ：     MsgBox ("和暦として存在しない。")
'         ： End If
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function isExistJpCalendar(arg_era As String, arg_year As Integer, arg_month As Integer, arg_day As Integer) As Boolean
    Dim date1 As Date
    date1 = getDominicalDate(arg_era, arg_year, arg_month, arg_day)
    
    If date1 = 0 Then
        ' 和暦として存在しない。
        Exit Function
    End If
    
    ' 和暦として存在する。
    isExistJpCalendar = True
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 引数指定の和暦年月日を日付型に変換する。
' 　　　　： ただし明治6年以降を対象としています。
' 　　　　： うるう年考慮。
' 　　　　： 旧暦からの改暦によるずれは非考慮。つまり明治1~5年は正しく処理できてません。
' 引数１　： 元号。
' 引数２　： 年。
' 引数３　： 月。
' 引数４　： 日。
' 返り値　： 和暦として存在する場合、日付型を返す。
' 　　　　： 変換に失敗した場合、「0:00:00（数値としてはゼロ）」を返す。
'
' 依存Mod ： 要「Mod_Date」モジュール
' 使用方法： date1 = Mod_JpCalendar.getDominicalDate("昭和", 22, 3, 4)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function getDominicalDate(arg_era As String, arg_year As Integer, arg_month As Integer, arg_day As Integer) As Date
    If arg_year < 1 Then
        ' 存在しない年。
        Exit Function
    End If
    If arg_month < 1 Or 12 < arg_month Then
        ' 存在しない月。
        Exit Function
    End If
    If arg_day < 1 Or 31 < arg_day Then
        ' 存在しない日。
        Exit Function
    End If
    
    Select Case Strings.Trim(arg_era)
        ' 【豆知識】
        ' 慶応4年9月8日。
        ' 「今日から明治を使うよー。
        ' でもきりが悪いから今までの慶応4年1月1日～9月8日は明治元年1月1日～9月8日でもあるよー。
        ' かぶるけど気にしないでね。」
        ' 【豆知識】
        ' 明治45年7月30日。
        ' 「今日から大正を使うよー。今日だけは明治45年7月30日でも大正元年7月30日でもあるよー。
        ' かぶるけど気にしないでね。」
        ' 【豆知識】
        ' 大正15年12月25日。
        ' 「今日から昭和を使うよー。今日だけは大正15年12月25日でも昭和元年12月25日でもあるよー。
        ' かぶるけど気にしないでね。」
        ' 【豆知識】
        ' 昭和64年1月7日。
        ' 「今日は昭和64年1月7日、明日は平成元年1月8日だよー。今回はかぶらないから安心してね。」
        Case "明治"
            ' 1868/01/01 - 1912/07/30（明治45年まで）
            start_date = #1/1/1868#
            end_date = #7/30/1912#
        Case "大正"
            ' 1912/07/30 - 1926/12/25（大正15年まで）
            start_date = #7/30/1912#
            end_date = #12/25/1926#
        Case "昭和"
            ' 1926/12/25 - 1989/01/07（昭和64年まで）
            start_date = #12/25/1926#
            end_date = #1/7/1989#
        Case "平成"
            ' 1989/01/08 - 20XX/XX/XX（ちよに やちよに）
            start_date = #1/8/1989#
            end_date = #12/31/9999#
        Case Else
            ' 存在しない元号。
            Exit Function
    End Select
    
    ' 西暦（文字列）に変換する。
    Dim base_year, dominical_year As Integer
    base_year = Year(start_date)
    dominical_year = base_year + arg_year - 1
    str_yyyymmdd = dominical_year & "/" & CStr(arg_month) & "/" & CStr(arg_day)
    If dominical_year < 1873 Then
        ' 明治6年より前は対象外とする。
        Exit Function
    End If
    
    
    Dim last_day As Variant
    last_day = Array(-1, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    ' モジュール呼び出し（うるう年判定）
    If Mod_Date.isIntercalaryYear(dominical_year) Then
        ' うるう年2月は29日まで。
        last_day(2) = 29
    End If
    If last_day(arg_month) < arg_day Then
        ' 存在しない日。
        Exit Function
    End If
    
    ' 西暦（日付型）に変換する。
    Dim dominical_date As Date
    dominical_date = DateValue(str_yyyymmdd)
    
    
    If dominical_date < start_date Or end_date < dominical_date Then
        ' 存在しない日付。年号の最終日を超えている。
        Exit Function
    End If
    
    
    ' 和暦として存在する。
    getDominicalDate = dominical_date
End Function

