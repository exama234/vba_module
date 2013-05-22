Attribute VB_Name = "Mod_Date"
Public Function isIntercalaryYear(arg_year As Integer) As Boolean
    If arg_year Mod 400 = 0 Then
        isIntercalaryYear = True
        Exit Function
    End If
    If arg_year Mod 100 = 0 Then
        isIntercalaryYear = False
        Exit Function
    End If
    If arg_year Mod 4 = 0 Then
        isIntercalaryYear = True
        Exit Function
    End If
    isIntercalaryYear = False
    Exit Function
End Function

