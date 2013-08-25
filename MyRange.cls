VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private rng As Range

' コンストラクタ
Private Sub Class_Initialize()
    ' 現在の選択セルで初期化。
    Set rng = Selection
End Sub


Public Sub init(argRange As Range)
    ' 指定の選択セルで初期化。
    Set rng = argRange
End Sub


' セル範囲のSetter
Public Sub setRange(ByVal argRng As Range)
    Set rng = argRng
End Sub

' セル範囲のGetter
Public Function getRange() As Range
    Set getRange = rng
End Function


Public Sub selectRange()
    rng.Select
End Sub


Private Function common(command As String, arg As Integer) As MyRange
    ' 再選択した範囲を取得する。
    Dim newRange As Range
    Set newRange = getRangeCell(rng, command, arg)
    
    
    ' 返り値は新たなオブジェクトとして返す。
    ' つまり自分自身の範囲は書き換えない。
    Dim newObj As MyRange
    Set newObj = New MyRange
    Call newObj.setRange(newRange)
    
    ' 返り値
    Set common = newObj
End Function

Public Function top(arg As Integer) As MyRange
    ' 返り値
    Set top = common("top", arg)
End Function
Public Function left(arg As Integer) As MyRange
    ' 返り値
    Set left = common("left", arg)
End Function
Public Function right(arg As Integer) As MyRange
    ' 返り値
    Set right = common("right", arg)
End Function
Public Function bottom(arg As Integer) As MyRange
    ' 返り値
    Set bottom = common("bottom", arg)
End Function


Public Function topLimit(arg As Integer) As MyRange
    ' 返り値
    Set topLimit = common("topLimit", arg)
End Function
Public Function leftLimit(arg As Integer) As MyRange
    ' 返り値
    Set leftLimit = common("leftLimit", arg)
End Function
Public Function rightLimit(arg As Integer) As MyRange
    ' 返り値
    Set rightLimit = common("rightLimit", arg)
End Function
Public Function bottomLimit(arg As Integer) As MyRange
    ' 返り値
    Set bottomLimit = common("bottomLimit", arg)
End Function


Public Function topResize(arg As Integer) As MyRange
    ' 返り値
    Set topResize = common("topResize", arg)
End Function
Public Function leftResize(arg As Integer) As MyRange
    ' 返り値
    Set leftResize = common("leftResize", arg)
End Function
Public Function rightResize(arg As Integer) As MyRange
    ' 返り値
    Set rightResize = common("rightResize", arg)
End Function
Public Function bottomResize(arg As Integer) As MyRange
    ' 返り値
    Set bottomResize = common("bottomResize", arg)
End Function


Public Function getRowsCount() As Long
    getRowsCount = rng.Rows.count
End Function
Public Function getColumnsCount() As Long
    getColumnsCount = rng.Columns.count
End Function


Public Function toString() As String
    toString = rng.Address
End Function
