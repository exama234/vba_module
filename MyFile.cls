VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private streamObj As Object
Private nonBomFlg As Boolean
Private filename_fullpath As String


' コンストラクタ
Private Sub Class_Initialize()
    ' 現在の選択セルで初期化。
    Set streamObj = CreateObject("ADODB.Stream")
    nonBomFlg = True
End Sub



' ファイルのエンコードを指定するメソッド群。
' 指定できる文字エンコーディングの種類は、
' レジストリの以下パスにあるサブキーを確認
' HKEY_CLASSES_ROOT\MIME\Database\Charset
Public Sub setChara_SJIS()
    streamObj.Charset = "Shift-JIS"
End Sub
Public Sub setChara_JIS()
    streamObj.Charset = "iso-2022-jp"
End Sub
Public Sub setChara_EUC()
    streamObj.Charset = "euc-jp"
End Sub
Public Sub setChara_Unicode()
    streamObj.Charset = "Unicode"
End Sub
Public Sub setChara_UTF7()
    streamObj.Charset = "UTF-7"
End Sub
Public Sub setChara_UTF8()
    streamObj.Charset = "UTF-8"
    nonBomFlg = False
End Sub
Public Sub setChara_UTF8N()
    streamObj.Charset = "UTF-8"
    nonBomFlg = True
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： 指定したファイルのストリームを開く。
' 引数１　： ファイル名。（絶対パス）
' 返り値　： なし。
' 使用方法： Dim file As MyFile
'            Set file = New MyFile
'            file.openStream ("C:\test.txt")
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub openStream(arg1 As String)
     ' 文字コードを指定してファイルをオープン
    streamObj.Open
    streamObj.LoadFromFile (arg1)
    ' 指定のファイル名で初期化。
    filename_fullpath = arg1
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： ファイルのストリームを閉じる。
' 引数　　： なし。
' 返り値　： なし。
' 使用方法： Dim file As MyFile
'            Set file = New MyFile
'            file.openStream ("C:\test.txt")
'            file.closeStream()
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub closeStream()
    streamObj.Close
    Set streamObj = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： ファイルを保存する。
' 引数　　： なし。
' 返り値　： なし。
' 使用方法： Dim file As MyFile
'            Set file = New MyFile
'            file.openStream ("C:\test.txt")
'            file.saveStream()
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub saveStream()
    If streamObj.Charset = "UTF-8" Then
        If nonBomFlg Then
            ' 先頭のBOM取り
            streamObj.Position = 0
            streamObj.Type = 1 '「1:TypeBinary」
            streamObj.Position = 3
            bytData = streamObj.read
            streamObj.Close
            
            ' UTF-8Nコードのデータを保存
            streamObj.Open
            streamObj.Type = 1 '「1:TypeBinary」
            streamObj.write bytData
        End If
    End If
    
    ' ファイルに書き込み
    streamObj.SaveToFile filename_fullpath, 2 '「2:adSaveCreateOverWrite」
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： ファイルの終端かを判断する真偽値を取得する。
' 引数　　： なし。
' 返り値　： なし。
' 使用方法： Dim file As MyFile
'            Set file = New MyFile
'            file.openStream ("C:\test.txt")
'            Do While Not file.isEos
'                MsgBox file.readLine
'            Loop
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function isEos() As Boolean
    isEos = streamObj.eos
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： ストリームから１行文字列を読み込む。
' 引数１　： 読み込み行番号。
' 返り値　： １行文字列。
' 使用方法： Dim file As MyFile
'            Set file = New MyFile
'            file.openStream ("C:\test.txt")
'            Do While Not file.isEos
'                MsgBox file.readLine
'            Loop
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function readLine(Optional line_number As Long) As String
    Dim str As String
    Dim pos As Long
    
    If line_number > 0 Then
        ' 読み込み行が指定された場合、ストリーム位置を先頭にする。
        streamObj.Position = 0
        
        ' 指定行 - 1回だけ行スキップする。
        Dim skip_count As Long
        skip_count = line_number - 1
        For idx = 1 To skip_count
            streamObj.SkipLine
        Next
    End If
    
    pos = streamObj.Position
    ' 行取得｡
    str = streamObj.ReadText(-2) ' 「-2:adReadLine」

    readLine = str
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： ストリームへ文字列を書き込む。
' 引数１　： 書き込む文字列。
' 返り値　： なし。
' 使用方法： Dim file As MyFile
'            Set file = New MyFile
'            file.openStream ("C:\test.txt")
'            file.writeStream ("Hello." & vbNewLine)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub writeStream(line As String)
    ' 書き込み前に現在位置を保持。
    pos = streamObj.Position
    ' ストリームの終端に移動後、書き込み。
    streamObj.Position = streamObj.Size
    streamObj.WriteText line
    ' 現在位置を戻す。
    streamObj.Position = pos
End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''
' 説明　　： ストリームへ文字列を書き込む。
' 引数１　： 書き込む文字列。
' 返り値　： なし。
' 使用方法： Dim file As MyFile
'            Set file = New MyFile
'            file.openStream ("C:\test.txt")
'            file.writeStream ("Hello.")
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub setStreamPosition(line_number As Long)
    If line_number < 1 Then
        ' 処理対象外。
        Exit Sub
    End If
    
    ' 読み込み行が指定された場合、ストリーム位置を先頭にする。
    streamObj.Position = 0
    
    ' 指定行 - 1回だけ行スキップする。
    Dim skip_count As Long
    skip_count = line_number - 1
    For idx = 1 To skip_count
        streamObj.SkipLine
    Next
End Sub


