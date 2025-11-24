Attribute VB_Name = "CommonModule"
Option Explicit
Option Private Module

'''文字列を配列に格納する
'''引数
'''  Texts     / IO / 格納する配列。
'''  Text      / I  / 格納される文字列。
Public Sub AddArrayText(ByRef Texts() As String, ByVal Text As String)
    Dim Size As Integer

    If (Not Texts) = -1 Then
        Size = 0

    Else
        Size = UBound(Texts) + 1
    End If

    ReDim Preserve Texts(Size)
    Texts(Size) = Text
End Sub


'''指定シートの最終行を取得する。
'''引数
'''  sh      / I / 指定シート
'''  needVal / I / Trueの場合、ブランク行をスキップする。(省略時はTrue)
'''戻り値
'''  有効値のある最終行
Public Function GetLastUsedRow(ByRef sh As Worksheet, Optional needVal As Boolean = True) As Long
    With sh.UsedRange
        '使用Rangeの最終行を取得
        GetLastUsedRow = .Rows(.Rows.Count).Row
    End With

    If Not needVal Then Exit Function

    With sh
        'UsedRangeは書式だけ変更された行もカウントするので
        'ブランク行を後からスキップする。
        Do While IsBlankRange(sh.Rows(GetLastUsedRow))
            GetLastUsedRow = GetLastUsedRow - 1
        Loop
    End With
End Function

'''指定Rangeが全てブランクかどうかを判定する。
'''引数
'''  Rng / I / 指定Range
'''戻り値
'''  指定Rangeが全てブランクならばTrue。そうでない場合はFalseを返却。
Public Function IsBlankRange(ByRef Rng As Range) As Boolean
    IsBlankRange = (WorksheetFunction.CountBlank(Rng) = Rng.Count)
End Function

'''指定された文字コードでテキストファイルの一括読込を行う
'''引数
'''  Result / O / 読み込んだテキストデータ
'''  Path   / I / 読み込み対象のファイルパス
'''  charSet/ I / 文字コード
'''戻り値
'''  成功した場合はTrue、失敗した場合はFalseを返却
Public Function ReadTextFile(ByRef Result As String, ByVal Path As String, Optional ByVal charSet As String = "UTF-8") As Boolean
    Dim buf As String

    ReadTextFile = False

    On Error GoTo ErrProc
    With CreateObject("ADODB.Stream")
        .Type = 2
        .charSet = charSet
        .Open
        .LoadFromFile Path
        Result = .ReadText
        .Close
    End With
    On Error GoTo 0

    If Right(Result, 4) = vbCrLf & vbCrLf Then
        Result = Left(Result, Len(Result) - 2)
    End If

    ReadTextFile = True
    Exit Function
    
ErrProc:
End Function

'''バイナリファイルの一括読込を行う
'''引数
'''  Result / O / 読み込んだバイナリデータ
'''  Path   / I / 読み込み対象のファイルパス
'''戻り値
'''  成功した場合はTrue、失敗した場合はFalseを返却
Public Function ReadBinaryFile(ByRef Result() As Byte, ByVal Path As String) As Boolean
    Dim buf As String

    ReadBinaryFile = False

    On Error GoTo ErrProc
    With CreateObject("ADODB.Stream")
        .Type = 1
        .Open
        .LoadFromFile Path
        Result = .Read
        .Close
    End With
    On Error GoTo 0

    ReadBinaryFile = True
    Exit Function
    
ErrProc:
End Function

'''指定したバイナリをBase64に変換する
'''引数
'''  buf     / I / 入力バイナリ
'''  folding / I / 改行有無(Trueの場合、入力57byte、出力76文字単位で改行する)
'''戻り値
'''  Base64化した文字列
Public Function ConvertBase64(ByRef buf() As Byte, Optional ByVal folding As Boolean = True) As String
    Static B64CHR() As String
    Dim Result As String
    Dim idx As Long
    Dim Pos As Integer

    If (Not B64CHR) = -1 Then
        For idx = Asc("A") To Asc("Z"): AddArrayText B64CHR, Chr(idx): Next idx
        For idx = Asc("a") To Asc("z"): AddArrayText B64CHR, Chr(idx): Next idx
        For idx = Asc("0") To Asc("9"): AddArrayText B64CHR, Chr(idx): Next idx
        AddArrayText B64CHR, "+"
        AddArrayText B64CHR, "/"
    End If

    '8bit * 3 AAAAAABB-BBBBCCCC-CCDDDDDD
    'を
    '6bit * 4 AAAAAA-BBBBBB-CCCCCC-DDDDDD
    'に変換し、6bitを64種の文字に変換して出力

    Result = ""

    idx = LBound(buf)
    Do While idx <= UBound(buf)
        'AAAAAAxx ⇒ AAAAAA
        Pos = Int(buf(idx) / 4)
        Result = Result & B64CHR(Pos)

        'xxxxxxBB ⇒ BB0000
        Pos = (buf(idx) Mod 4) * 16

        idx = idx + 1
        If idx > UBound(buf) Then
            Result = Result & B64CHR(Pos) & "=="
            Exit Do
        End If

        'BBBBxxxx ⇒ BBBB
        'BB0000 + BBBB = BBBBBB
        Pos = Pos + Int(buf(idx) / 16)
        Result = Result & B64CHR(Pos)

        'xxxxCCCC ⇒ CCCC00
        Pos = (buf(idx) Mod 16) * 4

        idx = idx + 1
        If idx > UBound(buf) Then
            Result = Result & B64CHR(Pos) & "="
            Exit Do
        End If

        'CCxxxxxx ⇒ CC
        'CCCC00 + CC = CCCCCC
        Pos = Pos + Int(buf(idx) / 64)
        Result = Result & B64CHR(Pos)

        'xxDDDDDD ⇒ DDDDDD
        Pos = buf(idx) Mod 64
        Result = Result & B64CHR(Pos)

        idx = idx + 1

        If folding And idx Mod 57 = 0 Then '19 * 4 = 76, 19 * 3 = 57
            Result = Result & vbLf
        End If
    Loop
    ConvertBase64 = Result
End Function
