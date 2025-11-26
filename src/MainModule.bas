Attribute VB_Name = "MainModule"
Option Explicit

'''変換対象文字列をQRコードに変換し、指定Rangeに書き込みます。
''' @param pRng    / IO / QRコード出力先
''' @param pInfo   / O / エラー情報またはバージョン情報
''' @param pTarget / I / 変換対象文字列
''' @param pECL    / I / エラー補正レベル(省略時はECL_L)
Public Sub WriteQRCode(ByRef pRng As Range, ByRef pInfo As String, ByVal pTarget As String, Optional ByVal pECL As eErrorCorrectionLevel = ECL_L)
    Dim Subject As String
    Dim ar() As Variant

    If Len(pTarget) <= 100 Then
        Subject = pTarget

    Else
        Subject = Left(pTarget, 48) & " … " & Right(pTarget, 48)
    End If
    Subject = Replace(Subject, vbCrLf, " ")
    Subject = Replace(Subject, vbLf, " ")
    pRng.Value = Subject

    If GetQRCode(ar, pInfo, pTarget, pECL) Then
        OutputRange pRng.Offset(1, 0), ar

    Else
        With pRng.Offset(1, 0)
            .RowHeight = 18.75
            .Value = pInfo
        End With
    End If
End Sub

'''QRコードが格納された二次元配列を指定Rangeに書き込みます。
'''配列周りに4ドット分の空白を入れます。
''' @param pRng / IO / 指定Range
''' @param ar   / I / 黒ドット部分に1が設定された二次元配列
Private Sub OutputRange(ByRef pRng As Range, ByRef ar() As Variant)
    Dim Rng As Range
    Dim fc As FormatCondition

    Set Rng = pRng.Resize(UBound(ar, 1) + 8, UBound(ar, 2) + 8)
    With Rng
        .ClearContents
        .ClearFormats
        .ColumnWidth = 0.45
        .RowHeight = 4.5
        .Interior.Color = vbWhite
    End With

    Set Rng = pRng.Resize(UBound(ar, 1), UBound(ar, 2)).Offset(4, 4)
    With Rng
        Set fc = .FormatConditions.Add(xlCellValue, xlEqual, "1")
        fc.Interior.Color = vbBlack
        Set fc = .FormatConditions.Add(xlCellValue, xlNotEqual, "1")
        fc.Interior.Color = vbWhite
        .Value = ar
    End With
End Sub

'''指定したファイルを指定文字コードで読み込み、QRコードで表現できるサイズで分割します。
''' @param Result  / O / 分割結果配列
''' @param Path    / I / 読み込み対象のファイルパス
''' @param charSet / I / 指定文字コード(省略時は"UTF-8")
'''                      "binary"で始まる文字列が指定された場合、ファイルパスの指す先を
'''                      バイナリファイルとして読み込んでBase64化したものを対象とします｡
''' @param pECL    / I / エラー補正レベル(省略時はECL_L)
''' @return 成功した場合はTrue
Public Function SplitFile(ByRef Result() As String, ByVal Path As String, Optional ByVal charSet As String = "UTF-8", Optional ByVal pECL As eErrorCorrectionLevel = ECL_L) As Boolean
    Dim buf() As Byte
    Dim Lines() As String
    Dim curText As String, oldText As String
    Dim idx As Long
    Dim v As Integer

    SplitFile = False

    If LCase(charSet) Like "binary*" Then
        If Not ReadBinaryFile(buf, Path) Then
            Exit Function
        End If

        curText = Trim(ConvertBase64(buf))
        curText = "begin-base64 664 " & Mid(Path, InStrRev(Path, "\") + 1) & vbLf _
                & curText & vbLf & "===="

    ElseIf Not ReadTextFile(curText, Path, charSet) Then
        Exit Function
    End If

    Lines = Split(curText, vbLf)
    curText = ""
    oldText = ""

    For idx = LBound(Lines) To UBound(Lines)
        curText = curText & Lines(idx) & vbLf
        v = CheckQRCode(curText, pECL)
        If v = 0 Then
            If oldText = "" Then Exit Function
            AddArrayText Result, oldText

            curText = ""
            idx = idx - 1
        End If

        oldText = curText
    Next idx
    AddArrayText Result, oldText
    SplitFile = True
End Function
