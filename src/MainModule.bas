Attribute VB_Name = "MainModule"
Option Explicit

'''変換対象文字列をQRコードに変換し、指定Rangeに書き込む。
''' @param pRng     / IO / QRコード出力先
''' @param pInfo    / O / エラー情報またはバージョン情報
''' @param pTarget  / I / 変換対象文字列
''' @param pECL     / I / エラー補正レベル(省略時はECL_L)
''' @param pMaskPtn / I / マスクパターン(省略時はMSK_AUTO)
''' @param pModeSet / I / モードセット(省略時はMOD_ALL)
Public Sub WriteQRCode(ByRef pRng As Range, ByRef pInfo As String, ByVal pTarget As String _
        , Optional ByVal pECL As eErrorCorrectionLevel = ECL_L _
        , Optional ByVal pMaskPtn As eMaskType = MSK_AUTO _
        , Optional ByVal pModeSet As eModeBit = MOD_ALL _
        )
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
    pRng.RowHeight = 19.5

    If GetQRCode(ar, pInfo, pTarget, pECL, pMaskPtn, pModeSet) Then
        OutputRange pRng.Offset(1, 0), ar

    Else
        With pRng.Offset(1, 0)
            .RowHeight = 18.75
            .Value = pInfo
        End With
    End If
End Sub

'''QRコードが格納された二次元配列を指定Rangeに書き込む。
'''配列周りに4ドット分の空白を入れる。
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

'''指定したファイルを指定文字コードで読み込み、QRコードで表現できるサイズで分割する。
''' @param Result   / O / 分割結果配列
''' @param Path     / I / 読み込み対象のファイルパス
''' @param charSet  / I / 指定文字コード(省略時は"UTF-8")
'''                       "binary"で始まる文字列が指定された場合、ファイルパスの指す先を
'''                       バイナリファイルとして読み込んでBase64化したものを対象とする｡
''' @param pECL     / I / エラー補正レベル(省略時はECL_L)
''' @param pModeSet / I / モードセット(省略時はMOD_ALL)
''' @return 成功した場合はTrue
Public Function SplitFile(ByRef Result() As String, ByVal Path As String _
        , Optional ByVal charSet As String = "UTF-8" _
        , Optional ByVal pECL As eErrorCorrectionLevel = ECL_L _
        , Optional ByVal pModeSet As eModeBit = MOD_ALL _
        ) As Boolean
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
        v = CheckQRCode(curText, pECL, pModeSet)
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

'''QRコードの画像を描画する。
''' @param pRng        / I / 出力先セル位置指定
''' @param pInfo       / O / QRコード実行結果
''' @param FilePath    / I / 出力ファイルパス。省略時は一時ファイルで作成する。
''' @param pElementRng / I / ドット要素として使用する画像情報を生成するソースとして使用するRange。
''' @param bgColor     / I / 背景色
''' @param pTarget     / I / 変換対象文字列
''' @param pECL        / I / エラー補正レベル(省略時はECL_L)
''' @param pMaskPtn    / I / マスクパターン(省略時はMSK_AUTO)
''' @param pModeSet    / I / モードセット(省略時はMOD_ALL)
Public Sub DrawQRCodeImage(ByRef pRng As Range, ByRef pInfo As String, ByVal FilePath As String, ByRef pElementRng As Range, ByVal bgColor As Long, ByVal pTarget As String _
        , Optional ByVal pECL As eErrorCorrectionLevel = ECL_L _
        , Optional ByVal pMaskPtn As eMaskType = MSK_AUTO _
        , Optional ByVal pModeSet As eModeBit = MOD_ALL _
        )
    Dim ar() As Variant
    Dim isLink As MsoTriState
    Dim eImg() As Variant
    Dim colors() As Variant
    Dim bmpBody() As Variant
    Dim c As Long, r As Long
    Dim sh As Worksheet

    'QRコードの生成
    If Not GetQRCode(ar, pInfo, pTarget, pECL, pMaskPtn, pModeSet) Then
        Exit Sub
    End If

    'ファイルパスが未指定の場合、テンポラリファイルパスを使用
    If FilePath = "" Then
        FilePath = CreateTemporaryFilePath("bmp")
        isLink = msoFalse

    Else
        If InStr(FilePath, "\") = 0 Then
            'ファイル名のみの場合
            FilePath = Application.ThisWorkbook.Path & "\" & FilePath
        End If

        isLink = msoTrue
    End If

    If pElementRng Is Nothing Then
        '要素Rangeが未指定の場合、4pixel角の黒四角を要素とする。
        ReDim eImg(1 To 4, 1 To 4)
        For r = 1 To 4: For c = 1 To 4: eImg(r, c) = 1: Next c, r
        colors = Array(bgColor, vbBlack)

    Else
        '要素Rangeからピクセル要素画像を生成する
        CreateElement eImg, colors, pElementRng, bgColor
    End If

    If (Not colors) = -1 Then
        'カラーインデックスがない場合
        'QRコードと要素画像をかけ合わせて24bitカラー2次元マップを生成
        BuildImage bmpBody, ar, eImg, bgColor

    Else
        'カラーインデックスがある場合
        'QRコードと要素画像をかけ合わせてインデックスカラー2次元マップを生成
        BuildImage bmpBody, ar, eImg, 0
    End If

    'BMP画像を出力
    If Not ExportBMPFile(FilePath, bmpBody, colors) Then
        Exit Sub
    End If

    '生成したBMP画像を表示する
    Set sh = pRng.Parent
    sh.Shapes.AddPicture FilePath, isLink, msoTrue, pRng.Left, pRng.Top, -1, -1

    '一時ファイルの削除
    If isLink = msoFalse Then
        Kill FilePath
    End If
End Sub

'''要素画像の生成
''' @param pBody   / O / 24bitカラー2次元マップ、またはインデックスカラー2次元マップ
''' @param pColors / O / カラーインデックス
''' @param pRng    / I / 要素Range
''' @param bgColor / I / 背景色
Private Sub CreateElement(ByRef pBody() As Variant, ByRef pColors() As Variant, ByRef pRng As Range, ByVal bgColor As Long)
    Dim ar() As Variant
    Dim rIdx As Long, cIdx As Long

    'カラーインデックスに背景色設定
    ReDim pColors(0 To 0)
    pColors(0) = bgColor

    '指定Range範囲の背景色から24bitカラー二次元マップを生成。
    ReDim ar(1 To pRng.Rows.Count, 1 To pRng.Columns.Count)
    For rIdx = LBound(ar, 1) To UBound(ar, 1)
        For cIdx = LBound(ar, 2) To UBound(ar, 2)
            ar(rIdx, cIdx) = pRng.Cells(rIdx, cIdx).DisplayFormat.Interior.Color
        Next cIdx
    Next rIdx

    'インデックスカラー2次元マップとカラーインデックスを作成
    If Not CreateIndexMap(pBody, pColors, ar) Then
        '失敗した場合はカラーインデックスをクリアし、24bitカラーマップを返却する。
        Erase pColors
        ReDim pBody(LBound(ar, 1) To UBound(ar, 1), LBound(ar, 2) To UBound(ar, 2))
        For rIdx = LBound(ar, 1) To UBound(ar, 1)
            For cIdx = LBound(ar, 2) To UBound(ar, 2)
                pBody(rIdx, cIdx) = ar(rIdx, cIdx)
            Next cIdx
        Next rIdx
    End If
End Sub

'''2次元バーコードと要素カラー2次元マップからカラー2次元マップを生成する。
''' @param pResult  / O / 出力2次元マップ
''' @param pD2Code  / I / 2次元バーコード
''' @param pElement / I / 要素カラー2次元マップ
''' @param bgColor  / I / 背景色
Private Sub BuildImage(ByRef pResult() As Variant, ByRef pD2Code() As Variant, ByRef pElement() As Variant, ByVal bgColor As Long)
    Dim qh As Long, qw As Long, eh As Long, ew As Long
    Dim rdIdx As Long, cdIdx As Long, reIdx As Long, ceIdx As Long, rrIdx As Long, crIdx As Long

    qh = UBound(pD2Code, 1) - LBound(pD2Code, 1) + 1
    qw = UBound(pD2Code, 2) - LBound(pD2Code, 2) + 1
    eh = UBound(pElement, 1) - LBound(pElement, 1) + 1
    ew = UBound(pElement, 2) - LBound(pElement, 2) + 1

    ReDim pResult(1 To qh * eh, 1 To qw * ew)

    For rdIdx = LBound(pD2Code, 1) To UBound(pD2Code, 1)
        For cdIdx = LBound(pD2Code, 2) To UBound(pD2Code, 2)
            For reIdx = LBound(pElement, 1) To UBound(pElement, 1)
                rrIdx = (rdIdx - LBound(pD2Code, 1)) * eh + reIdx - LBound(pElement, 1) + 1
                For ceIdx = LBound(pElement, 2) To UBound(pElement, 2)
                    crIdx = (cdIdx - LBound(pD2Code, 2)) * ew + ceIdx - LBound(pElement, 2) + 1
                    If pD2Code(rdIdx, cdIdx) = 1 Then
                        pResult(rrIdx, crIdx) = pElement(reIdx, ceIdx)

                    Else
                        pResult(rrIdx, crIdx) = bgColor
                    End If
                Next ceIdx
            Next reIdx
        Next cdIdx
    Next rdIdx
End Sub

'''画像の読み込み。BMPファイルならば直接読み込みます。
'''BMPファイル以外ならばBMPファイルに変換してから読み込みます。
''' @param Path  / I / 入力ファイルパス
''' @param pData / O / 読み込んだ24bitカラーマップ
''' @param sh    / IO / 作業に使用するワークシート(一時的にオブジェクトを追加しますが、最終的には削除します。)(省略時はActiveSheetを使用)
''' @return 成功した場合はTrue
Public Function ImportImage(ByVal Path As String, ByRef pData() As Long, Optional ByRef sh As Worksheet = Nothing) As Boolean
    Dim errCd As eErrorCode
    Dim TmpPath As String
    Dim w As Single, h As Single

    'BMPファイルの読み込み
    If ImportBMPFile(Path, pData, errCd) Then
        ImportImage = True
        Exit Function
    End If
        
    If errCd <> ERR_FORMAT And errCd <> ERR_WINDOWS And errCd <> ERR_COMPRESS Then
        '読めない画像形式以外のエラー
        ImportImage = False
        Exit Function
    End If

    If sh Is Nothing Then
        Set sh = ActiveSheet
    End If

    '一時ファイルパスの生成
    TmpPath = CreateTemporaryFilePath("bmp")

    Application.ScreenUpdating = False
    With sh.Pictures.Insert(Path)
        '画像の大きさを取得
        w = .Width
        h = .Height
        'ピクチャーの削除
        .Delete
    End With
    '画像の大きさのグラフを生成。
    With sh.ChartObjects.Add(0, 0, w, h)
        '背景画像として画像ファイルを読み込み
        .Chart.SetBackgroundPicture Path
        '枠線の消去
        .ShapeRange.Line.Visible = msoFalse
        'BMP画像としてエクスポート
        '※DPI96として出力されるので、読み込んだ画像のDPIが96以外だと大きさが変わるので注意。
        .Chart.Export TmpPath, "bmp"
        'グラフの削除
        .Delete
    End With
    Application.ScreenUpdating = True

    '一時作成したBMPファイルを読み込み
    ImportImage = ImportBMPFile(TmpPath, pData)

    '一時ファイルの削除
    Kill TmpPath
End Function

'''24bitカラー試験用の要素作成。
'''256個超過のセル(縦横17x17とか)を選択して実行。
Private Sub Test_Create24BitElement()
    Dim Rng As Range
    Dim r%, c%, a%, b%

    Set Rng = Selection
    a = 256 \ Rng.Columns.Count
    b = 256 \ Rng.Rows.Count
    For r = 1 To Rng.Rows.Count
        For c = 1 To Rng.Columns.Count
            Rng.Cells(r, c).Interior.Color = RGB((r * b) Mod 256, (c * a) Mod 256, (r * b + c * a) Mod 256)
        Next c
    Next r
End Sub
