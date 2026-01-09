Attribute VB_Name = "BMPModule"
Option Explicit
Option Private Module

'BMPファイルヘッダ構造体
Private Type tBITMAPFILEHEADER
    bfType As String * 2        'ファイルタイプ ("BM")
    bfSize As Long              'ビットマップ ファイルのサイズ (バイト単位)。
    bfReserved1 As Integer      '予約領域（※常に0）
    bfReserved2 As Integer      '予約領域（※常に0）
    bfOffBits As Long           'BITMAPFILEHEADER 構造体の先頭からビットマップ ビットまでのオフセット (バイト単位)。
End Type

'BMP情報ヘッダ構造体
Private Type tBITMAPINFOHEADER
    biSize As Long              'ヘッダサイズ。
    biWidth As Long             'イメージの幅(ピクセル単位)
    biHeight As Long            'イメージの高さ(ピクセル単位)
    biPlanes As Integer         'イメージプレーン数 (※常に1)
    biBitCount As Integer       'ピクセルあたりのビット数(1, 4, 8 or 24)
    biCompression As Long       '圧縮形式
    biSizeImage As Long         '圧縮されたイメージのサイズ(バイト単位)、または0
    biXPelsPerMeter As Long     '水平方向の解像度(1メートルあたりのピクセル数)
    biYPelsPerMeter As Long     '垂直方向の解像度(1メートルあたりのピクセル数)
    biClrUsed As Long           '使用するカラー数(biBitCountにより最大値決定。1bit=max2、4bit=max16、8bit=max256、24bit=0)
    biClrImportant As Long      '重要なカラー数(使用カラー数に制限のある環境用)
End Type

'カラーマップ構造体
Private Type tRGBQUAD
    rgbBlue As Byte             'カラーマップエントリの青の値
    rgbGreen As Byte            'カラーマップエントリの緑の値
    rgbRed As Byte              'カラーマップエントリの赤の値
    rgbReserved As Byte         '予約領域 (※常に0)
End Type

'ビットマップ情報
Private Type tBITMAPINFO
    fileHedaer As tBITMAPFILEHEADER 'BMPファイルヘッダ
    infoHeader As tBITMAPINFOHEADER 'BMP情報ヘッダ
    colorMap() As tRGBQUAD          'カラーマップ
    Body() As Byte                  '画像本体
End Type

'''指定したパスに生成したビットマップを出力する。
''' @param Path    / I / 出力パス
''' @param pData   / I / 24bitカラー2次元マップ、またはインデックスカラー2次元マップ
''' @param pColors / I / カラーインデックス(省略した場合はpDataを24bitカラー2次元マップとみなす)
''' @param pDpi    / I / 解像度DPI(インチあたりのピクセル数)(省略時は72)
''' @return 成功した場合はTrue
Public Function ExportBMPFile(ByVal Path As String, ByRef pData() As Variant, Optional ByRef pColors As Variant, Optional ByVal pDpi As Integer = 72) As Boolean
    Dim bmpInfo As tBITMAPINFO
    Dim fileNo As Long

    ExportBMPFile = False

    If Not CreateBMPInfo(bmpInfo, pData, pColors, pDpi) Then
        Exit Function
    End If

    On Error GoTo OutputError

    If Dir(Path) <> "" Then Kill Path

    fileNo = FreeFile()
    Open Path For Binary As #fileNo

    With bmpInfo
        'ファイルヘッダの出力
        Put #fileNo, , .fileHedaer
        '情報ヘッダの出力
        Put #fileNo, , .infoHeader

        If .infoHeader.biClrUsed > 0 Then
            'カラーマップの出力
            Put #fileNo, , .colorMap
        End If

        'ビットマップ情報の出力
        Put #fileNo, , .Body
    End With

    Close #fileNo

    ExportBMPFile = True
OutputError:
End Function

'''24bitカラー2次元マップからカラーインデックスとカラーインデックス2次元マップを生成する。
''' @param pIndexMap   / O / カラーインデックス2次元マップ
''' @param pColors     / IO / カラーインデックス。最大256色。あらかじめ固定色を指定可能。
''' @param p24ColorMap / I / 24bitカラー2次元マップ
''' @return 成功した場合はTrue。カラーインデックスが256色を超える場合はFalse
Public Function CreateIndexMap(ByRef pIndexMap() As Variant, ByRef pColors() As Variant, ByRef p24ColorMap() As Variant) As Boolean
    Dim clCand() As Long
    Dim base As Long
    Dim rIdx As Long, cIdx As Long, lIdx As Long
    Dim clr As Long

    '色候補先頭に固定色設定
    If (Not pColors) <> -1 Then
        base = LBound(pColors)
        ReDim clCand(0 To 1, 0 To UBound(pColors) - base)
        For lIdx = base To UBound(pColors)
            clCand(0, lIdx - base) = pColors(lIdx)
            clCand(1, lIdx - base) = 0
        Next lIdx
    End If

    '色ごとに同色ピクセル数を算出
    For rIdx = LBound(p24ColorMap, 2) To UBound(p24ColorMap, 2)
        For cIdx = LBound(p24ColorMap, 1) To UBound(p24ColorMap, 1)
            clr = p24ColorMap(rIdx, cIdx)

            If (Not clCand) = -1 Then
                ReDim Preserve clCand(0 To 1, 0 To 0)
                clCand(0, 0) = clr
                clCand(1, 0) = 1

            Else
                For lIdx = LBound(clCand, 2) To UBound(clCand, 2)
                    If clCand(0, lIdx) = clr Then Exit For
                Next lIdx
                If lIdx > UBound(clCand, 2) Then
                    ReDim Preserve clCand(0 To 1, 0 To lIdx)
                    clCand(0, lIdx) = clr
                    clCand(1, lIdx) = 1

                Else
                    clCand(1, lIdx) = clCand(1, lIdx) + 1
                End If
            End If
        Next cIdx
    Next rIdx

    If UBound(clCand, 2) > 255 Then
        '色候補数が256を超える場合
        CreateIndexMap = False

        Exit Function
    End If

    '固定色の次のインデックス
    base = 0
    If (Not pColors) <> -1 Then
        base = UBound(pColors) - LBound(pColors) + 1
    End If

    '固定色を除外して色候補をピクセル数が多い順にソート。
    If base <= UBound(clCand, 2) - 1 Then
        For rIdx = base To UBound(clCand, 2) - 1
            For cIdx = rIdx + 1 To UBound(clCand, 2)
                If clCand(1, rIdx) < clCand(1, cIdx) Then
                   clr = clCand(0, rIdx): clCand(0, rIdx) = clCand(0, cIdx): clCand(0, cIdx) = clr
                   clr = clCand(1, rIdx): clCand(1, rIdx) = clCand(1, cIdx): clCand(1, cIdx) = clr
                End If
            Next cIdx
        Next rIdx
    End If

    'カラーインデックスの作成
    ReDim pColors(LBound(clCand, 2) To UBound(clCand, 2))
    For lIdx = LBound(clCand, 2) To UBound(clCand, 2)
        pColors(lIdx) = clCand(0, lIdx)
    Next lIdx

    'カラーインデックスに対応したインデックス2次元マップ作成
    ReDim pIndexMap(LBound(p24ColorMap, 1) To UBound(p24ColorMap, 1), LBound(p24ColorMap, 2) To UBound(p24ColorMap, 2))
    For rIdx = LBound(pIndexMap, 2) To UBound(pIndexMap, 2)
        For cIdx = LBound(pIndexMap, 1) To UBound(pIndexMap, 1)
            For lIdx = LBound(pColors) To UBound(pColors)
                If pColors(lIdx) = p24ColorMap(rIdx, cIdx) Then
                    pIndexMap(rIdx, cIdx) = lIdx
                    Exit For
                End If
            Next lIdx
        Next cIdx
    Next rIdx

    CreateIndexMap = True
End Function

'''ビットマップ情報作成
''' @param info    / O / ビットマップ情報
''' @param pData   / I / 1ピクセル1要素の2次元配列(インデックスマップまたは24bitカラーマップ)
''' @param pColors / I / カラーインデックス
''' @param pDpi    / I / DPI
''' @return 成功した場合はTrue
Private Function CreateBMPInfo(ByRef info As tBITMAPINFO, ByRef pData() As Variant, ByRef pColors As Variant, ByVal pDpi As Integer) As Boolean
    Dim Width As Long, bWidth As Long, bPadding As Long
    Dim Height As Long
    Dim ClrUsed As Long
    Dim BitCount As Long
    Dim idx As Long, Row As Long, Col As Long, x As Long
    Dim word As Byte
    Dim n As Long

    CreateBMPInfo = False

    '画像サイズ計算
    Width = -1
    Height = -1
    On Error Resume Next
    Width = UBound(pData, 2) - LBound(pData, 2) + 1
    Height = UBound(pData, 1) - LBound(pData, 1) + 1
    On Error GoTo 0

    If Width <= 0 Or Height <= 0 Then
        Exit Function
    End If

    '色数とビット数の計算
    If IsMissing(pColors) Then
        ClrUsed = 0
        BitCount = 24

    Else
        ClrUsed = -1
        On Error Resume Next
        ClrUsed = UBound(pColors) - LBound(pColors) + 1
        On Error GoTo 0

        If ClrUsed <= 0 Then
            Exit Function
        End If

        If ClrUsed <= 2 Then
            BitCount = 1

        ElseIf ClrUsed <= 16 Then
            BitCount = 4

        ElseIf ClrUsed <= 256 Then
            BitCount = 8

        Else
            Exit Function
        End If

        ReDim info.colorMap(LBound(pColors) To UBound(pColors))
        For idx = LBound(pColors) To UBound(pColors)
            With info.colorMap(idx)
                SplitRGB CLng(pColors(idx)), .rgbRed, .rgbGreen, .rgbBlue
            End With
        Next idx
    End If

    'ビット数と画像幅からバイト単位画面幅とパディングを計算
    Select Case BitCount
    Case 1:     bWidth = (Width + 7) \ 8
    Case 4:     bWidth = (Width + 1) \ 2
    Case 8:     bWidth = Width
    Case 24:    bWidth = Width * 3
    Case Else:  Exit Function
    End Select
    bPadding = (4 - bWidth Mod 4) Mod 4

    'ビットマップ情報のサイズ定義
    ReDim info.Body(0 To (bWidth + bPadding) * Height - 1)

    'ファイルヘッダの定義
    With info.fileHedaer
        .bfType = "BM"
        .bfOffBits = Len(info.fileHedaer) + Len(info.infoHeader) + 4 * ClrUsed
        .bfSize = .bfOffBits + UBound(info.Body) + 1
    End With

    '情報ヘッダの定義
    With info.infoHeader
        .biSize = Len(info.infoHeader)
        .biWidth = Width
        .biHeight = Height
        .biPlanes = 1
        .biBitCount = BitCount
        .biCompression = 0
        .biSizeImage = 0
        .biXPelsPerMeter = Int(pDpi * 39.3701)
        .biYPelsPerMeter = Int(pDpi * 39.3701)
        .biClrUsed = ClrUsed
        .biClrImportant = ClrUsed
    End With

    'ビットマップ情報の生成
    idx = 0
    For Row = UBound(pData, 2) To LBound(pData, 2) Step -1
        x = 0
        word = 0
        For Col = LBound(pData, 1) To UBound(pData, 1)
            n = CLng(pData(Row, Col))
            Select Case BitCount
            Case 1
                word = word * 2 + (n And 1)
                x = x + 1
                If x Mod 8 = 0 Then
                    info.Body(idx) = word
                    idx = idx + 1
                    word = 0
                End If

            Case 4
                word = word * &H10 + (n And &HF)
                x = x + 1
                If x Mod 2 = 0 Then
                    info.Body(idx) = word
                    idx = idx + 1
                    word = 0
                End If

            Case 8
                info.Body(idx) = n And &HFF
                idx = idx + 1

            Case 24
                SplitRGB n, info.Body(idx + 2), info.Body(idx + 1), info.Body(idx)
                idx = idx + 3
            End Select
        Next Col

        Select Case BitCount
        Case 1
            If x Mod 8 > 0 Then
                word = word * 2 ^ (8 - x Mod 8)
                info.Body(idx) = word
                idx = idx + 1
            End If

        Case 4
            If x Mod 2 > 0 Then
                word = word * &H10
                info.Body(idx) = word
                idx = idx + 1
            End If
        End Select

        idx = idx + bPadding
    Next Row

    CreateBMPInfo = True
End Function

Private Sub Test_ExportBMPFile()
    Dim Path As String
    Dim data() As Variant
    Dim colors() As Variant
    Dim i%, j%

    Path = ThisWorkbook.Path

    'モノクロビットマップ試験
    ReDim data(0 To 100, 0 To 100)
    For i = 0 To 100: For j = 0 To 100: data(i, j) = 0: Next j, i
    For i = 0 To 100: data(i, i) = 1: Next i
    colors = Array(vbWhite, vbBlack)
    ExportBMPFile Path & "\test1.bmp", data, colors, 72

    '4bitビットマップ試験
    ReDim data(0 To 99, 0 To 99)
    For i = 0 To 99: For j = 0 To 99: data(i, j) = (i + j) Mod 8: Next j, i
    colors = Array(vbBlack, vbBlue, vbRed, vbMagenta, vbCyan, vbYellow, vbGreen, vbWhite)
    ExportBMPFile Path & "\test4.bmp", data, colors, 72

    '8bitビットマップ試験
    ReDim colors(0 To 255)
    For i = 0 To 255: colors(i) = RGB(i, i, i): Next i
    ReDim data(0 To 255, 0 To 255)
    For i = 0 To 255: For j = 0 To 255: data(i, j) = (i + j) Mod 256: Next j, i
    ExportBMPFile Path & "\test8.bmp", data, colors, 72

    '24bitビットマップ試験
    ReDim data(0 To 255, 0 To 255)
    For i = 0 To 255: For j = 0 To 255: data(i, j) = RGB(i, j, (i + j) Mod 256): Next j, i
    ExportBMPFile Path & "\test24.bmp", data, , 10
End Sub
