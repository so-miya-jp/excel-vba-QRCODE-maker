Attribute VB_Name = "BMPModule"
Option Explicit
Option Private Module

'''エラーコード
Public Enum eErrorCode
    NML_SUCCESS = 0
    ERR_FILENOTFOUND = 1
    ERR_OPEN = 2
    ERR_READ = 3
    ERR_FORMAT = 4
    ERR_WINDOWS = 5
    ERR_COMPRESS = 6
    ERR_UNKNOWN = 9
End Enum

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
''' @param pDpi    / I / 解像度DPI(インチあたりのピクセル数)(省略時は96)
''' @return 成功した場合はTrue
Public Function ExportBMPFile(ByVal Path As String, ByRef pData() As Variant, Optional ByRef pColors As Variant, Optional ByVal pDpi As Integer = 96) As Boolean
    Dim bmpInfo As tBITMAPINFO

    ExportBMPFile = False

    'ビットマップ情報の生成
    If Not CreateBMPInfo(bmpInfo, pData, pColors, pDpi) Then
        Exit Function
    End If

    'ビットマップファイルの出力
    ExportBMPFile = WriteBMPFile(Path, bmpInfo)
End Function

'''指定したパスをBMPファイルとして読み込み、24bitカラーマップとして読み込む
''' @param Path  / I / 入力ファイルパス
''' @param pData / O / 24bitカラーマップ
''' @param errCd / O / 失敗時のエラーコード(省略可能)
''' @return Trueの場合、成功。
Public Function ImportBMPFile(ByVal Path As String, ByRef pData() As Long, Optional ByRef errCd As eErrorCode) As Boolean
    Dim bmpInfo As tBITMAPINFO

    ImportBMPFile = False

    'ビットマップファイルの入力
    errCd = ReadBMPFile(bmpInfo, Path)
    If errCd <> NML_SUCCESS Then
        Exit Function
    End If

    'カラーマップの作成
    CreateColorMap pData, bmpInfo

    ImportBMPFile = True
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
    For rIdx = LBound(p24ColorMap, 1) To UBound(p24ColorMap, 1)
        For cIdx = LBound(p24ColorMap, 2) To UBound(p24ColorMap, 2)
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
                    If lIdx > 255 Then
                        '色候補数が256を超える場合
                        CreateIndexMap = False

                        Exit Function
                    End If
                    
                    ReDim Preserve clCand(0 To 1, 0 To lIdx)
                    clCand(0, lIdx) = clr
                    clCand(1, lIdx) = 1

                Else
                    clCand(1, lIdx) = clCand(1, lIdx) + 1
                End If
            End If
        Next cIdx
    Next rIdx

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
    For rIdx = LBound(pIndexMap, 1) To UBound(pIndexMap, 1)
        For cIdx = LBound(pIndexMap, 2) To UBound(pIndexMap, 2)
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
            ClrUsed = 0
            BitCount = 24

        ElseIf ClrUsed <= 2 Then
            BitCount = 1

        ElseIf ClrUsed <= 16 Then
            BitCount = 4

        ElseIf ClrUsed <= 256 Then
            BitCount = 8

        Else
            Exit Function
        End If

        If ClrUsed > 0 Then
            ReDim info.colorMap(LBound(pColors) To UBound(pColors))
            For idx = LBound(pColors) To UBound(pColors)
                With info.colorMap(idx)
                    SplitRGB CLng(pColors(idx)), .rgbRed, .rgbGreen, .rgbBlue
                End With
            Next idx
        End If
    End If

    'ビット数と画像幅からバイト単位画面幅とパディングを計算
    GetByteWidth bWidth, bPadding, BitCount, Width

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
    For Row = UBound(pData, 1) To LBound(pData, 1) Step -1
        x = 0
        word = 0
        For Col = LBound(pData, 2) To UBound(pData, 2)
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

'''BMPファイルの書き込み
''' @param Path    / I / 出力ファイルパス
''' @param bmpInfo / I / ビットマップ情報
''' @return 成功した場合はTrue
Private Function WriteBMPFile(ByVal Path As String, ByRef bmpInfo As tBITMAPINFO) As Boolean
    Dim fileNo As Long

    WriteBMPFile = False
    On Error GoTo OutputError

    '既存ファイルの削除
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

    WriteBMPFile = True
OutputError:
End Function

'''BMPファイルの読み込み
''' @param bmpInfo / O / BMP情報
''' @param Path    / I / 入力ファイルパス
''' @return 0の場合成功。1以上の場合はエラーID
Private Function ReadBMPFile(ByRef bmpInfo As tBITMAPINFO, ByVal Path As String) As eErrorCode
    Dim fileNo As Long
    Dim biSize As Long
    Dim ClrUsed As Long
    Dim bWidth As Long, bPadding As Long

    ReadBMPFile = ERR_UNKNOWN
    On Error GoTo InputError
    
    If Dir(Path) = "" Then
#If DEBUG_ > 0 Then
        Debug.Print "ファイルが見つかりません。(" & Path & ")"
#End If
        ReadBMPFile = ERR_FILENOTFOUND
        Exit Function
    End If

    ReadBMPFile = ERR_OPEN
    fileNo = FreeFile()
    Open Path For Binary As #fileNo
        
    With bmpInfo
        ReadBMPFile = ERR_READ
        'ファイルヘッダの読み取り
        Get #fileNo, , .fileHedaer

        If .fileHedaer.bfType <> "BM" Then
#If DEBUG_ > 0 Then
            Debug.Print "ファイル形式がBMP形式ではありません。(" & Path & ")"
#End If
            ReadBMPFile = ERR_FORMAT
            GoTo InputError
        End If

        '情報ヘッダの読み取り
        Get #fileNo, , .infoHeader

        With .infoHeader
            If .biSize <> 40 Or .biPlanes <> 1 Then
                'OS/2形式だとbiSize=12。カラーマップが各3byteですが、対応しません。
#If DEBUG_ > 0 Then
                Debug.Print "情報ヘッダがWindows形式ではありません。(" & Path & ")"
#End If
                ReadBMPFile = ERR_WINDOWS
                GoTo InputError
            End If

            If .biCompression <> 0 Then
#If DEBUG_ > 0 Then
                Debug.Print "圧縮形式のBMPには対応していません。(" & Path & ")"
#End If
                ReadBMPFile = ERR_COMPRESS
                GoTo InputError
            End If

            If .biClrUsed = 0 Then
                Select Case .biBitCount
                Case 1: ClrUsed = 2
                Case 4: ClrUsed = 16
                Case 8: ClrUsed = 256
                Case Else: ClrUsed = 0
                End Select

            Else
                ClrUsed = .biClrUsed
            End If
        
            'ビット数と画像幅からバイト単位画面幅を計算
            GetByteWidth bWidth, bPadding, .biBitCount, .biWidth
            bWidth = bWidth + bPadding
        End With

        'カラーマップ
        If ClrUsed > 0 Then
            ReDim .colorMap(0 To ClrUsed - 1)
            Get #fileNo, Len(.fileHedaer) + .infoHeader.biSize + 1, .colorMap
        End If

        'ビットマップ情報のサイズ定義
        ReDim .Body(0 To bWidth * .infoHeader.biHeight - 1)

        'ビットマップ情報
        Get #fileNo, .fileHedaer.bfOffBits + 1, .Body
    End With

    ReadBMPFile = NML_SUCCESS
InputError:
    Close #fileNo
End Function

'''カラーマップの作成
''' @param pData   / O / カラーマップ
''' @param bmpInfo / I / ビットマップ情報
Private Sub CreateColorMap(ByRef pData() As Long, ByRef bmpInfo As tBITMAPINFO)
    Dim bWidth As Long, bPadding As Long
    Dim rIdx As Long, cIdx As Long, bIdx As Long, lIdx As Long
    Dim word As Byte

    With bmpInfo.infoHeader
        '24bitカラーマップの高さと幅を確定
        ReDim pData(1 To .biHeight, 1 To .biWidth)
        
        'ビット数と画像幅からバイト単位画面幅を計算
        GetByteWidth bWidth, bPadding, .biBitCount, .biWidth
        bWidth = bWidth + bPadding
    End With

    For rIdx = LBound(pData, 1) To UBound(pData, 1)
        bIdx = (UBound(pData, 1) - rIdx) * bWidth
        For cIdx = LBound(pData, 2) To UBound(pData, 2)
            With bmpInfo
                Select Case .infoHeader.biBitCount
                Case 1
                    If (cIdx - 1) Mod 8 = 0 Then
                        word = .Body(bIdx)
                        bIdx = bIdx + 1
                    End If
                    lIdx = (8 - (cIdx Mod 8)) Mod 8
                    lIdx = word / 2 ^ lIdx
                    lIdx = lIdx And 1

                Case 4
                    If (cIdx - 1) Mod 2 = 0 Then
                        word = .Body(bIdx)
                        bIdx = bIdx + 1
                        lIdx = (word / &H10) And &HF

                    Else
                        lIdx = word And &HF
                    End If

                Case 8
                    word = .Body(bIdx)
                    bIdx = bIdx + 1
                    lIdx = word And &HFF

                Case 24
                    pData(rIdx, cIdx) = RGB(.Body(bIdx + 2), .Body(bIdx + 1), .Body(bIdx))
                    bIdx = bIdx + 3
                End Select

                Select Case .infoHeader.biBitCount
                Case 1, 4, 8
                    With .colorMap(lIdx)
                        pData(rIdx, cIdx) = RGB(.rgbRed, .rgbGreen, .rgbBlue)
                    End With
                End Select
            End With
        Next cIdx
    Next rIdx
End Sub

'''ビット数と画像幅からバイト単位画面幅とパディングを計算
''' @param bWidth   / O / バイト単位画面幅
''' @param bPadding / O / バイト単位パディング
''' @param BitCount / I / ビット数
''' @param Width    / I / ピクセル単位画像幅
Private Sub GetByteWidth(ByRef bWidth As Long, ByRef bPadding As Long, ByVal BitCount As Long, ByVal Width As Long)
    Select Case BitCount
    Case 1:     bWidth = (Width + 7) \ 8
    Case 4:     bWidth = (Width + 1) \ 2
    Case 8:     bWidth = Width
    Case 24:    bWidth = Width * 3
    End Select
    bPadding = (4 - bWidth Mod 4) Mod 4
End Sub

'''ExportBMPFile試験
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
    ExportBMPFile Path & "\test1.bmp", data, colors, 96

    '4bitビットマップ試験
    ReDim data(0 To 99, 0 To 99)
    For i = 0 To 99: For j = 0 To 99: data(i, j) = (i + j) Mod 8: Next j, i
    colors = Array(vbBlack, vbBlue, vbRed, vbMagenta, vbCyan, vbYellow, vbGreen, vbWhite)
    ExportBMPFile Path & "\test4.bmp", data, colors, 96

    '8bitビットマップ試験
    ReDim colors(0 To 255)
    For i = 0 To 255: colors(i) = RGB(i, i, i): Next i
    ReDim data(0 To 255, 0 To 255)
    For i = 0 To 255: For j = 0 To 255: data(i, j) = (i + j) Mod 256: Next j, i
    ExportBMPFile Path & "\test8.bmp", data, colors, 96

    '24bitビットマップ試験
    ReDim data(0 To 127, 0 To 127)
    For i = 0 To 127: For j = 0 To 127: data(i, j) = RGB(i * 2, j * 2, ((i + j) * 2) Mod 256): Next j, i
    ExportBMPFile Path & "\test24.bmp", data, , 96
End Sub

'''ImportBMPFile試験
Private Sub Test_ImportBMPFile()
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim Path As String
    Dim data() As Long
    Dim errCd As eErrorCode
    Dim i%, j%, v

    Set wb = Workbooks.Add
    Set sh = Nothing

    Path = ThisWorkbook.Path

    For Each v In Array("test1", "test4", "test8", "test24")
        If ImportBMPFile(Path & "\" & v & ".bmp", data, errCd) Then
            If sh Is Nothing Then
                Set sh = wb.Worksheets(1)

            Else
                Set sh = wb.Worksheets.Add(After:=sh)
            End If
            ActiveWindow.Zoom = 30
            sh.Name = v
            sh.Cells.RowHeight = 10
            sh.Cells.ColumnWidth = 1
            For i = LBound(data, 1) To UBound(data, 1)
                For j = LBound(data, 2) To UBound(data, 2)
                    sh.Cells(i, j).Interior.Color = data(i, j)
                Next j
                DoEvents
            Next i

        Else
            Debug.Print v, errCd
            Exit Sub
        End If
    Next v
End Sub
