Attribute VB_Name = "QRCodeModule"
Rem このQRCodeModuleは、以下のバーコード生成ライブラリからQRコードの生成に必要な分を抜き出し、
Rem 修正したものです。
Rem オリジナルとは以下の点が異なります。
Rem ・漢字モードに対応しました。
Rem ・UTF-8の3byte以上の文字が指定されたとき、正常に動作するようにしました。
Rem ・オリジナルは出力先がShapeでしたが、2次元配列出力としてあります。
Rem ・解析のため、大きなサブルーチンを分解しました。
Rem ・解析のため、ユーザー定義を使用するようにしました。
Rem ・文字列解析を作り直しました。
Rem 　オリジナルは数字のみの入力は英数字モードで出力しましたが、数字モードで出力します。
Rem 　オリジナルはモード別ブロックが最大20でしたが、無制限としました。
Rem ----
Rem https://code.google.com/archive/p/barcode-vba-macro-only/downloads
Rem  *****  BASIC  *****
Rem This software is distributd under The MIT License (MIT)
Rem Copyright ｩ 2013 Madeta a.s. Jiri Gabriel
Rem Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
Rem The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
Rem THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
Rem
Option Explicit
Option Private Module

Private Const QR_ALNUM$ = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ $%*+-./:"
Private Const CC_SIZ_BLKS$ = _
          "D01A01K01G01" & "J01D01V01P01" & "T01I01P02L02" & "L02N01J04T02" _
        & "R02T01P04L04" & "J04L02V04R04" & "L04N02T05L06" & "P04R02T06P06" _
        & "P05X02R08N08" & "T05L04V08R08" & "X05N04R11V08" & "P08R04V11T10" _
        & "P09T04P16R12" & "R09X04R16N16" & "R10P06R18X12" & "V10R06X16R17" _
        & "V11V06V19V16" & "T13X06V21V18" & "T14V07T25T21" & "T16V08V25X20" _
        & "T17V08X25V23" & "V17V09R34X23" & "V18X09X30X25" & "V20X10X32X27" _
        & "V21T12X35X29" & "V23V12X37V34" & "V25X12X40X34" & "V26X13X42X35" _
        & "V28X14X45X38" & "V29X15X48X40" & "V31X16X51X43" & "V33X17X54X45" _
        & "V35X18X57X48" & "V37X19X60X51" & "V38X19X63X53" & "V40X20X66X56" _
        & "V43X21X70X59" & "V45X22X74X62" & "V47X24X77X65" & "V49X25X81X68"

Public Enum eErrorCorrectionLevel
    ECL_M = 0
    ECL_L = 1
    ECL_H = 2
    ECL_Q = 3
End Enum

Private Enum eType
    TYP_UNKNOWN = 0
    TYP_NUM = 1
    TYP_ALNUM = 2
    TYP_BYTE = 3
    TYP_KANJI = 4
End Enum

Private Type tEcxItem
    Pos As Integer
    Cnt As Integer
End Type

Private Type tEbItem
    Typ As eType
    Pos As Integer
    Cnt As Integer
End Type

Private Type tParams
    Ver As Integer
    Siz As Integer
    ccSiz As Integer
    ccBlks As Integer
    ttlByt As Integer
    syncs() As Integer
    verInf As Long
End Type

Private IsMs As Boolean
Private ErrTxt As String
Public COUNT_LENGTH() As Variant

Private Sub Init()
    IsMs = (VarType(Asc("A")) = vbInteger)
    ErrTxt = ""
    COUNT_LENGTH = [{10, 12, 14; 9, 11, 13; 8, 16, 16; 8, 10, 12}]
End Sub

Public Function GetQRCode(ByRef ar() As Variant, ByRef pInfo As String, ByVal pTarget As String, Optional ByVal pECL As eErrorCorrectionLevel = ECL_M) As Boolean
    Dim barCode As String

    GetQRCode = False

    Init

    barCode = QR_gen(pTarget, pECL)
    If barCode = "" Then
        pInfo = ErrTxt
        Exit Function
    End If

    If Not BC_to2Dim(barCode, ar) Then
        pInfo = ErrTxt
        Exit Function
    End If
    
    pInfo = "Ver." & Int((UBound(ar, 2) - 17) / 4)

    GetQRCode = True

End Function

Public Function CheckQRCode(ByVal pText As String, ByVal pECL As eErrorCorrectionLevel) As Integer
    Dim eb() As tEbItem
    Dim Ver As Integer
    Dim dummy1%, dummy2&, dummy3%, dummy4%

    Init

    QR_anlyz pText, eb

    QR_search_params pECL, eb, Ver, dummy1, dummy2, dummy3, dummy4

    CheckQRCode = Ver
End Function

'指定された文字のUTF-8のバイトコードをLong型で返却する
Private Function AscL(ByVal ch As String) As Long
    If IsMs Then
        AscL = AscW(ch)
        If AscL < 0 Then AscL = AscL + &H10000
    Else
        AscL = Asc(ch)
    End If
End Function

'指定された文字がSJISとして表現可能である場合、QRコード用の13bitのバイトコードに変換して返却する。
'SJISとして表現できない場合、-1を返却する
Private Function QR_AscK(ByVal ch As String) As Long
    Dim code As Long

    QR_AscK = -1
    If Not IsMs Then Exit Function

    'AscはASCIIコードまたはSJISコードをInteger型で返却する。
    'SJISに解釈できない場合、"?"のアスキーコードを返却する。
    code = Asc(ch)
    If code < 0 Then code = code + &H10000

    If code >= &H8140& And code <= &H9FFC& Then
        code = code - &H8140&

    ElseIf code >= &HE040& And code <= &HEBBF& Then
        code = code - &HC140&

    Else
        Exit Function
    End If

    QR_AscK = Int(code / &H100&) * &HC0& + (code And &HFF&)
End Function

Private Sub outErr(ByVal msg As String)
    If ErrTxt <> "" Then ErrTxt = ErrTxt & vbLf
    ErrTxt = ErrTxt & msg
    Debug.Print msg
End Sub

Private Function QR_gen(ByVal pText As String, ByVal pECL As eErrorCorrectionLevel) As String
    Dim encoded1() As Byte
    Dim eb() As tEbItem
    Dim qrArr() As Byte
    Dim qrParam As tParams

    If pText = "" Then
        outErr "QR_gen : no data"
        Exit Function
    End If

    QR_anlyz pText, eb

    QR_params pECL, eb, qrParam

#If DEBUG_ = 1 Then
    QR_debugEb pText, eb, qrParam
#End If

    If qrParam.Ver <= 0 Then
        outErr "QR_gen : too long"
        Exit Function
    End If

#If DEBUG_ = 1 Then
    Debug.Print "ver:" & qrParam.Ver & "  Siz:" & qrParam.Siz & "  ccs:" & qrParam.ccSiz & "  ccb:" & qrParam.ccBlks _
              & "  d:" & (qrParam.ttlByt - qrParam.ccSiz * qrParam.ccBlks)
#End If

    QR_encd pText, qrParam, eb, encoded1

    QR_makeArr qrParam, encoded1, qrArr

    QR_maskArr qrParam, pECL, qrArr

    QR_gen = QR_makeAscMtrx(qrParam, qrArr)
End Function

Private Sub QR_anlyz(ByVal pText As String, ByRef eb() As tEbItem)
    Dim idx As Long
    Dim ch As String
    Dim utfCd As Long   'UTF code
    Dim sjisCd As Long  'SJIS code
    Dim anIdx As Long   'ALNUM index
    Dim ltSiz As Long   'letter size
    Dim nxTyp As eType
    Dim ecx(TYP_NUM To TYP_KANJI) As tEcxItem
    Dim bytSiz() As Integer

    For idx = LBound(ecx) To UBound(ecx): With ecx(idx)
        .Cnt = 0: .Pos = 0
    End With: Next idx

    ReDim bytSiz(1 To Len(pText))
    For idx = 1 To Len(pText)
        ch = Mid(pText, idx, 1)
        utfCd = AscL(ch)
        If utfCd >= &H1FFFFF Then
            bytSiz(idx) = 4

        ElseIf utfCd >= &H7FF Then
            bytSiz(idx) = 3

        ElseIf utfCd >= &H80 Then
            bytSiz(idx) = 2

        Else
            bytSiz(idx) = 1
        End If
    Next idx

    For idx = 1 To Len(pText)
        ch = Mid(pText, idx, 1)
        ltSiz = bytSiz(idx)
        nxTyp = TYP_BYTE
        Select Case ltSiz
        Case 1
            anIdx = InStr(QR_ALNUM, ch) - 1
            sjisCd = -1
            If anIdx >= 0 Then
                nxTyp = TYP_NUM
                If anIdx > 9 Then nxTyp = TYP_ALNUM
            End If

        Case 2
            anIdx = -1
            sjisCd = -1

        Case Else
            anIdx = -1
            sjisCd = QR_AscK(ch)
            If sjisCd >= 0 Then nxTyp = TYP_KANJI
        End Select

        If ecx(TYP_NUM).Cnt > 0 And (anIdx < 0 Or anIdx > 9) Then
            GoSub FIN_NUMERIC

        ElseIf ecx(TYP_ALNUM).Cnt > 0 And anIdx < 0 Then
            GoSub FIN_ALPH_NUMERIC

        ElseIf ecx(TYP_KANJI).Cnt > 0 And sjisCd < 0 Then
            GoSub FIN_KANJI

        End If

        QR_anlyz_cntLtr ecx, TYP_NUM, idx, (anIdx < 0 Or anIdx > 9)
        QR_anlyz_cntLtr ecx, TYP_ALNUM, idx, (anIdx < 0)
        QR_anlyz_cntLtr ecx, TYP_BYTE, idx, False, ltSiz
        QR_anlyz_cntLtr ecx, TYP_KANJI, idx, (sjisCd < 0)
    Next idx

    nxTyp = TYP_UNKNOWN
    If ecx(TYP_NUM).Cnt > 0 Then GoSub FIN_NUMERIC
    If ecx(TYP_ALNUM).Cnt > 0 Then GoSub FIN_ALPH_NUMERIC
    If ecx(TYP_KANJI).Cnt > 0 Then GoSub FIN_KANJI
    If ecx(TYP_BYTE).Cnt > 0 Then QR_addEb eb, ecx, TYP_BYTE

    Exit Sub

FIN_NUMERIC:
    '- ALNUMと数字が隣接している場合、数字が8桁未満ならばALNUMとして出力したほうがビット数が少ない
    '- byteと数字が隣接している場合、数字が4桁未満ならばbyteとして出力したほうがビット数が少ない
    '- 数字と漢字はグループが異なる
    If ecx(TYP_NUM).Cnt < 8 Then
        '数字が8桁未満
        If nxTyp = TYP_ALNUM Then
            '次の文字が数字以外のALNUMなのでALNUMとして出力予定
            ecx(TYP_NUM).Cnt = 0
            Return
        End If

        If ecxcmp(ecx, TYP_BYTE, TYP_ALNUM) = 0 And nxTyp = TYP_UNKNOWN Then
            'ALNUMの前後に未出力のbyteが存在しない場合

        ElseIf ecxcmp(ecx, TYP_ALNUM, TYP_NUM) > 0 Then
            '数字の前に数字以外の未出力のALNUMが存在する
            If ecx(TYP_ALNUM).Cnt < 7 Then
                '未出力のALNUMが7桁未満なのでbyteとして出力予定
                ecx(TYP_NUM).Cnt = 0
                ecx(TYP_ALNUM).Cnt = 0
                Return
            End If

        ElseIf ecx(TYP_NUM).Cnt < 4 Then
            '数字の前後がALNUM以外のbyteで、数字が4桁未満の場合、byteとして出力予定
            ecx(TYP_NUM).Cnt = 0
            ecx(TYP_ALNUM).Cnt = 0
            Return
        End If
    End If

    If ecxcmp(ecx, TYP_BYTE, TYP_ALNUM) > 0 Then
        'ALNUM前に未出力のbyteが存在する場合それを出力
        QR_addEb eb, ecx, TYP_BYTE, TYP_ALNUM
    End If

    If ecxcmp(ecx, TYP_ALNUM, TYP_NUM) = 0 Then
        'ALNUMと数字の未出力桁数が同じ場合は数字として出力
        QR_addEb eb, ecx, TYP_NUM

    ElseIf ecx(TYP_NUM).Cnt < 8 Then
        '数字の前に未出力のALNUMが存在し、数字が8桁未満の場合
        QR_addEb eb, ecx, TYP_ALNUM

    Else
        '数字の前に未出力のALNUMが存在し、数字が8桁以上の場合
        QR_addEb eb, ecx, TYP_ALNUM, TYP_NUM
        QR_addEb eb, ecx, TYP_NUM
    End If

    ecx(TYP_NUM).Cnt = 0
    ecx(TYP_ALNUM).Cnt = 0
    ecx(TYP_BYTE).Cnt = 0

    Return

FIN_ALPH_NUMERIC:
    '- byteとALNUMが隣接している場合、ALNUMが7桁未満ならばbyteとして出力したほうがビット数が少ない
    '- ALNUMと漢字はグループが異なる
    If ecxcmp(ecx, TYP_BYTE, TYP_ALNUM) = 0 And nxTyp = TYP_UNKNOWN Then
        'ALNUMの前後に未出力のbyteが存在しない場合

    ElseIf ecx(TYP_ALNUM).Cnt < 7 Then
        'ALNUMが7文字未満なのでbyteとして出力予定
        ecx(TYP_ALNUM).Cnt = 0
        Return
    End If

    If ecxcmp(ecx, TYP_BYTE, TYP_ALNUM) > 0 Then
        QR_addEb eb, ecx, TYP_BYTE, TYP_ALNUM
    End If
    QR_addEb eb, ecx, TYP_ALNUM

    ecx(TYP_ALNUM).Cnt = 0
    ecx(TYP_BYTE).Cnt = 0

    Return

FIN_KANJI:
    '・byteと漢字が隣接している場合、漢字が1文字ならばbyteとして出力したほうがビット数が少ない
    '　が、UTF-8とSJISが混在すると文字化けするリーダーがあるので出力する

    'If ecxcmp(ecx, TYP_BYTE, TYP_KANJI) = 0 And nxTyp = TYP_UNKNOWN Then
    '    '漢字の前後に未出力のbyteが存在しない場合
    '
    'ElseIf ecx(TYP_KANJI).Cnt = 1 Then
    '    ecx(TYP_KANJI).Cnt = 0
    '    Return
    'End If

    If ecxcmp(ecx, TYP_BYTE, TYP_KANJI) > 0 Then
        QR_addEb eb, ecx, TYP_BYTE, TYP_KANJI, bytSiz
    End If

    QR_addEb eb, ecx, TYP_KANJI

    ecx(TYP_KANJI).Cnt = 0
    ecx(TYP_BYTE).Cnt = 0

    Return
End Sub

Private Function ecxcmp(ByRef ecx() As tEcxItem, ByVal sTyp As eType, ByVal dTyp As eType) As Integer
    If sTyp = TYP_BYTE And dTyp = TYP_KANJI Then
        ecxcmp = ecx(dTyp).Pos - ecx(sTyp).Pos

    ElseIf sTyp = TYP_KANJI And dTyp = TYP_BYTE Then
        ecxcmp = ecx(dTyp).Pos - ecx(sTyp).Pos

    Else
        ecxcmp = ecx(sTyp).Cnt - ecx(dTyp).Cnt
    End If
End Function

Private Sub QR_anlyz_cntLtr(ByRef ecx() As tEcxItem, ByVal tTyp As eType, ByVal idx As Integer, ByVal reset As Boolean, Optional ByVal ltSiz As Integer = 1)
    If reset Then
        ecx(tTyp).Cnt = 0
        Exit Sub
    End If

    If ecx(tTyp).Cnt = 0 Then ecx(tTyp).Pos = idx
    ecx(tTyp).Cnt = ecx(tTyp).Cnt + ltSiz
End Sub

Private Sub QR_addEb(ByRef eb() As tEbItem, ByRef ecx() As tEcxItem, ByVal tTyp As eType, Optional ByVal sTyp As eType = TYP_UNKNOWN, Optional ByRef bytSiz As Variant)
    Dim idx As Integer

    If tTyp = TYP_UNKNOWN Then
        Exit Sub
    End If

    If (Not eb) = -1 Then
        idx = 1

    Else
        idx = UBound(eb) + 1
    End If

    ReDim Preserve eb(1 To idx)
    With eb(idx)
        .Typ = tTyp
        .Pos = ecx(tTyp).Pos
        If sTyp = TYP_UNKNOWN Then
            .Cnt = ecx(tTyp).Cnt

        ElseIf IsArray(bytSiz) Then
            .Cnt = 0
            For idx = ecx(tTyp).Pos To ecx(sTyp).Pos - 1
                .Cnt = .Cnt + bytSiz(idx)
            Next idx
        
        Else
            .Cnt = ecx(tTyp).Cnt - ecx(sTyp).Cnt
        End If
    End With
End Sub

Private Sub QR_debugEb(ByVal pText As String, ByRef eb() As tEbItem, ByRef qrParam As tParams)
    Dim txt$, idx%, ln%, b&, utfCd&, ttlSiz&, ttlBit&
    txt = ""
    ttlSiz = 0
    For idx = LBound(eb) To UBound(eb): With eb(idx)
        ln = .Cnt
        Select Case .Typ
        Case TYP_UNKNOWN:   txt = "U"
        Case TYP_NUM:       txt = "N"
        Case TYP_ALNUM:     txt = "A"
        Case TYP_KANJI:     txt = "K"
        Case TYP_BYTE:      txt = "B"
            ln = 0: b = 0
            Do While b < .Cnt
                utfCd = AscL(Mid(pText, .Pos + ln, 1))
                b = b + 1 - (utfCd >= &H80) - (utfCd >= &H7FF) - (utfCd >= &H1FFFFF)
                ln = ln + 1
            Loop
        End Select

        b = QR_getBitSize(.Typ, .Cnt, qrParam.Ver)
        ttlBit = ttlBit + b
        ttlSiz = ttlSiz + ln
        txt = "eb(" & idx & ")=" & txt
        txt = txt & "(" & ln & " letters"
        txt = txt & " / " & b & "bits)"
        txt = txt & ", [" & Mid(pText, .Pos, ln) & "]"

        Debug.Print txt
    End With: Next idx

    Debug.Print "Total : " & ttlSiz & " letters  / " & ttlBit & " bits  / " & Int((ttlBit + 7) / 8) & " bytes"

    If ttlSiz <> Len(pText) Then
        outErr "QR_debugEb : eb error  ttlSiz=" & ttlSiz & "  Len(pText)=" & Len(pText)
    End If
End Sub

Private Sub QR_params(ByVal pECL As eErrorCorrectionLevel, ByRef eb() As tEbItem, ByRef qrParam As tParams)
    Dim Siz As Integer
    Dim ttlByt As Long
    Dim idx As Integer
    Dim t As Integer
    Dim verInf As Long
    Dim syncSiz As Integer
    Dim ccSiz As Integer
    Dim ccBlks As Integer
    Dim maxSiz(0 To 2) As Long
    Dim Ver As Integer

    'init
    With qrParam
        .Ver = 0
        .Siz = 0
        .ccSiz = 0
        .ccBlks = 0
        .ttlByt = 0
        Erase .syncs
        .verInf = 0
    End With

    QR_search_params pECL, eb, Ver, Siz, ttlByt, ccSiz, ccBlks
    If Ver = 0 Then Exit Sub

    With qrParam
        If Ver > 1 Then
            syncSiz = Int(Ver / 7) + 1
            ReDim .syncs(0 To syncSiz)
            .syncs(0) = 6
            .syncs(syncSiz) = Siz - 7
            If syncSiz >= 2 Then
                t = Int((Siz - 13) / 2 / syncSiz + 0.7) * 2
                .syncs(1) = .syncs(syncSiz) - t * (syncSiz - 1)
                If syncSiz >= 3 Then
                    For idx = 3 To syncSiz
                        .syncs(idx - 1) = .syncs(idx - 2) + t
                    Next idx
                End If
            End If
        End If
        .Ver = Ver
        .Siz = Siz
        .ccSiz = ccSiz
        .ccBlks = ccBlks
        .ttlByt = ttlByt
        If Ver >= 7 Then
            verInf = Ver
            QR_bch_calc verInf, &H1F25
            .verInf = verInf
        End If
    End With
End Sub

Private Sub QR_search_params(ByVal pECL As eErrorCorrectionLevel, ByRef eb() As tEbItem _
                           , ByRef Ver As Integer, ByRef Siz As Integer, ByRef ttlByt As Long _
                           , ByRef ccSiz As Integer, ByRef ccBlks As Integer)
    Dim txt As String
    Dim idx As Integer, syncs As Integer
    Dim reqSiz(0 To 2) As Long

    Ver = 0

    reqSiz(0) = 0
    reqSiz(1) = 0
    reqSiz(2) = 0
    For idx = LBound(eb) To UBound(eb): With eb(idx)
        reqSiz(0) = reqSiz(0) + QR_getBitSize(.Typ, .Cnt, 1)
        reqSiz(1) = reqSiz(1) + QR_getBitSize(.Typ, .Cnt, 10)
        reqSiz(2) = reqSiz(2) + QR_getBitSize(.Typ, .Cnt, 27)
    End With: Next idx
    reqSiz(0) = Int((reqSiz(0) + 7) / 8)
    reqSiz(1) = Int((reqSiz(1) + 7) / 8)
    reqSiz(2) = Int((reqSiz(2) + 7) / 8)

    If pECL = ECL_M And reqSiz(2) > 2334 _
    Or pECL = ECL_L And reqSiz(2) > 2956 _
    Or pECL = ECL_H And reqSiz(2) > 1276 _
    Or pECL = ECL_Q And reqSiz(2) > 1666 Then
        Exit Sub
    End If

    For Ver = 1 To 40
        Siz = 4 * Ver + 17
        txt = Mid(CC_SIZ_BLKS, (Ver - 1) * 12 + pECL * 3 + 1, 3)
        ccSiz = Asc(Left(txt, 1)) - 65 + 7
        ccBlks = Val(Mid(txt, 2))
        If Ver = 1 Then
            ttlByt = 26

        Else
            syncs = ((Int(Ver / 7) + 2) ^ 2) - 3
            ttlByt = (((Siz - 1) ^ 2) / 8) - (3& * syncs) - 24
            If Ver > 6 Then ttlByt = ttlByt - 4
            If syncs = 1 Then ttlByt = ttlByt - 1
        End If

        Select Case Ver
        Case 1 To 9:    idx = 0
        Case 10 To 26:  idx = 1
        Case 27 To 40:  idx = 2
        End Select
        If ttlByt >= reqSiz(idx) + ccSiz * ccBlks Then
            Exit Sub
        End If
    Next Ver

    Ver = 0
End Sub

Private Function QR_getBitSize(ByVal pTyp As eType, ByVal pCnt As Integer, ByVal pVer As Integer) As Long
    Dim bitSiz As Long

    bitSiz = 4 + QR_getCntLen(pTyp, pVer)

    Select Case pTyp
    Case TYP_NUM
        bitSiz = bitSiz + Int(pCnt / 3) * 10
        bitSiz = bitSiz + (pCnt Mod 3) * 3
        If pCnt Mod 3 > 0 Then bitSiz = bitSiz + 1

    Case TYP_ALNUM
        bitSiz = bitSiz + Int(pCnt / 2) * 11
        bitSiz = bitSiz + (pCnt Mod 2) * 6

    Case TYP_BYTE
        bitSiz = bitSiz + pCnt * 8

    Case TYP_KANJI
        bitSiz = bitSiz + pCnt * 13
    End Select

    QR_getBitSize = bitSiz
End Function

Private Function QR_getCntLen(ByVal pTyp As eType, ByVal pVer As Integer) As Integer
    If pVer < 10 Then
        QR_getCntLen = COUNT_LENGTH(pTyp, 1)

    ElseIf pVer < 27 Then
        QR_getCntLen = COUNT_LENGTH(pTyp, 2)

    Else
        QR_getCntLen = COUNT_LENGTH(pTyp, 3)
    End If
End Function

Private Sub QR_encd(ByVal pText As String, ByRef qrParam As tParams, ByRef eb() As tEbItem, ByRef encArr() As Byte)
    Dim encIdx As Integer
    Dim r As Integer
    Dim c As Integer
    Dim bits As Long
    Dim bIdx As Long, eIdx As Long, idx As Long
    Dim k As Long
    Dim m As Long

    ReDim encArr(qrParam.ttlByt + 2)

    encIdx = 0
    For eIdx = LBound(eb) To UBound(eb): With eb(eIdx)
        Select Case .Typ
        Case TYP_NUM:   k = 1
        Case TYP_ALNUM: k = 2
        Case TYP_BYTE:  k = 4
        Case TYP_KANJI: k = 8
        End Select

        c = QR_getCntLen(.Typ, qrParam.Ver)
        k = k * (2 ^ c) + .Cnt

        BB_putBits encArr, encIdx, k, c + 4

        bIdx = 0
        m = .Pos
        r = 0
        Do While bIdx < .Cnt
            k = AscL(Mid(pText, m, 1))
            m = m + 1

            Select Case .Typ
            Case TYP_NUM
                r = (r * 10) + ((k - &H30) Mod 10)
                If bIdx Mod 3 = 2 Then
                    BB_putBits encArr, encIdx, r, 10
                    r = 0
                End If
                bIdx = bIdx + 1

            Case TYP_ALNUM
                r = (r * 45) + ((InStr(QR_ALNUM, Chr(k)) - 1) Mod 45)
                If bIdx Mod 2 = 1 Then
                    BB_putBits encArr, encIdx, r, 11
                    r = 0
                End If
                bIdx = bIdx + 1

            Case TYP_KANJI
                bits = QR_AscK(Mid(pText, m - 1, 1))
                BB_putBits encArr, encIdx, bits, 13
                bIdx = bIdx + 1

            Case TYP_BYTE
                If k > &H1FFFFF Then
                    bits = &HF0 + Int(k / &H40000) Mod 8
                    BB_putBits encArr, encIdx, bits, 8

                    bits = &H80 + Int(k / &H1000) Mod &H40
                    BB_putBits encArr, encIdx, bits, 8

                    bits = &H80 + Int(k / &H40) Mod &H40
                    BB_putBits encArr, encIdx, bits, 8

                    bits = &H80 + k Mod &H40
                    BB_putBits encArr, encIdx, bits, 8
                    bIdx = bIdx + 4

                ElseIf k > &H7FF Then
                    bits = &HE0 + Int(k / &H1000) Mod &H10
                    BB_putBits encArr, encIdx, bits, 8

                    bits = &H80 + Int(k / &H40) Mod &H40
                    BB_putBits encArr, encIdx, bits, 8

                    bits = &H80 + k Mod &H40
                    BB_putBits encArr, encIdx, bits, 8
                    bIdx = bIdx + 3

                ElseIf k > &H7F Then
                    bits = &HC0 + Int(k / &H40) Mod &H20
                    BB_putBits encArr, encIdx, bits, 8

                    bits = &H80 + k Mod &H40
                    BB_putBits encArr, encIdx, bits, 8
                    bIdx = bIdx + 2

                Else
                    bits = k Mod &H100
                    BB_putBits encArr, encIdx, bits, 8
                    bIdx = bIdx + 1
                End If
            End Select
        Loop

        Select Case .Typ
        Case TYP_NUM
            If (bIdx Mod 3) = 1 Then
                BB_putBits encArr, encIdx, r, 4

            ElseIf (bIdx Mod 3) = 2 Then
                BB_putBits encArr, encIdx, r, 7
            End If

        Case TYP_ALNUM
            If (bIdx Mod 2) = 1 Then
                BB_putBits encArr, encIdx, r, 6
            End If
        End Select
    End With: Next eIdx

    'end of chain
    BB_putBits encArr, encIdx, 0, 4
    If (encIdx Mod 8) <> 0 Then
        BB_putBits encArr, encIdx, 0, 8 - (encIdx Mod 8)
    End If

    'padding
    idx = (qrParam.ttlByt - qrParam.ccSiz * qrParam.ccBlks) * 8
    If encIdx > idx Then
        outErr "QR_encd : encode length error"

        Exit Sub
    End If

    'padding 0xEC, 0x11, 0xEC, 0x11
    Do While encIdx < idx
        BB_putBits encArr, encIdx, &HEC11, 16
    Loop

    'supplement ECC
    idx = qrParam.ccSiz * qrParam.ccBlks
    QR_rs &H11D, encArr, qrParam.ttlByt - idx, idx, qrParam.ccBlks
End Sub

'read solomon
Private Sub QR_rs(pPoly As Integer, encArr() As Byte, ByVal pSize As Integer, ByVal pLen As Integer, ByVal pBlocks As Integer)
    Dim v_x As Integer, v_y As Integer, v_z As Integer, v_a As Integer, v_b As Integer
    Dim pA As Integer, rp As Integer
    Dim v_bs As Integer, v_b2c As Integer
    Dim vpo As Integer
    Dim vdo As Integer
    Dim v_es As Integer
    Dim poly(512) As Byte
    Dim v_ply() As Byte

    'generate read solomon expTable and logTable
    ' QR uses GF256(0x11d) // 0x11d = 285 = x^8 + x^4 + x^3 + x^2 + 1
    v_x = 1
    For v_y = 0 To 255
        poly(v_y) = v_x
        poly(v_x + 256) = v_y
        v_x = v_x * 2
        If v_x > 255 Then
            v_x = v_x Xor pPoly
        End If
    Next v_y

    For v_x = 1 To pLen
        encArr(v_x + pSize) = 0
    Next v_x

    v_b2c = pBlocks
    v_bs = Int(pSize / pBlocks) 'minimum block size
    v_es = Int(pLen / pBlocks)  'ecc block size
    v_x = pSize Mod pBlocks     'remain bytes
    v_b2c = pBlocks - v_x       'on block number

    ReDim v_ply(v_es + 1)
    v_z = 0
    v_ply(1) = 1
    v_x = 2
    Do While v_x <= v_es + 1
        v_ply(v_x) = v_ply(v_x - 1)
        v_y = v_x - 1
        Do While v_y > 1
            rp = QR_rs_prod(poly, v_ply(v_y), poly(v_z))

            v_ply(v_y) = v_ply(v_y - 1) Xor rp
            v_y = v_y - 1
        Loop

        rp = QR_rs_prod(poly, v_ply(1), poly(v_z))

        v_ply(1) = rp
        v_z = v_z + 1
        v_x = v_x + 1
    Loop

    For v_b = 0 To pBlocks - 1
        vpo = v_b * v_es + 1 + pSize    ' ECC start
        vdo = v_b * v_bs + 1            ' data start
        If v_b > v_b2c Then
            ' x longers before
            vdo = vdo + v_b - v_b2c
        End If

        'generate "nc" check words in the array
        v_x = 0
        v_z = v_bs
        If v_b >= v_b2c Then v_z = v_z + 1

        Do While v_x < v_z
            pA = encArr(vpo) Xor encArr(vdo + v_x)
            v_y = vpo
            v_a = v_es
            Do While v_a > 0
                rp = QR_rs_prod(poly, pA, v_ply(v_a))

                If v_a = 1 Then
                    encArr(v_y) = rp

                Else
                    encArr(v_y) = encArr(v_y + 1) Xor rp
                End If
                v_y = v_y + 1
                v_a = v_a - 1
            Loop

            v_x = v_x + 1
        Loop
    Next v_b
End Sub

Private Function QR_rs_prod(ByRef poly() As Byte, ByVal pA As Integer, ByVal pB As Integer) As Integer
    QR_rs_prod = 0
    If pA > 0 And pB > 0 Then
        QR_rs_prod = poly((0& + poly(256 + pA) + poly(256 + pB)) Mod 255&)
    End If
End Function

Private Sub QR_makeArr(ByRef qrParam As tParams, ByRef encArr() As Byte, ByRef qrArr() As Byte)
    Dim ch As Integer
    Dim Siz As Integer
    Dim r As Integer, c As Integer
    Dim idx As Long
    Dim k As Long
    Dim qrSync1(1 To 8) As Byte
    Dim qrSync2(1 To 5) As Byte

    Siz = qrParam.Siz

    ReDim qrArr(0 To 1, 0 To (Siz + 1) * 24&)
    qrArr(0, 0) = 0
    ch = 0

    BB_putBits qrSync1, ch, Array(&HFE, &H82, &HBA, &HBA, &HBA, &H82, &HFE, 0), 64
    QR_mask qrArr, qrSync1, 8, 0, 0
    QR_mask qrArr, 0, 8, 8, 0
    QR_mask qrArr, qrSync1, 8, 0, Siz - 7
    QR_mask qrArr, 0, 8, 8, Siz - 8
    QR_mask qrArr, qrSync1, 8, Siz - 7, 0
    QR_mask qrArr, 0, 8, Siz - 8, 0

    For idx = 0 To 6
        QR_bit qrArr, -1, idx, 8, 0
        QR_bit qrArr, -1, idx, Siz - 8, 0
        QR_bit qrArr, -1, Siz - 1 - idx, 8, 0
    Next idx

    QR_bit qrArr, -1, 7, 8, 0
    QR_bit qrArr, -1, 7, Siz - 8, 0
    QR_bit qrArr, -1, 8, 8, 0
    QR_bit qrArr, -1, Siz - 8, 8, 1

    'version info
    If qrParam.verInf <> 0& Then
        k = qrParam.verInf
        c = 0
        r = 0
        For idx = 0 To 17
            ch = k Mod 2
            QR_bit qrArr, -1, r, Siz - 11 + c, ch
            QR_bit qrArr, -1, Siz - 11 + c, r, ch
            c = c + 1
            If c > 2 Then
                c = 0
                r = r + 1
            End If
            k = Int(k / 2&)
        Next idx
    End If

    'sync line
    c = 1
    For idx = 8 To Siz - 9
        QR_bit qrArr, -1, idx, 6, c
        QR_bit qrArr, -1, 6, idx, c
        c = (c + 1) Mod 2
    Next idx

    'other sync
    ch = 0
    BB_putBits qrSync2, ch, Array(&H1F, &H11, &H15, &H11, &H1F), 40
    With qrParam
        If (Not .syncs) <> -1 Then
            ch = UBound(.syncs)
            For c = 0 To ch
                For r = 0 To ch
                    If (c <> 0 Or r <> 0) And _
                       (c <> ch Or r <> 0) And _
                       (c <> 0 Or r <> ch) Then
                        QR_mask qrArr, qrSync2, 5, .syncs(r) - 2, .syncs(c) - 2
                    End If
                Next r
            Next c
        End If
    End With

    With qrParam
        QR_fill qrArr, Siz, encArr, .ccBlks, .ttlByt - .ccSiz * .ccBlks, .ttlByt
    End With
End Sub

Private Sub QR_mask(ByRef qrArr() As Byte, ByRef pVal As Variant, ByVal pBits As Integer, ByVal pRow As Integer, ByVal pCol As Integer)
    Dim bIdx As Integer
    Dim word As Long
    Dim idx As Integer, rIdx As Integer, cIdx As Integer

    If pBits > 8 Or pBits < 1 Then
        outErr "QR_mask : pBits=" & CStr(pBits) & " is error"
        Exit Sub
    End If
    rIdx = pRow
    cIdx = pCol

    Select Case VarType(pVal)
    Case vbByte, vbInteger, vbLong, vbDouble
        word = Int(pVal)

        GoSub DoMask

    Case Else
        If InStr("Byte(),Integer(),Long(),Variant()", TypeName(pVal)) > 0 Then
            For idx = LBound(pVal) To UBound(pVal)
                cIdx = pCol
                word = Int(pVal(idx))

                GoSub DoMask

                rIdx = rIdx + 1
            Next idx

        Else
            outErr "QR_mask: " & TypeName(pVal) & " Unknown type"
            Exit Sub
        End If
    End Select

    Exit Sub

DoMask:
    bIdx = 2 ^ (pBits - 1)
    Do While bIdx > 0
        QR_bit qrArr, -1, rIdx, cIdx, word And bIdx
        cIdx = cIdx + 1
        bIdx = Int(bIdx / 2)
    Loop

    Return
End Sub

Private Sub QR_fill(ByRef qrArr() As Byte, ByVal pSiz As Integer, ByRef encArr() As Byte, ByVal pBlkSiz As Integer, ByVal pDatLen As Integer, ByVal pTtlLen As Integer)
    Dim vds As Integer, ves As Integer, vDnLen As Integer, vsb As Integer
    Dim vx As Integer, vy As Integer
    Dim vb As Integer
    Dim cIdx As Integer, rIdx As Integer
    Dim wa As Integer, wb As Integer, w As Integer
    Dim smer As Integer

    vds = Int(pDatLen / pBlkSiz)
    ves = Int((pTtlLen - pDatLen) / pBlkSiz)
    vDnLen = vds * pBlkSiz
    vsb = pBlkSiz - (pDatLen Mod pBlkSiz)

    cIdx = pSiz - 1
    rIdx = cIdx
    smer = 0
    vb = 1
    w = encArr(1)
    vx = 0
    Do While cIdx >= 0 And vb <= pTtlLen
        If QR_bit(qrArr, pSiz, rIdx, cIdx, (w And 128)) Then
            vx = vx + 1
            If vx = 8 Then
                GoSub qrf_NextByte
                vx = 0

            Else
                w = (w * 2) Mod 256
            End If
        End If

        Select Case smer
        Case 0, 2
            cIdx = cIdx - 1
            smer = smer + 1

        Case 1
            If rIdx = 0 Then
                cIdx = cIdx - 1
                If cIdx = 6 And pSiz >= 21 Then
                    cIdx = cIdx - 1
                End If
                smer = 2

            Else
                cIdx = cIdx + 1
                rIdx = rIdx - 1
                smer = 0
            End If

        Case 3
            If rIdx = pSiz - 1 Then
                cIdx = cIdx - 1
                If cIdx = 6 And pSiz >= 21 Then
                    cIdx = cIdx - 1
                End If
                smer = 0

            Else
                cIdx = cIdx + 1
                rIdx = rIdx + 1
                smer = 2
            End If
        End Select
    Loop

    Exit Sub

qrf_NextByte:
    If vb < pDatLen Then
        wa = vb
        If vb >= vDnLen Then
            wa = wa + vsb
        End If

        wb = wa Mod pBlkSiz
        wa = Int(wa / pBlkSiz)
        If wb > vsb Then
            wa = wa + wb - vsb
        End If

        w = encArr(1 + vds * wb + wa)

    ElseIf vb < pTtlLen Then
        wa = vb - pDatLen
        wb = wa Mod pBlkSiz
        wa = Int(wa / pBlkSiz)
        w = encArr(1 + pDatLen + ves * wb + wa)
    End If
    vb = vb + 1

    Return
End Sub

Private Sub QR_maskArr(ByRef qrParam As tParams, ByVal pECL As eErrorCorrectionLevel, ByRef qrArr() As Byte)
    Dim Siz As Integer
    Dim mask As Integer, idx As Integer
    Dim score As Long, minScore As Long

    Siz = qrParam.Siz

    minScore = -1
    For idx = 0 To 7
        QR_maskArr_addMM qrArr, pECL, Siz, idx
        score = QR_xorMask(qrArr, Siz, idx, False)
        If score < minScore Or minScore = -1 Then
            minScore = score
            mask = idx
        End If
    Next idx

    QR_maskArr_addMM qrArr, pECL, Siz, mask
    score = QR_xorMask(qrArr, Siz, mask, True)
End Sub

Private Sub QR_maskArr_addMM(ByRef qrArr() As Byte, ByVal pECL As eErrorCorrectionLevel, ByVal Siz As Integer, ByVal mask As Integer)
    Dim idx As Long
    Dim k As Long
    Dim ch As Integer
    Dim rIdx As Integer, cIdx As Integer

    k = pECL * 8 + mask
    QR_bch_calc k, &H537
    k = k Xor &H5412

    rIdx = 0
    cIdx = Siz - 1
    For idx = 0 To 14
        ch = k Mod 2
        k = Int(k / 2)
        QR_bit qrArr, -1, rIdx, 8, ch
        QR_bit qrArr, -1, 8, cIdx, ch
        cIdx = cIdx - 1
        rIdx = rIdx + 1
        If idx = 7 Then cIdx = 7: rIdx = Siz - 7
        If idx = 5 Then rIdx = rIdx + 1
        If idx = 8 Then cIdx = cIdx - 1
    Next idx
End Sub

'pMod = 0b000 : (c + r) Mod 2 = 0
'pMod = 0b001 : r Mod 2 = 0
'pMod = 0b010 : c Mod 3 = 0
'pMod = 0b011 : (c + r) Mod 3 = 0
'pMod = 0b100 : (Int(r / 2) + Int(c / 3)) Mod 2 = 0
'pMod = 0b101 : (c * r) Mod 2 + (c * r) Mod 3 = 0
'pMod = 0b110 : ((c * r) Mod 2 + (c * r) Mod 3) Mod 2 = 0
'pMod = 0b111 : ((c + r) Mod 2 + (c * r) Mod 3) Mod 2 = 0
Private Function QR_xorMask(ByRef qrArr() As Byte, ByVal pSiz As Integer, ByVal pMod As Integer, ByVal pFinal As Boolean) As Long
    Dim cIdx As Integer, rIdx As Integer
    Dim idx As Integer
    Dim i As Integer
    Dim m As Integer
    Dim wArr() As Byte

    ReDim wArr(pSiz * 24)

    For rIdx = 0 To pSiz - 1
        m = 1
        idx = 24 * rIdx
        wArr(idx) = qrArr(1, idx)

        For cIdx = 0 To pSiz - 1
            If (qrArr(0, idx) And m) = 0 Then
                Select Case pMod
                Case 0: i = (cIdx + rIdx) Mod 2
                Case 1: i = rIdx Mod 2
                Case 2: i = cIdx Mod 3
                Case 3: i = (cIdx + rIdx) Mod 3
                Case 4: i = (Int(rIdx / 2) + Int(cIdx / 3)) Mod 2
                Case 5: i = (cIdx * rIdx) Mod 2 + (cIdx * rIdx) Mod 3
                Case 6: i = ((cIdx * rIdx) Mod 2 + (cIdx * rIdx) Mod 3) Mod 2
                Case 7: i = ((cIdx + rIdx) Mod 2 + (cIdx * rIdx) Mod 3) Mod 2
                End Select

                If i = 0 Then wArr(idx) = wArr(idx) Xor m
            End If

            If m = 128 Then
                m = 1
                If pFinal Then qrArr(1, idx) = wArr(idx)
                idx = idx + 1
                wArr(idx) = qrArr(1, idx)
            Else
                m = m * 2
            End If
        Next cIdx

        If m <> 128 And pFinal Then qrArr(1, idx) = wArr(idx)
    Next rIdx

    If pFinal Then
        QR_xorMask = 0

    Else
        QR_xorMask = QR_xorMask_scoring(qrArr, wArr, pSiz)
    End If
End Function

'score computing
Private Function QR_xorMask_scoring(ByRef qrArr() As Byte, ByRef wArr() As Byte, ByVal pSiz As Integer) As Long
    Dim score As Long
    Dim bl As Long
    Dim rp As Long
    Dim rc As Long
    Dim cIdx As Integer, rIdx As Integer
    Dim idx As Integer
    Dim i As Integer
    Dim m As Integer
    Dim cols() As Long

    score = 0
    bl = 0
    ReDim cols(1, pSiz)
    rp = 0
    rc = 0
    For rIdx = 0 To pSiz - 1
        m = 1
        idx = 24 * rIdx
        rp = 0
        rc = 0
        For cIdx = 0 To pSiz - 1
            rp = (rp And &H3FF) * 2
            cols(1, cIdx) = (cols(1, cIdx) And &H3FF) * 2
            If (wArr(idx) And m) <> 0 Then
                If rc < 0 Then
                    If rc <= -5 Then score = score - 2 - rc
                    rc = 0
                End If
                rc = rc + 1
                If cols(0, cIdx) < 0 Then
                    If cols(0, cIdx) <= -5 Then score = score - 2 - cols(0, cIdx)
                    cols(0, cIdx) = 0
                End If
                cols(0, cIdx) = cols(0, cIdx) + 1
                rp = rp Or 1
                cols(1, cIdx) = cols(1, cIdx) Or 1
                bl = bl + 1

            Else
                If rc > 0 Then
                    If rc >= 5 Then score = score - 2 + rc
                    rc = 0
                End If
                rc = rc - 1
                If cols(0, cIdx) > 0 Then
                    If cols(0, cIdx) >= 5 Then score = score - 2 + cols(0, cIdx)
                    cols(0, cIdx) = 0
                End If
                cols(0, cIdx) = cols(0, cIdx) - 1
            End If

            If cIdx > 0 And rIdx > 0 Then
                i = rp And 3
                If (cols(1, cIdx - 1) And 3) >= 2 Then i = i + 8
                If (cols(1, cIdx) And 3) >= 2 Then i = i + 4
                If i = 0 Or i = 15 Then
                    score = score + 3
                End If
            End If

            If cIdx >= 10 And (rp = &H5D Or rp = &H5D0) Then
                score = score + 40
            End If
            If rIdx >= 10 And (cols(1, cIdx) = &H5D Or cols(1, cIdx) = &H5D0) Then
                score = score + 40
            End If

            If m = 128 Then
                m = 1
                idx = idx + 1
            Else
                m = m * 2
            End If
        Next cIdx
        If Abs(rc) >= 5 Then score = score - 2 + Abs(rc)
    Next rIdx

    For cIdx = 0 To pSiz - 1
        If Abs(cols(0, cIdx)) >= 5 Then score = score - 2 + Abs(cols(0, cIdx))
    Next cIdx
    bl = Int(Abs((bl * 100&) / (pSiz * pSiz) - 50&) / 5) * 10

    QR_xorMask_scoring = score + bl
End Function

Private Function QR_bit(ByRef qrArr() As Byte, ByVal pSiz As Integer, ByVal pRow As Integer, ByVal pCol As Integer, ByVal pBit As Integer)
    Dim idx As Integer
    Dim Value As Integer
    Dim rIdx As Integer, cIdx As Integer

    rIdx = pRow
    cIdx = pCol
    QR_bit = False
    idx = rIdx * 24 + Int(cIdx / 8)
    If idx > UBound(qrArr, 2) Or idx < 0 Then
        outErr "QR_bit : out of range"
        Exit Function
    End If

    cIdx = 2 ^ (cIdx Mod 8)
    Value = qrArr(0, idx)
    If pSiz > 0 Then
        If (Value And cIdx) = 0 Then
            If pBit <> 0 Then
                qrArr(1, idx) = qrArr(1, idx) Or cIdx
            End If
            QR_bit = True

        Else
            QR_bit = False
        End If

    Else
        QR_bit = True
        qrArr(1, idx) = qrArr(1, idx) And (255 - cIdx)
        If pBit > 0 Then qrArr(1, idx) = qrArr(1, idx) Or cIdx
        If pSiz < 0 Then qrArr(0, idx) = qrArr(0, idx) Or cIdx
    End If
End Function

Private Sub QR_bch_calc(ByRef data As Long, ByVal poly As Long)
    Dim b%, n%, rv&, x&
    If data = 0 Then Exit Sub

    b = QR_numbits(poly) - 1
    x = data * 2 ^ b

    rv = x
    Do
        n = QR_numbits(rv)
        If n <= b Then Exit Do
        rv = rv Xor (poly * 2 ^ (n - b - 1))
    Loop

    data = x + rv
End Sub

Private Function QR_numbits(ByVal Num As Long) As Integer
    Dim n%, a&

    a = 1
    n = 0
    Do While a <= Num
        a = a * 2
        n = n + 1
    Loop

    QR_numbits = n
End Function

Private Sub BB_putBits(ByRef pArr() As Byte, ByRef pPos As Integer, ByRef pVal As Variant, ByVal pLen As Integer)
    Dim idx As Integer, sIdx As Integer, bits As Integer
    Dim word As Long
    Dim dWd As Double
    Dim restLen As Integer
    Dim tmpArr(7) As Byte
    Dim sArr As Variant

    Select Case VarType(pVal)
    Case vbByte, vbInteger, vbLong, vbDouble
        If pLen > 56 Then
            outErr "BB_putbits: " & CStr(pLen) & " Too long"
            Exit Sub
        End If
        dWd = pVal
        If pLen < 56 Then dWd = dWd * 2 ^ (56 - pLen)
        idx = 0
        Do While idx < 6 And dWd > 0#
            word = Int(dWd / 2 ^ 48)
            tmpArr(idx) = word Mod 256
            dWd = dWd - 2 ^ 48 * word
            dWd = dWd * 256
            idx = idx + 1
        Loop
        sArr = tmpArr

    Case Else
        If InStr("Byte(),Integer(),Long(),Variant()", TypeName(pVal)) > 0 Then
            sArr = pVal

        Else
            outErr "BB_putbits: " & TypeName(pVal) & " Unknown type"
            Exit Sub
        End If
    End Select

    idx = Int(pPos / 8) + 1
    bits = pPos Mod 8
    sIdx = LBound(sArr)
    restLen = pLen
    Do While restLen > 0
        If sIdx <= UBound(sArr) Then
            word = sArr(sIdx)
            sIdx = sIdx + 1

        Else
            word = 0
        End If

        If restLen < 8 Then
            word = word And (&H100 - 2 ^ (8 - restLen))
        End If

        If bits > 0 Then
            word = word * 2 ^ (8 - bits)
            pArr(idx) = pArr(idx) Or Int(word / &H100)
            pArr(idx + 1) = pArr(idx + 1) Or (word And &HFF)

        Else
            pArr(idx) = pArr(idx) Or (word And &HFF)
        End If

        If restLen < 8 Then
            pPos = pPos + restLen
            restLen = 0

        Else
            pPos = pPos + 8
            idx = idx + 1
            restLen = restLen - 8
        End If
    Loop
End Sub

Private Function QR_makeAscMtrx(ByRef qrParam As tParams, ByRef qrArr() As Byte) As String
    Dim ascMtrx As String
    Dim rIdx As Integer, cIdx As Integer, bIdx As Integer
    Dim ch As Integer, idx As Long

    ascMtrx = ""

    For rIdx = 0 To qrParam.Siz Step 2
        bIdx = 0
        For cIdx = 0 To qrParam.Siz Step 2
            If (cIdx Mod 8) = 0 Then
                ch = qrArr(1, bIdx + 24 * rIdx)
                If rIdx < qrParam.Siz Then
                    idx = qrArr(1, bIdx + 24 * (rIdx + 1))

                Else
                    idx = 0
                End If
                bIdx = bIdx + 1
            End If
            ascMtrx = ascMtrx & Chr(97 + (ch Mod 4) + 4 * (idx Mod 4))
            ch = Int(ch / 4)
            idx = Int(idx / 4)
        Next cIdx
        ascMtrx = ascMtrx & vbLf
    Next rIdx

    QR_makeAscMtrx = ascMtrx
End Function

Private Function BC_to2Dim(ByVal pBarCode As String, ByRef ar() As Variant) As Boolean
    Dim rIdx As Integer, cIdx As Integer, t As Integer
    Dim txt As String, lenTxt As Integer
    Dim Pos As Integer
    Dim ch As String

    BC_to2Dim = False
    rIdx = 0
    cIdx = 0
    t = 0
    txt = Trim(pBarCode)
    lenTxt = Len(txt)
    For Pos = 1 To lenTxt
        ch = Mid(txt, Pos, 1)
        If ch >= "a" And ch <= "p" Then
            t = t + 2

        ElseIf ch = vbLf Or Pos = lenTxt Then
            If cIdx < t Then cIdx = t
            t = 0
            rIdx = rIdx + 2
        End If
    Next Pos
    If cIdx <= 0 Then
        outErr "BC_to2Dim : no data"
        Exit Function
    End If

    ReDim ar(1 To rIdx, 1 To cIdx)
    rIdx = 0
    cIdx = 0
    For Pos = 1 To lenTxt
        ch = Mid(txt, Pos, 1)
        If ch = vbLf Then
            cIdx = 0
            rIdx = rIdx + 2

        ElseIf ch >= "a" And ch <= "p" Then
            t = Asc(ch) - 97
            If (t And 1) = 1 Then ar(rIdx + 1, cIdx + 1) = 1
            If (t And 2) = 2 Then ar(rIdx + 1, cIdx + 2) = 1
            If (t And 4) = 4 Then ar(rIdx + 2, cIdx + 1) = 1
            If (t And 8) = 8 Then ar(rIdx + 2, cIdx + 2) = 1
            cIdx = cIdx + 2
        End If
    Next Pos

    BC_to2Dim = True
End Function
