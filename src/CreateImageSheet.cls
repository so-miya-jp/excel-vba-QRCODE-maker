VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CreateImageSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'''[変換]ボタン押下時の処理
Sub CreateQRCodeImage_Click()
    Dim sh As Worksheet
    Dim ar As Variant
    Dim txtTarget As String
    Dim ecl As eErrorCorrectionLevel
    Dim txtFilePath As String
    Dim elementRng As Range
    Dim bgColor As Long
    Dim Result As String

    Set sh = CreateImageSheet
    ar = sh.Range("C2:C6").Value

    '変換対象文字列
    txtTarget = ar(1, 1)

    'エラー補正レベル
    Select Case UCase(ar(2, 1))
    Case "L": ecl = ECL_L
    Case "M": ecl = ECL_M
    Case "H": ecl = ECL_H
    Case "Q": ecl = ECL_Q
    Case Else
        ecl = ECL_L
        sh.Range("C3").Value = "L"
    End Select

    '出力ファイルパス
    txtFilePath = ar(3, 1)
    
    '要素
    Set elementRng = Nothing
    If Trim(ar(4, 1)) <> "" Then
        On Error Resume Next
        Set elementRng = Range(ar(4, 1))
        On Error GoTo 0
    End If

    '背景色
    bgColor = vbWhite
    If Trim(ar(5, 1)) <> "" Then
        On Error Resume Next
        bgColor = WorksheetFunction.Hex2Dec(ar(5, 1))
        On Error GoTo 0
    End If

    'QRコード画像の生成
    DrawQRCodeImage sh.Range("C10"), Result, txtFilePath, elementRng, bgColor, txtTarget, ecl
    sh.Range("C9").Value = Result
End Sub

'''出力ファイル[選択]ボタン押下時の処理
Sub SelectFile_Click()
    Dim sh As Worksheet
    Dim initPath As String
    Dim Result As Variant

    Set sh = CreateImageSheet

    initPath = sh.Range("C4").Value
    If initPath = "" Then
        initPath = Application.ThisWorkbook.Path
    End If

    Result = Application.GetSaveAsFilename(initPath, "bitmap files, *.bmp", 1, "出力するBMPファイル")
    If Result <> False Then
        sh.Range("C4").Value = Result
    End If
End Sub

'''要素位置[選択]ボタン押下時の処理
Sub SelectCell_Click()
    Dim sh As Worksheet
    Dim initText As String
    Dim Rng As Range

    Set sh = CreateImageSheet

    initText = sh.Range("C5").Value
    On Error Resume Next
    Set Rng = Application.InputBox("マウスでセル範囲指定をしてください", Type:=8, Default:=initText)

    If Not Rng Is Nothing Then
        If Not Rng.Worksheet Is sh Then
            initText = Rng.Worksheet.Name & "!" & Rng.Address
        Else
            initText = Rng.Address
        End If
        sh.Range("C5").Value = initText
    End If
End Sub

'''背景色[選択]ボタン押下時の処理
Sub SelectColor_Click()
    Dim sh As Worksheet
    Dim txtBGColor As String
    Dim iR As Byte, iG As Byte, iB As Byte

    Set sh = CreateImageSheet
    txtBGColor = sh.Range("C6").Value
    If txtBGColor = "" Then
        SplitRGB vbWhite, iR, iG, iB

    Else
        SplitRGB WorksheetFunction.Hex2Dec(txtBGColor), iR, iG, iB
    End If

    With Application.Dialogs(xlDialogEditColor)
        If .Show(1, Int(iR), Int(iG), Int(iB)) Then
            sh.Range("C6").Value = Right("00000" & Hex(ActiveWorkbook.colors(1)), 6)
        End If
    End With
End Sub
