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
    Dim setting As New MySetting

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
        Set elementRng = Application.Range(CStr(ar(4, 1)))
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
    DrawQRCodeImage sh.Range("C10"), Result, txtFilePath, elementRng, bgColor, txtTarget, ecl, setting.MaskPattern, setting.ModeSet
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
            initText = Rng.Address(External:=True)

        Else
            initText = Rng.Address
        End If
        If Left(initText, 1) = "'" Then initText = "'" & initText
        sh.Range("C5").Value = initText
    End If
End Sub

'''要素位置[画像読込]ボタン押下時の処理
Sub LoadImage_Click()
    Dim sh As Worksheet
    Dim initText As String
    Dim Path As String
    Dim data() As Long
    Dim Rng As Range
    Dim r As Long, c As Long
    Dim setting As New MySetting

    If setting.IsError Then
        MsgBox "設定にエラーがあります。" & vbLf & setting.ErrorMessage
    End If

    Set sh = CreateImageSheet

    initText = sh.Range("C5").Value

    Path = ""
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "画像ファイルを選択してください。"
        .Filters.Add "全ての画像形式", "*.bmp; *.dib; *.rle; *.png; *.gif; *.jpg; *.jpeg; *.jfif; *.jpe; *.webp; *.ico; *.tif; *.tiff; *.heif; *.heic; *.hif; *.avif; *.svg; *.emf; *.wmf; *.emz; *.wmz"
        .Filters.Add "Windowsビットマップ", "*.bmp; *.dib; *.rle"
        .Filters.Add "PNG形式", "*.png"
        .Filters.Add "GIF形式", "*.gif"
        .Filters.Add "JPEG形式", "*.jpg; *.jpeg; *.jfif; *.jpe"
        .Filters.Add "WebP", "*.webp"
        .Filters.Add "TIFF形式", "*.tif; *.tiff"
        .Filters.Add "HEIF形式", "*.heif; *.heic; *.hif"
        .Filters.Add "アイコン", "*.ico"
        .Filters.Add "AV1画像ファイル形式", "*.avif"
        .Filters.Add "スケーラブルベクターグラフィックス", "*.svg"
        .Filters.Add "Windows 拡張メタファイル", "*.emf"
        .Filters.Add "Windows メタファイル", "*.wmf"
        .Filters.Add "圧縮Windows拡張メタファイル", "*.emz"
        .Filters.Add "圧縮Windowsメタファイル", "*.wmz"
        .AllowMultiSelect = False
        If setting.ImageFolder = "" Then
            .InitialFileName = Application.ThisWorkbook.Path

        Else
            .InitialFileName = setting.ImageFolder
        End If

        If .Show Then
            Path = .SelectedItems(1)
        End If
    End With
    
    If Path = "" Then Exit Sub
    
    If Not ImportImage(Path, data, sh) Then
        MsgBox "画像の読み込みに失敗しました。"
        Exit Sub
    End If

    Set Rng = Range(initText)
    If Rng Is Nothing Then
        Set Rng = sh.Range("E9")

    ElseIf Rng.Areas.Count > 1 Then
        Set Rng = Rng.Areas(1)
    End If
    Set Rng = Rng.Resize(UBound(data, 1), UBound(data, 2))
    
    Do
        If Not Rng.Worksheet Is sh Then
            initText = Rng.Address(External:=True)

        Else
            initText = Rng.Address
        End If

        On Error Resume Next
        Set Rng = Nothing
        Set Rng = Application.InputBox("画像サイズが" & UBound(data, 1) & "×" & UBound(data, 2) & "あります。" & vbLf & "マウスで画像を背景色に反映するセル範囲を指定をしてください", Type:=8, Default:=initText)
        On Error GoTo 0

        If Rng Is Nothing Then Exit Do
        If Rng.Areas.Count > 1 Then
            Set Rng = Rng.Areas(1)

        ElseIf Rng.Rows.Count = UBound(data, 1) Then
            If Rng.Columns.Count = UBound(data, 2) Then
                Exit Do
            End If
        End If
        Set Rng = Rng.Resize(UBound(data, 1), UBound(data, 2))
    Loop
    
    If Rng Is Nothing Then
        MsgBox "画像の読み込みを中断しました。"
        Exit Sub
    End If

    With Rng
        .RowHeight = 10
        .ColumnWidth = 1
        
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous

        For r = 1 To .Rows.Count
            For c = 1 To .Columns.Count
                .Cells(r, c).Interior.Color = data(r, c)
            Next c
        Next r

        If Left(initText, 1) = "'" Then initText = "'" & initText
        sh.Range("C5").Value = initText
        sh.Range("A1").Select
        sh.Range("C5").Select
    End With
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
