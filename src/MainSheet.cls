VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MainSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Const COL_TEXT% = 2
Const COL_RESULT% = 3
Const COL_ECL% = 4
Const COL_SHEET% = 5
Const COL_FLAG% = 6

'''[変換]ボタン押下時の処理
Sub OutputQRCode_Click()
    Dim Row As Long
    Dim ListObj As ListObject
    Dim tSh As Worksheet, pSh As Worksheet
    Dim targetTxt As String, eclTxt As String, shName As String, exeFlg As String
    Dim info As String
    Dim ecl As eErrorCorrectionLevel

    Set ListObj = MainSheet.ListObjects("変換対象一覧")
    Set pSh = MainSheet

    For Row = 1 To ListObj.ListRows.Count: With ListObj.ListRows(Row)
        targetTxt = .Range(COL_TEXT).Value
        eclTxt = .Range(COL_ECL).Value
        shName = .Range(COL_SHEET).Value
        exeFlg = .Range(COL_FLAG).Value

        If targetTxt = "" Then GoTo Continue_For
        If exeFlg <> "" Then GoTo Continue_For
            
        Set tSh = Nothing
        If shName <> "" Then
            'シート名が記入済みの場合、その名前のシートを対象シートとして取得
            On Error Resume Next
            Set tSh = Application.ThisWorkbook.Sheets(shName)
            On Error GoTo 0
        End If
        If tSh Is Nothing Then
            '対象シートがない場合、生成する。
            Set tSh = Application.ThisWorkbook.Sheets.Add(After:=pSh)
            MainSheet.Activate
            If shName <> "" Then
                'シート名が記入済みの場合、シート名の変更を試みる
                On Error Resume Next
                tSh.Name = shName
                On Error GoTo 0
            End If
            If tSh.Name <> shName Then
                'シート名列と対象シート名が異なる場合、シート名列を上書き
                shName = tSh.Name
                .Range(COL_SHEET).Value = shName
            End If
        End If
        Set pSh = tSh
        tSh.Cells.Clear
            
        Select Case UCase(eclTxt)
        Case "L": ecl = ECL_L
        Case "M": ecl = ECL_M
        Case "H": ecl = ECL_H
        Case "Q": ecl = ECL_Q
        Case Else
            ecl = ECL_L
            .Range(COL_ECL).Value = "L"
        End Select

        WriteQRCode tSh.Range("A1"), info, targetTxt, ecl
        .Range(COL_RESULT).Value = info
        If info Like "ver.*" Then
            .Range(COL_FLAG).Value = "済"
        End If
    
Continue_For:
    End With: Next Row
End Sub

'''[選択]ボタン押下時の処理
Sub SelectFile_Click()
    Dim Rng As Range
    Dim Path As String

    Set Rng = MainSheet.Cells(3, 2)
    Path = Rng.Value
    If Path = "" Then
        Path = Application.ThisWorkbook.Path
    End If
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "分割対象のテキストファイルを選択してください。"
        .AllowMultiSelect = False
        .InitialFileName = Path
        If .Show Then
            Rng.Value = .SelectedItems(1)
        End If
    End With
End Sub

'''[分割]ボタン押下時の処理
Sub SplitFile_Click()
    Dim SrcRng As Range, ListRng As Range
    Dim Path As String, charSet As String, eclTxt As String, prefix As String
    Dim msg As String
    Dim ecl As eErrorCorrectionLevel
    Dim Result() As String
    Dim msgRet As VbMsgBoxResult
    Dim idx As Long, Row As Long
    Dim fmt As String

    Set SrcRng = MainSheet.ListObjects("分割対象").DataBodyRange
    Set ListRng = MainSheet.ListObjects("変換対象一覧").DataBodyRange

    With SrcRng
        Path = .Cells(1, 1).Value
        charSet = .Cells(1, 2).Value
        eclTxt = .Cells(1, 3).Value
        prefix = .Cells(1, 4).Value
    End With

    msg = ""
    If Path = "" Then
        msg = msg & vbLf & "分割対象のファイルパスを指定してください。"

    ElseIf Dir(Path) = "" Then
        msg = msg & vbLf & "分割対象のファイルパスが存在しません。"
    End If

    If charSet = "" Then
        charSet = "Shift_JIS"
        SrcRng.Cells(1, 2).Value = charSet
    End If

    Select Case UCase(eclTxt)
    Case "L": ecl = ECL_L
    Case "M": ecl = ECL_M
    Case "H": ecl = ECL_H
    Case "Q": ecl = ECL_Q
    Case Else
        ecl = ECL_L
        eclTxt = "L"
        SrcRng.Cells(1, 3).Value = "L"
    End Select

    If msg <> "" Then
        MsgBox msg
        Exit Sub
    End If

    If Not SplitFile(Result, Path, charSet, ecl) Then
        MsgBox "ファイル分割に失敗しました。"
        Exit Sub
    End If

    '変換対象一覧の初期化
    Clear_Click

    If prefix <> "" Then
        'Format用の書式作成。分割数の桁数分の連続した"0"を作成。
        fmt = Replace(Space(Len(CStr(UBound(Result)))), " ", "0")
    End If

    Row = 1
    For idx = LBound(Result) To UBound(Result)
        Do While ListRng.Cells(Row, COL_TEXT) <> "": Row = Row + 1: Loop
        ListRng.Cells(Row, COL_TEXT).Value = Result(idx)
        ListRng.Cells(Row, COL_ECL).Value = eclTxt
        If prefix <> "" Then
            ListRng.Cells(Row, COL_SHEET).Value = prefix & Format(idx, fmt)
        End If
        ListRng.Cells(Row, COL_FLAG).Value = ""
        Row = Row + 1
    Next idx
End Sub

'''[クリア]ボタン押下時の処理
Sub Clear_Click()
    Dim ListObj As ListObject
    Dim Rng As Range

    Set ListObj = MainSheet.ListObjects("変換対象一覧")

    With ListObj
        Set Rng = Union( _
            .ListColumns(COL_TEXT).DataBodyRange, _
            .ListColumns(COL_RESULT).DataBodyRange, _
            .ListColumns(COL_ECL).DataBodyRange, _
            .ListColumns(COL_SHEET).DataBodyRange, _
            .ListColumns(COL_FLAG).DataBodyRange)
    End With
    If Not IsBlankRange(Rng) Then
        Select Case MsgBox("変換対象一覧を初期化しますか？", vbYesNo)
        Case vbYes
            Rng.ClearContents
        End Select
    End If
End Sub

