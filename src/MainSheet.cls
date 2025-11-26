VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MainSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Enum LIST_COL
    COL_NO = 1
    COL_TEXT = 2
    COL_RESULT = 3
    COL_ECL = 4
    COL_SHEET = 5
    COL_FLAG = 6
End Enum
Const NUMBER_FORMULA As String = "=IF([@QRコード変換対象文字列]="""", """", ROW() - ROW(変換対象一覧[#Headers]))"

'''[変換]ボタン押下時の処理
Sub OutputQRCode_Click()
    Dim Row As Long, outRow As Long
    Dim ListObj As ListObject
    Dim tSh As Worksheet, pSh As Worksheet
    Dim targetTxt As String, eclTxt As String, shName As String, exeFlg As String
    Dim shNameChunk As String
    Dim info As String
    Dim ecl As eErrorCorrectionLevel

    Set ListObj = MainSheet.ListObjects("変換対象一覧")
    Set pSh = MainSheet

    shNameChunk = "#"
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
            Set pSh = tSh
            ActiveWindow.Zoom = 70
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

        If InStr(shNameChunk, "#" & tSh.Name & "#") > 0 Then
            outRow = GetLastUsedRow(tSh) + 2

        Else
            shNameChunk = shNameChunk & tSh.Name & "#"
            tSh.Cells.Clear
            outRow = 1
        End If

        Select Case UCase(eclTxt)
        Case "L": ecl = ECL_L
        Case "M": ecl = ECL_M
        Case "H": ecl = ECL_H
        Case "Q": ecl = ECL_Q
        Case Else
            ecl = ECL_L
            .Range(COL_ECL).Value = "L"
        End Select

        WriteQRCode tSh.Cells(outRow, 1), info, targetTxt, ecl
        .Range(COL_RESULT).Value = info
        If LCase(info) Like "ver.*" Then
            .Range(COL_FLAG).Value = "済"
        End If

Continue_For:
    End With: Next Row
End Sub

'''[選択]ボタン押下時の処理
Sub SelectFile_Click()
    Dim SrcObj As ListObject
    Dim selRng As Range, r As Range
    Dim idx As Long, tIdx As Long
    Dim Path As String

    Set SrcObj = MainSheet.ListObjects("分割対象")
    tIdx = 0
    If TypeOf Selection Is Range Then
        Set selRng = Intersect(Selection, SrcObj.DataBodyRange)
        If Not selRng Is Nothing Then
            '分割対象が選択されていた場合
            '分割対象の選択範囲のうち、最後のファイルパスをデフォルトパスとし、上書き対象とする。
            For idx = SrcObj.ListRows.Count To 1 Step -1
                If Not Intersect(selRng, SrcObj.ListRows(idx).Range) Is Nothing Then
                    Path = SrcObj.ListRows(idx).Range(1).Value
                    If Path <> "" Then
                        tIdx = idx
                        Exit For
                    End If
                End If
            Next idx
        End If
    End If

    If Path = "" Then
        '選択なし、または選択範囲が空文字の場合
        If SrcObj.ListRows.Count = 1 Then
            '分割対象が1行の場合、上書き対象とする。
            tIdx = 1
            Path = SrcObj.ListRows(1).Range(1).Value

        Else
            '分割対象が複数行の場合
            '分割対象の最終行をデフォルトパスとする。
            '設定行は最初の空行とする。
            For idx = 1 To SrcObj.ListRows.Count
                If SrcObj.ListRows(idx).Range(1).Value = "" Then
                    If tIdx = 0 Then tIdx = idx

                Else
                    Path = SrcObj.ListRows(idx).Range(1).Value
                End If
            Next idx
        End If
    End If

    If Path = "" Then
        Path = Application.ThisWorkbook.Path
    End If

    If tIdx = 0 Then
        MsgBox "選択したファイルを設定する余裕がありません。空行を用意するか、設定行を選択してください。"
        Exit Sub
    End If

    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "分割対象のテキストファイルを選択してください。"
        .AllowMultiSelect = False
        .InitialFileName = Path

        If .Show Then
            SrcObj.ListRows(tIdx).Range(1).Value = .SelectedItems(1)
        End If
    End With
End Sub

'''[分割]ボタン押下時の処理
Sub SplitFile_Click()
    Dim SrcObj As ListObject, ListObj As ListObject
    Dim SrcAr As Variant
    Dim Path As String, charSet As String, eclTxt As String, prefix As String
    Dim msg As String
    Dim ecl As eErrorCorrectionLevel
    Dim Result() As String
    Dim msgRet As VbMsgBoxResult
    Dim idx As Long, srcRow As Long, lstRow As Long
    Dim fmt As String

    Set SrcObj = MainSheet.ListObjects("分割対象")
    Set ListObj = MainSheet.ListObjects("変換対象一覧")
    SrcAr = SrcObj.DataBodyRange

    '入力チェック及び未設定項目への初期値設定
    msg = ""
    idx = 0
    For srcRow = LBound(SrcAr, 1) To UBound(SrcAr, 1)
        Path = SrcAr(srcRow, 1)
        charSet = SrcAr(srcRow, 2)
        eclTxt = SrcAr(srcRow, 3)

        If Path <> "" Then
            If Dir(Path) = "" Then
                msg = msg & vbLf & CStr(srcRow) & "行目の分割対象のファイルが存在しません。"

            Else
                idx = idx + 1
            End If

            If charSet = "" Then
                SrcObj.ListRows(srcRow).Range(2).Value = "Shift_JIS"
            End If

            Select Case eclTxt
            Case "L", "M", "H", "Q"

            Case "l", "m", "h", "q"
                SrcObj.ListRows(srcRow).Range(3).Value = UCase(eclTxt)

            Case Else
                SrcObj.ListRows(srcRow).Range(3).Value = "L"
            End Select
        End If
    Next srcRow

    If idx = 0 Then
        msg = msg & vbLf & "分割対象のファイルパスを指定してください。"
    End If

    If msg <> "" Then
        MsgBox Mid(msg, 2)
        Exit Sub
    End If

    SrcAr = SrcObj.DataBodyRange

    '変換対象一覧の初期化
    Clear_Click

    lstRow = 0
    For srcRow = LBound(SrcAr, 1) To UBound(SrcAr, 1)
        Path = SrcAr(srcRow, 1)
        charSet = SrcAr(srcRow, 2)
        eclTxt = SrcAr(srcRow, 3)
        prefix = SrcAr(srcRow, 4)

        If Path = "" Then GoTo Continue_For

        Select Case eclTxt
        Case "L": ecl = ECL_L
        Case "M": ecl = ECL_M
        Case "H": ecl = ECL_H
        Case "Q": ecl = ECL_Q
        End Select

        Erase Result
        If Not SplitFile(Result, Path, charSet, ecl) Then
            MsgBox "ファイル分割に失敗しました。" & vbLf & Path
            GoTo Continue_For
        End If

        If prefix <> "" Then
            'Format用の書式作成。分割数の桁数分の"0"を作成。
            fmt = Replace(Space(Len(CStr(UBound(Result)))), " ", "0")
        End If

        For idx = LBound(Result) To UBound(Result)
            Do
                lstRow = lstRow + 1
                If lstRow > ListObj.ListRows.Count Then
                    ListObj.ListRows.Add
                    With ListObj.ListRows(lstRow)
                        .Range.RowHeight = 30#
                        If .Range(COL_NO).Formula = "" Then
                            .Range(COL_NO).Formula = NUMBER_FORMULA
                        End If
                    End With
                    Exit Do
                End If
            Loop While ListObj.ListRows(lstRow).Range(COL_TEXT) <> ""

            With ListObj.ListRows(lstRow)
                .Range(COL_TEXT).Value = Result(idx)
                .Range(COL_ECL).Value = eclTxt
                If prefix <> "" Then
                    .Range(COL_SHEET).Value = prefix & Format(idx, fmt)
                End If
                .Range(COL_FLAG).Value = ""
            End With
        Next idx
Continue_For:
    Next srcRow
End Sub

'''[クリア]ボタン押下時の処理
Sub Clear_Click()
    Dim ListObj As ListObject
    Dim TargetRng As Range, Rng As Range
    Dim sh As Worksheet

    Set ListObj = MainSheet.ListObjects("変換対象一覧")

    With ListObj
        Set TargetRng = Union( _
            .ListColumns(COL_TEXT).DataBodyRange, _
            .ListColumns(COL_RESULT).DataBodyRange, _
            .ListColumns(COL_ECL).DataBodyRange, _
            .ListColumns(COL_SHEET).DataBodyRange, _
            .ListColumns(COL_FLAG).DataBodyRange)
    End With
    If Not IsBlankRange(TargetRng) Then
        Select Case MsgBox("変換対象一覧を初期化します。シート名に対応するシートも削除しますか？", vbYesNoCancel)
        Case vbYes
            '対応するシートが存在すれば削除する。
            Application.DisplayAlerts = False
            For Each Rng In ListObj.ListColumns(COL_SHEET).DataBodyRange
                On Error Resume Next
                Set sh = Nothing
                Set sh = Application.ThisWorkbook.Sheets(Rng.Value)
                On Error GoTo 0
                If Not sh Is Nothing Then sh.Delete
            Next Rng
            Application.DisplayAlerts = True

        Case vbNo
            'nop

        Case Else
            Exit Sub
        End Select

        TargetRng.ClearContents
    End If
End Sub

