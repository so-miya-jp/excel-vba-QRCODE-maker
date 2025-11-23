Attribute VB_Name = "MainModule"
Option Explicit

Public Sub WriteQRCode(ByRef pRng As Range, ByRef pInfo As String, ByVal pTarget As String, Optional ByVal pECL As eErrorCorrectionLevel = ECL_L)
    Dim Subject As String
    Dim ar() As Variant

    If Len(pTarget) <= 100 Then
        Subject = pTarget

    Else
        Subject = Left(pTarget, 48) & " c " & Right(pTarget, 48)
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

Public Function SplitFile(ByRef Result() As String, ByVal Path As String, Optional ByVal charSet As String = "UTF-8", Optional ByVal pECL As eErrorCorrectionLevel = ECL_L) As Boolean
    Dim Lines() As String
    Dim curText As String, oldText As String
    Dim idx As Long
    Dim v As Integer

    SplitFile = False

    If Not ReadTextFile(curText, Path, charSet) Then
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

Private Function ReadTextFile(ByRef Result As String, ByVal Path As String, Optional ByVal charSet As String = "UTF-8") As Boolean
    Dim buf As String

    ReadTextFile = False

    On Error GoTo ErrProc
    With CreateObject("ADODB.Stream")
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
