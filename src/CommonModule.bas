Attribute VB_Name = "CommonModule"
Option Explicit
Option Private Module

'''文字列を配列に格納する
'''引数
'''  Texts     / IO / 格納する配列。
'''  Text      / I  / 格納される文字列。
Public Sub AddArrayText(ByRef Texts() As String, ByVal Text As String)
    Dim Size As Integer

    If (Not Texts) = -1 Then
        Size = 0

    Else
        Size = UBound(Texts) + 1
    End If

    ReDim Preserve Texts(Size)
    Texts(Size) = Text
End Sub


'''指定シートの最終行を取得する。
'''引数
'''  sh      / I / 指定シート
'''  needVal / I / Trueの場合、ブランク行をスキップする。(省略時はTrue)
'''戻り値
'''  有効値のある最終行
Public Function GetLastUsedRow(ByRef sh As Worksheet, Optional needVal As Boolean = True) As Long
    With sh.UsedRange
        '使用Rangeの最終行を取得
        GetLastUsedRow = .Rows(.Rows.Count).Row
    End With

    If Not needVal Then Exit Function

    With sh
        'UsedRangeは書式だけ変更された行もカウントするので
        'ブランク行を後からスキップする。
        Do While IsBlankRange(sh.Rows(GetLastUsedRow))
            GetLastUsedRow = GetLastUsedRow - 1
        Loop
    End With
End Function

'''指定Rangeが全てブランクかどうかを判定する。
'''引数
'''  Rng / I / 指定Range
'''戻り値
'''  指定Rangeが全てブランクならばTrue。そうでない場合はFalseを返却。
Public Function IsBlankRange(ByRef Rng As Range) As Boolean
    IsBlankRange = (WorksheetFunction.CountBlank(Rng) = Rng.Count)
End Function
