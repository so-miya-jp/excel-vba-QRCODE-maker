VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MySetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'---- メンバー変数
Private DirtyFlag_ As Boolean
Private MaskPattern_ As eMaskType
Private ModeSet_ As eModeBit
Private ImageFolder_ As String
'----

'''クラス初期処理
Private Sub Class_Initialize()
    DirtyFlag_ = True
    LoadSetting True
End Sub

'''設定の読み込み
''' @param Silenced / I / Trueの場合、エラー時にメッセージを非表示とする(省略時はFalse)
Public Sub LoadSetting(Optional ByVal Silenced As Boolean = False)
    Dim sh As Worksheet
    Dim Rng As Range
    Dim ar() As Variant
    Dim msg As String
    Dim msgChunk As String
    Dim idx As Long
    Dim sKey As String, sValue As String, lValue As Long

    If Not DirtyFlag_ Then Exit Sub

    Set sh = SettingSheet
    Set Rng = sh.ListObjects("設定表").DataBodyRange

    If Not Silenced Then
        If sh.Comments.Count > 0 Then
            If Not Intersect(sh.Cells.SpecialCells(xlCellTypeComments), Rng) Is Nothing Then
                'コメントが存在する場合、コメントを削除
                Rng.ClearComments
#If DEBUG_ > 0 Then
                Debug.Print "LoadSetting ClearComments"
#End If
            End If
        End If

        If Rng.Interior.ColorIndex <> xlNone Then
            '背景色の変更をなしにする。
            Rng.Interior.ColorIndex = xlNone
#If DEBUG_ > 0 Then
            Debug.Print "LoadSetting ColorIndex = xlNone"
#End If
        End If
    End If

    ar = Rng

    msgChunk = ""

    For idx = LBound(ar, 1) To UBound(ar, 1)
        sKey = UCase(Trim(ar(idx, 1)))
        sValue = Trim(ar(idx, 2))
        If InStr(sValue, ":") > 0 Then
            sValue = Trim(Split(sValue, ":")(0))
        End If
        lValue = Val(sValue)

        Select Case sKey
        Case "MASK_PATTERN" 'マスクパターン
            If UCase(sValue) = "A" Then
                MaskPattern_ = MSK_AUTO

            ElseIf lValue >= MSK_000 And lValue <= MSK_111 Then
                MaskPattern_ = lValue

            Else
                msg = "想定外の値"
                GoSub ErrorSub

                MaskPattern_ = MSK_AUTO
            End If

        Case "MODE_SET" 'モードセット
            If lValue >= 0 And lValue <= 15 And (lValue And 4) = 4 Then
                ModeSet_ = lValue

            Else
                msg = "想定外の値"
                GoSub ErrorSub

                ModeSet_ = MOD_ALL
            End If

        Case "IMAGE_FOLDER" '画像フォルダ
            If sValue = "" Then
                ImageFolder_ = ""

            ElseIf Dir(sValue, vbDirectory) <> "" Then
                ImageFolder_ = sValue

            Else
                msg = "ディレクトリが存在しません"
                GoSub ErrorSub

                ImageFolder_ = ""
            End If
        End Select
    Next idx

    If Not Silenced And msgChunk <> "" Then
        MsgBox Mid(msgChunk, 2)
    End If

    DirtyFlag_ = False

    Exit Sub

ErrorSub:
    msgChunk = msgChunk & vbLf & sKey & ":" & msg & "(" & sValue & ")"
    If Not Silenced Then
        With Rng.Cells(idx, 2)
            .AddComment msg
            .Interior.Color = rgbPink
        End With
    End If

    Return
End Sub

'''設定変更時に呼び出される。
Public Sub Update()
    DirtyFlag_ = True
End Sub

'''マスクパターンプロパティ
''' @return マスクパターン
Public Property Get MaskPattern() As eMaskType
    LoadSetting
    MaskPattern = MaskPattern_
End Property

'''モードセットプロパティ
''' @return モードセット
Public Property Get ModeSet() As eModeBit
    LoadSetting
    ModeSet = ModeSet_
End Property

'''画像フォルダプロパティ
''' @return 画像フォルダ
Public Property Get ImageFolder() As String
    LoadSetting
    ImageFolder = ImageFolder_
End Property
