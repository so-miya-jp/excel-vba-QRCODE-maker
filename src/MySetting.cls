VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MySetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---- メンバー変数
Private MaskPattern_ As eMaskType
Private ModeSet_ As eModeBit
Private ImageFolder_ As String
Private ErrorMessage_ As String
'----

'''クラス初期処理
Private Sub Class_Initialize()
    LoadSetting
End Sub

Public Sub LoadSetting()
    Dim sh As Worksheet
    Dim ar() As Variant
    Dim idx As Long
    Dim sKey As String, sValue As String, lValue As Long

    Set sh = SettingSheet
    ar = sh.ListObjects("設定表").DataBodyRange

    ErrorMessage_ = ""

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
                OutputError sKey, idx, "想定外の値"
                MaskPattern_ = MSK_AUTO
            End If

        Case "MODE_SET" 'モードセット
            If lValue >= 0 And lValue <= 15 And (lValue And 4) = 4 Then
                ModeSet_ = lValue

            Else
                OutputError sKey, idx, "想定外の値"
                ModeSet_ = MOD_ALL
            End If

        Case "IMAGE_FOLDER" '画像フォルダ
            If sValue = "" Then
                ImageFolder_ = ""

            ElseIf Dir(sValue, vbDirectory) <> "" Then
                ImageFolder_ = sValue

            Else
                OutputError sKey, idx, "ディレクトリが存在しません。"
                ImageFolder_ = ""
            End If
        End Select
    Next idx

End Sub

'''エラー出力
''' @param ID      / I / ID
''' @param Line    / I / 行番号
''' @param Message / I / メッセージ
Private Sub OutputError(ByVal ID As String, ByVal Line As Long, ByVal Message As String)
    If ErrorMessage_ <> "" Then ErrorMessage_ = ErrorMessage_ & vbLf
    ErrorMessage_ = ErrorMessage_ & ID & "：" & Message
End Sub

'''エラー有無プロパティ
''' @return エラー有無
Public Property Get IsError() As Boolean
    IsError = ErrorMessage_ <> ""
End Property

'''エラーメッセージプロパティ
''' @return エラーメッセージ
Public Property Get ErrorMessage() As String
    ErrorMessage = ErrorMessage_
End Property

'''マスクパターンプロパティ
''' @return マスクパターン
Public Property Get MaskPattern() As eMaskType
    MaskPattern = MaskPattern_
End Property

'''モードセットプロパティ
''' @return モードセット
Public Property Get ModeSet() As eModeBit
    ModeSet = ModeSet_
End Property

'''画像フォルダプロパティ
''' @return 画像フォルダ
Public Property Get ImageFolder() As String
    ImageFolder = ImageFolder_
End Property
