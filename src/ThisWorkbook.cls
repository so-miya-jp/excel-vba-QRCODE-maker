VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Rem --------------------------------------------------------------------------------
Rem QRCODE-maker
Rem ThisWorkbook
Rem --------------------------------------------------------------------------------

Option Explicit

'''エクセルブックオープン時の処理
Private Sub Workbook_Open()
    MySetting.LoadSetting True
End Sub
