VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Dim MyErrorCheckValue As Boolean

Private Sub Workbook_Open()
    MyErrorCheckValue = Application.ErrorCheckingOptions.NumberAsText
    Application.ErrorCheckingOptions.NumberAsText = False
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.ErrorCheckingOptions.NumberAsText = MyErrorCheckValue
End Sub
Private Sub Workbook_Activate()
    Application.ErrorCheckingOptions.NumberAsText = False
End Sub
Private Sub Workbook_Deactivate()
    Application.ErrorCheckingOptions.NumberAsText = MyErrorCheckValue
End Sub
