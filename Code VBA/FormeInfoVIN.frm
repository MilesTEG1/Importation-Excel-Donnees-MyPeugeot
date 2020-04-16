VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormeInfoVIN 
   Caption         =   "Saisie des informations d'un VIN"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10170
   OleObjectBlob   =   "FormeInfoVIN.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormeInfoVIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OkBouton_Click()
    If Me.DescriptionVIN = "" Then
        MsgBox "Il faut saisir la description de ce VIN", vbCritical
    Else
        Me.Hide
    End If
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True
End Sub
Private Sub UserForm_Initialize()
    Me.DescriptionVIN.SetFocus
End Sub

