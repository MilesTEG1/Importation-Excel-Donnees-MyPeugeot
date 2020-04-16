VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormeVIN 
   Caption         =   "Choix du (des) VIN"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6735
   OleObjectBlob   =   "FormeVIN.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormeVIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FormeVIN_AnnulerBouton_Click()
    FormeVIN.BoutonChoisi.Value = 2
    FormeVIN.Hide
End Sub
Private Sub FormeVIN_OkBouton_Click()
    FormeVIN.BoutonChoisi.Value = 1
    FormeVIN.Hide
End Sub
Private Sub userform_Activate()
    FormeVIN.FormeVIN_ListeVIN.SetFocus
    FormeVIN.BoutonChoisi.Value = 10
End Sub
Private Sub UserForm_Initialize()
    FormeVIN.FormeVIN_ListeVIN.SetFocus
    FormeVIN.BoutonChoisi.Value = 10
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True
End Sub
