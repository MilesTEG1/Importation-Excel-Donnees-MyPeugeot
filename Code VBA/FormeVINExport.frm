VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormeVINExport 
   Caption         =   "Choix du (des) VIN"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6735
   OleObjectBlob   =   "FormeVINExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormeVINExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AnnulerBouton_Click()
    FormeVINExport.BoutonChoisi.Value = 2
    FormeVINExport.Hide
End Sub
Private Sub OkBouton_Click()
    FormeVINExport.BoutonChoisi.Value = 1
    FormeVINExport.Hide
End Sub
Private Sub userform_Activate()
    FormeVINExport.ListeVIN.SetFocus
    FormeVINExport.BoutonChoisi.Value = 10
End Sub
Private Sub UserForm_Initialize()
    FormeVINExport.ListeVIN.SetFocus
    FormeVINExport.BoutonChoisi.Value = 10
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True
End Sub
