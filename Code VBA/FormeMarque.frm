VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormeMarque 
   Caption         =   "Choix de la marque à exporter"
   ClientHeight    =   2235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6165
   OleObjectBlob   =   "FormeMarque.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormeMarque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AnnulerBouton_Click()
    FormeMarque.BoutonChoisi.Value = 2
    FormeMarque.Hide
End Sub
Private Sub OkBouton_Click()
    FormeMarque.BoutonChoisi.Value = 1
    FormeMarque.Hide
End Sub
Private Sub userform_Activate()
    FormeMarque.ListeMarque.SetFocus
    FormeMarque.BoutonChoisi.Value = 10
End Sub
Private Sub UserForm_Initialize()
    FormeMarque.ListeMarque.SetFocus
    FormeMarque.BoutonChoisi.Value = 10
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True
End Sub
