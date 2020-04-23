VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim Flag As Boolean
Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
    Dim i_TCD, Nb_VIN_Selection, Derniere_Lat, Derniere_Long, Derniere_Adr, Critere(G_Nb_VIN_Max)
    
    ' On Error GoTo fin
    If Not DicoRempli Then
        Call RemplisDicoJSON
    End If
    If Not Flag Then
        Flag = True
        Nb_VIN_Selection = 0
        If Not Macro_en_cours Then
            Sheets("accueil").Activate
            For Each i_TCD In ActiveSheet.PivotTables("TCD_VIN").PivotFields("VIN(s)").PivotItems
                If Not (i_TCD.Name = "(blank)") Then
                    If i_TCD.Visible Then
                      Sheets("Trajets").Activate
                      Nb_VIN_Selection = Nb_VIN_Selection + 1
                      Critere(Nb_VIN_Selection) = i_TCD.Name
                      Select Case Nb_VIN_Selection
                        Case 1
                            ActiveSheet.Range("$A$2:$B$65000").AutoFilter Field:=1, Criteria1:=Critere(1)
                        Case 2
                            ActiveSheet.Range("$A$2:$B$65000").AutoFilter Field:=1, Criteria1:=Critere(1), _
                               Operator:=xlOr, Criteria2:=Critere(2)
                        Case Else
                            ActiveSheet.Range("$A$2:$B$65000").AutoFilter Field:=1, Criteria1:=Critere, _
                               Operator:=xlFilterValues
                      End Select
                    End If
                End If
            Next
        End If
    ' V1.9.3 : Détermination de la dernière adresse atteinte
        Derniere_Lat = ""
        Derniere_Long = ""
        Derniere_Adr = ""
        If Nb_VIN_Selection = 1 Then
          ' uniquement si un seul vin sélectionné
            ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Select
            Derniere_Lat = Sheets("Trajets").Cells(Selection.Row, DicoJSON("endPosLatitude")).Value
            Derniere_Long = Sheets("Trajets").Cells(Selection.Row, DicoJSON("endPosLongitude")).Value
            Derniere_Adr = Sheets("Trajets").Cells(Selection.Row, DicoJSON("endPosAddress")).Value
        End If
        Sheets("Accueil").Activate
        Range(CELL_lat) = Derniere_Lat
        Range(CELL_long) = Derniere_Long
        Range(CELL_adr) = Derniere_Adr
    End If
fin:
    Flag = False
End Sub



