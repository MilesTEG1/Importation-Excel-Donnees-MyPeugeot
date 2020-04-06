Attribute VB_Name = "Module1"
'
' Fonction pour lire les données depuis un fichier JSON et les écrire dans le tableau
'
Sub MYP_JSON_Decode()
    Dim jsonText As String
    Dim jsonObject As Object, item As Object, item_item As Object
    Dim i As Long
    Dim ws As Worksheet
    
    Dim FSO As New FileSystemObject
    Dim JsonTS As TextStream
    
    Dim FichierMYP As String    ' nom du fichier myp
    Dim nbTrip, nbTripReel As Integer ' nombre de trajets
    Dim kilometrage As Single   ' Pour stocker le kilométrage total de la voiture.
    Dim conso_totale As Single  ' Pour stocker la consommation totale de tout le tableau
    
    Set ws = Worksheets("Trajets-MyPeugeot")
    
'    jsonText = ws.Range("Q1")   ' Pour le 1er test, je mets le contenu du fichier test JSON dans la cellule Q1
    
    FichierMYP = Application.GetOpenFilename("Fichiers trajets Peugeot App (*.myp),*.myp")  ' On demande la sélection du fichier
    If FichierMYP = "" Then
        MsgBox "Aucun fichier n'a été selectionné !"
        Exit Sub
    Else
        Set JsonTS = FSO.OpenTextFile(FichierMYP, ForReading)
        jsonText = JsonTS.ReadAll
        JsonTS.Close
    
        ' Comme le fichier existe, on efface tout
        Effacage_Donnees
        
        nbTrip = 0    ' On réinitialise le nombre de trajets
    
        ws.Range("N1") = FichierMYP     ' On écrit le chemin d'accès du fichier.
    End If
    
  
    Set jsonObject = JsonConverter.ParseJson(jsonText)
  
    i = 5   ' Le n° de la ligne où on va commencer à écrire les données
            ' ws.Cells(LIGNE, COLONNE)     Où LIGNE commence à 1
            '                              Où COLONNE commence à 1 = A
            ' COLONNE = 1 (A) -> Trip ID#
            '         = 2 (B) -> Date de Début  (à déterminer avec une conversion)
            '         = 3 (C) -> Date de Fin    (à déterminer avec une conversion)
            '         = 4 (D) -> Durée du trajet  (à calculer)
            '         = 5 (E) -> Distance du trajet en km
            '         = 6 (F) -> Distance totale au compteur km
            '         = 7 (G) -> Consommation du tajet en L
            '         = 8 (H) -> Consommation moyenne en L/100km
            '         = 9 (I) -> Position de départ - Latitude
            '         = 10 (J) -> Position de départ - Longitude
            '         = 11 (K) -> Position de départ - Adresse
            '         = 12 (L) -> Position d'arrivée - Latitude
            '         = 13 (M) -> Position d'arrivée - Longitude
            '         = 14 (N) -> Position d'arrivée - Adresse
            '         = 15 (O) -> Niveau de carburant en %
            '         = 16 (P) -> Autonomie restante en km
            

  
    For Each item In jsonObject
        For Each item_item In item("trips")   ' Boucle pour récupérer les trajets
            
            Dim MaDate_UNIX_GMT_dep As Long, MaDate_DST_dep As Date  ' Pour convertir la date unix de départ en date excel
            Dim MaDate_UNIX_GMT_arr As Long, MaDate_DST_arr As Date  ' Pour convertir la date unix d'arrivée en date excel
            Dim duree_trajet As Long, duree_trajet_bis As Date
            Dim distance_trajet As Single, conso_trajet As Single, niveau_carb As Single
                                              
            ws.Cells(i, 1) = item_item("id")
      
            ' Récupération des dates
            ' On stocke les deux dates (départ et arrivée) car il faut déterminer le temps de parcours
            ' qui ne doit pas être dépendant d'un éventuelle changement d'heure en cours de route
            MaDate_UNIX_GMT_dep = item_item("startDateTime")            ' Date de départ
            MaDate_UNIX_GMT_arr = item_item("endDateTime")              ' Date d'arrivée
            MaDate_DST_dep = Date_UNIX_To_Date_DST(MaDate_UNIX_GMT_dep)
            MaDate_DST_arr = Date_UNIX_To_Date_DST(MaDate_UNIX_GMT_arr)
            'ws.Cells(i, 2) = Format(MaDate_DST_dep, "dd/mm/yy - hh:mm")
            'ws.Cells(i, 3) = Format(MaDate_DST_arr, "dd/mm/yy - hh:mm")
            ws.Cells(i, 2) = MaDate_DST_dep
            ws.Cells(i, 2).NumberFormat = "dd/mm/yy - hh:mm"
            ws.Cells(i, 3) = MaDate_DST_arr
            ws.Cells(i, 3).NumberFormat = "dd/mm/yy - hh:mm"
            ' Calcul de la durée du trajet en cours
            duree_trajet = MaDate_UNIX_GMT_arr - MaDate_UNIX_GMT_dep
            duree_trajet_bis = Date_UNIX_To_Date_GMT(MaDate_UNIX_GMT_arr) - Date_UNIX_To_Date_GMT(MaDate_UNIX_GMT_dep)
            'ws.Cells(i, 4) = Format(duree_trajet_bis, "h:mm")
            ws.Cells(i, 4) = duree_trajet_bis
            ws.Cells(i, 4).NumberFormat = "h:mm"
            distance_trajet = item_item("endMileage") - item_item("startMileage")
            ws.Cells(i, 5) = distance_trajet
            ws.Cells(i, 5).NumberFormatLocal = "0,0"
            ws.Cells(i, 6) = item_item("endMileage")
            ws.Cells(i, 6).NumberFormatLocal = "0"
            ws.Cells(i, 7) = item_item("consumption")
            conso_trajet = ws.Cells(i, 7)
            ws.Cells(i, 7).NumberFormatLocal = "0,00"
            ' Pour le calcul de la consommation moyenne, il faut éviter la division par zéro dans le cas où
            ' la voiture à tourner à l'arret, la distance parcourue est nulle
            If distance_trajet <> 0 Then
                ws.Cells(i, 8) = conso_trajet / distance_trajet * 100
                ws.Cells(i, 8).NumberFormatLocal = "0,0"
            Else
                ws.Cells(i, 8) = "//"
                ws.Cells(i, 8).NumberFormat = "General"
            End If

            ws.Cells(i, 9) = item_item("startPosLatitude")
            ws.Cells(i, 10) = item_item("startPosLongitude")
            
            Dim adresse_dep As String, adresse_arr As String
            adresse_dep = item_item("startPosAddress")
            ws.Cells(i, 11) = Correction_Adresse(adresse_dep)
            ws.Cells(i, 11).WrapText = False
            ws.Cells(i, 12) = item_item("endPosLatitude")
            ws.Cells(i, 13) = item_item("endPosLongitude")
            
            adresse_arr = item_item("endPosAddress")
            ws.Cells(i, 14) = Correction_Adresse(adresse_arr)
            ws.Cells(i, 14).WrapText = False
            ws.Cells(i, 15) = item_item("fuelLevel") / 100
            ws.Cells(i, 15).NumberFormat = "0 %"
            ws.Cells(i, 16) = item_item("fuelAutonomy")
            
            i = i + 1
            nbTrip = nbTrip + 1
            kilometrage = item_item("endMileage")
        Next
    
        ws.Range("B1") = item("vin")   ' On écrit le VIN récupéré
            
    Next
    
    ws.Range("G1") = nbTrip   ' On a maintenant le nombre de trajet
    ws.Range("G2") = kilometrage ' On écrit le kilométrage total de la voiture
    ws.Range("G2").NumberFormatLocal = "0"
    
    ' Calcul de la consommation totale du tableau
    Dim j As Integer
    j = 4 + nbTrip      ' valeur qui délimite la dernière cellule de la colonne G contenant une consommation
    Set MaPlage = ws.Range("G5:G" & j)
    conso_totale = ws.Application.WorksheetFunction.Sum(MaPlage)
    ws.Range("K1") = conso_totale   ' On a maintenant la consommation totale de tous les trajets
    ws.Range("K1").NumberFormatLocal = "0,00"
    ws.Range("K2").NumberFormatLocal = "0,0"

End Sub

Sub Effacage_Donnees()
    ' Fonction pour effacer les données avant d'utiliser celles du fichier ouvert
    With Worksheets("Trajets-MyPeugeot")
        .Range("B1") = ""           ' Le VIN
        .Range("N1") = ""           ' Le nom du fichier
        .Range("G1:G2") = ""        ' Le nombre de trajet et le kilométrage total
        .Range("A5:P10000") = ""    ' Le grand tableau de valeurs de tous les trajets effectués
        .Range("K1") = ""           ' La consommation totale sur tous les trajets
    End With
    With Worksheets("DATA")
        .Range("A2:D10000") = ""  ' Un tableau récapitulatif mensuel
        .Range("F4:F53") = 0      ' Le nombre de trajets regroupé en plage de distance
        .Range("H4:H53") = 0      ' La distance des regroupements de trajets
    End With
End Sub

Private Function Date_UNIX_To_Date_GMT(date_unix_GMT As Long) As Date
    ' Fonction qui converti un temps UNIX en date GMT
    '@PARAM {Long} date à convertir, au format UNIX GMT
    '@RETURN {Date} Renvoi la date convertie en date GMT

    Date_UNIX_To_Date_GMT = (date_unix_GMT / 86400) + DateValue("01/01/1970")

End Function

Private Function Date_UNIX_To_Date_DST(date_unix_GMT As Long) As Date
    ' Fonction qui converti un temps UNIX en date avec DST (changement d'heure)
    '@PARAM {Long} date à convertir, au format UNIX GMT
    '@RETURN {Date} Renvoi la date convertie en date GMT
    
    Dim DST_val As Integer, date_unix_DST As Long
            
    ' La date ainsi calculée ne tient pas compte du passage à l'heure d'été ou à l'heure d'hiver
    ' Il faut vérifier si le jour du mois de cette date est avant ou après le dernier dimanche de mars ou d'octbre.
    ' Rappel :  l'heure passe en GMT +2 au dernier dimanche de mars, c'est l'heure d'été
    '           l'heure passe en GMT +1 au dernier dimanche d'octobre, c'est l'heure d'hiver
    ' On va donc devoir ajouter soit 1h=3600s au temps GMT si c'est l'heure d'hiver, soit 2h=7200s au temps GMT si c'est l'heure d'hiver
    ' Déterminons la valeur du facteur DST
    DST_val = DST(date_unix_GMT)
    
    ' Calcul de la nouvelle date avec DST
    date_unix_DST = date_unix_GMT + DST_val * 3600
    ' Conversion en Date de cette date UNIX : la nouvelle Date est maintenant DST
    Date_UNIX_To_Date_DST = Date_UNIX_To_Date_GMT(date_unix_DST)

End Function

Private Function DST(date_unix_GMT As Long) As Integer
    ' Fonction qui détermine le modificateur d'heure (Day Saving Time) à appliquer à l'heure GMT pour avoir l'heure FR
    ' Ce sera le modificateur d'heure par rapport au temps unix GMT :   1 = pour l'heure d'hiver
    '                                                                   2 = pour l'heure d'été
    '@PARAM {Long} date à tester, au format unix GMT
    '@RETURN {Integer} Renvoi un entier permettant la modification de l'heure unix


    ' On déclare les variables utilisées pour le jour, le mois, l'année, l'heure en temps GMT
    Dim jour_GMT As Integer, mois_GMT As Integer, annee_GMT As Integer, heure_GMT As Integer, minutes_GMT As Integer
    Dim date_temp As Date
    
    ' On déclare les variables utilisées pour le dernier dimanche de mars ou d'octobre : jour/mois/heure
    Dim jour_dD As Integer, mois_dD As Integer, heure_dD As Integer
    Dim num_jour_31 As Integer      ' C'est pour stocker le n° du jour de semaine pour le 31/03/annee ou 31/10/annee
   
    ' On convertir la date unix en date GMT
    date_temp = Date_UNIX_To_Date_GMT(date_unix_GMT)
    
    ' On récupère le jour, le mois, l'année, l'heure en temps GMT
    jour_GMT = Day(date_temp)
    mois_GMT = Month(date_temp)
    annee_GMT = Year(date_temp)
    heure_GMT = Hour(date_temp)
    ' minutes = Minute(date_temp)       ' Inutile ici
     
    Select Case mois_GMT
        Case 1, 2, 11, 12
            ' On est dans le cas où l'heure d'hiver est appliquée : de Novembre à Février
            DST = 1
            
        Case 4 To 9
            ' On est dans le cas où l'heure d'été est appliquée : de Avril à Septembre
            DST = 2
            
        Case 3
            ' On est en mars, il faut vérifier que le jour en question est avant ou après le dernier dimanche de mars
            ' Détermination du n° du jour de la semaine du dernier jour de mars
            num_jour_31 = Weekday("31/03/" & annee_GMT, vbMonday)
            If num_jour_31 = 7 Then
                ' Le 31 c'est le dimanche
                jour_dD = 7
            Else
                jour_dD = 31 - num_jour_31
            End If
                            
            If jour_GMT < jour_dD Then
                ' On est avant le dernier dimanche du mois, donc encore en heure d'hiver
                DST = 1
            ElseIf jour_GMT > jour_dD Then
                ' On est après le dernier dimanche du mois, donc en heure d'été
                DST = 2
            ElseIf jour_GMT = jour_dD Then
                ' On est le dernier dimanche du mois, donc il faut voir si on est avant 1h du matin ou pas.
                If heure_GMT < 1 Then
                    ' On est encore à l'heure d'hiver
                    DST = 1
                Else
                    ' On est passé à l'heure d'été
                    DST = 2
                End If
            End If
            
        Case 10
            ' On est en octobre, il faut vérifier que le jour en question est avant ou après le dernier dimanche de mars
            ' Détermination du n° du jour de la semaine du dernier jour d'octobre
            num_jour_31 = Weekday("31/10/" & annee_GMT, vbMonday)
            jour_dD = 31 - num_jour_31
            If num_jour_31 = 7 Then
                ' Le 31 c'est le dimanche, c'est donc lui le dernier dimanche du mois !
                jour_dD = 7
            Else
                ' Le 31 est un autre jour de la semaine, on calcule donc quel sera le jour XX du dernier dimanche du mois
                jour_dD = 31 - num_jour_31
            End If
            
            If jour_GMT < jour_dD Then
                ' On est avant le dernier dimanche du mois, donc encore en heure d'hiver
                DST = 1
            ElseIf jour_GMT > jour_dD Then
                ' On est après le dernier dimanche du mois, donc en heure d'été
                DST = 2
            ElseIf jour_GMT = jour_dD Then
                ' On est le dernier dimanche du mois, donc il faut voir si on est avant 1h du matin ou pas.
                If heure_GMT < 1 Then
                    ' On est encore à l'heure d'été
                    DST = 2
                Else
                    ' On est passé à l'heure d'hiver
                    DST = 1
                End If
            End If
        End Select
        
End Function

Private Function Correction_Adresse(ByVal adresse As String) As String

    Dim i As Integer
    Dim lettre As String * 2

    adresse = Replace(adresse, vbLf, ", ")   ' On remplace tous les retours à la ligne "\n" par des ", "


    For i = 1 To Len(adresse) - 1
        lettre = Mid(adresse, i, 2)
        Select Case lettre
            Case "Ã¨"   ' è
                adresse = Left(adresse, i - 1) + "è" + Mid(adresse, i + 2, 100)
            Case "Ã©"   ' é
                adresse = Left(adresse, i - 1) + "é" + Mid(adresse, i + 2, 100)
            Case "Ãª"   ' ê
                adresse = Left(adresse, i - 1) + "ê" + Mid(adresse, i + 2, 100)
            Case "Ã€"   ' À
                adresse = Left(adresse, i - 1) + "À" + Mid(adresse, i + 2, 100)
            Case "Ã¢"   ' â
                adresse = Left(adresse, i - 1) + "â" + Mid(adresse, i + 2, 100)
            Case "Å“"   ' oe
                adresse = Left(adresse, i - 1) + "œ" + Mid(adresse, i + 2, 100)
            Case "Ã "   ' à
                adresse = Left(adresse, i - 1) + "à" + Mid(adresse, i + 2, 100)
            Case "Ã§"   ' ç
                adresse = Left(adresse, i - 1) + "ç" + Mid(adresse, i + 2, 100)
            Case "Ã¯"   ' ï
                adresse = Left(adresse, i - 1) + "ï" + Mid(adresse, i + 2, 100)
            Case "Ã¹"   ' ù
                adresse = Left(adresse, i - 1) + "ù" + Mid(adresse, i + 2, 100)
            Case "Ã¼"   ' ü
                adresse = Left(adresse, i - 1) + "ü" + Mid(adresse, i + 2, 100)
            Case "Ã‰"   ' É
                adresse = Left(adresse, i - 1) + "É" + Mid(adresse, i + 2, 100)
            Case "Ãˆ"   ' È
                adresse = Left(adresse, i - 1) + "È" + Mid(adresse, i + 2, 100)
        End Select
    Next i
    
    Correction_Adresse = adresse

End Function


Sub test()

    
    Dim MaDate_UNIX_GMT As Long, MaDate_GMT As Date, MaDate_DST As Date  ' Pour convertir une date unix en date excel

    
    ' pour les essais = 25/10/2020 à 0h50min00s  ==  1603587000 UTC-UNIX <- c'est censé être encore l'heure d'été
    ' pour les essais = 29/03/2020 à 0h50min00s  ==  1585443000 UTC-UNIX <- c'est censé être encore l'heure d'hiver
    MaDate_UNIX_GMT = 1603590600
    'MaDate_GMT = (MaDate_UNIX_GMT / 86400) + DateValue("01/01/1970")

    MaDate_DST = Date_UNIX_To_Date_DST(MaDate_UNIX_GMT)
    
    MsgBox "Date UNIX GMT = " & MaDate_GMT & vbCrLf & "Date DST = " & MaDate_DST

End Sub


















