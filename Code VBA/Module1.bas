Attribute VB_Name = "Module1"
' Licence utilisée :
'                   GNU AFFERO GENERAL PUBLIC LICENSE
'                      Version 3, 19 November 2007
'
' Dépôt GitHub : https://github.com/MilesTEG1/Importation-Excel-Donnees-MyPeugeot/blob/master/LICENSE.md
'
' @author MilesTEG1@gmail.com
' @license  AGPL-3.0 (https://www.gnu.org/licenses/agpl-3.0.fr.html)
'
' Déinissons quelques constantes qui serviront pour les colonnes/lignes/plages de cellules.
'
Const L_premiere_valeur As Integer = 5      ' Première ligne à contenir des données (avant ce sont les lignes d'en-tête
Const C_vin             As Integer = 1      ' COLONNE = 1 (A) -> VIN pour le trajet (il est possible d'avoir plusieurs VIN dans le fichier de donnée).
Const C_id              As Integer = 2      '         = 2 (B) -> Trip ID#
Const C_date_dep        As Integer = 3      '         = 3 (C) -> Date de Départ (à déterminer avec une conversion)
Const C_date_arr        As Integer = 4      '         = 4 (D) -> Date de Fin    (à déterminer avec une conversion)
Const C_duree           As Integer = 5      '         = 5 (E) -> Durée du trajet  (à calculer)
Const C_dist            As Integer = 6      '         = 6 (F) -> Distance du trajet en km
Const C_dist_tot        As Integer = 7      '         = 7 (G) -> Distance totale au compteur km
Const C_conso           As Integer = 8      '         = 8 (H) -> Consommation du tajet en L
Const C_conso_moy       As Integer = 9      '         = 9 (I) -> Consommation moyenne en L/100km
Const C_pos_dep_lat     As Integer = 10     '         = 10 (J) -> Position de départ - Latitude
Const C_pos_dep_long    As Integer = 11     '         = 11 (K) -> Position de départ - Longitude
Const C_pos_dep_adr     As Integer = 12     '         = 12 (L) -> Position de départ - Adresse
Const C_pos_arr_lat     As Integer = 13     '         = 13 (M) -> Position d'arrivée - Latitude
Const C_pos_arr_long    As Integer = 14     '         = 14 (N) -> Position d'arrivée - Longitude
Const C_pos_arr_adr     As Integer = 15     '         = 15 (O) -> Position d'arrivée - Adresse
Const C_niv_carb        As Integer = 16     '         = 16 (P) -> Niveau de carburant en %
Const C_auto            As Integer = 17     '         = 17 (Q) -> Autonomie restante en km
Const CELL_vin_entete   As String = "B1"    ' Cellule qui contiendra le VIN associé aux données de l'en-tête
Const CELL_nb_trips     As String = "H1"    ' Cellule qui contiendra le nombre de trajets totale (tous VIN confondus)
Const CELL_fichierMYP   As String = "O1"    ' Cellule qui contiendra le nom du fichier importé
Const CELL_km           As String = "H2"    ' Cellule qui contiendra le kilométrage actuel de la voiture
Const CELL_conso_tot    As String = "L1"    ' Cellule qui contiendra la consommation totale en L
Const CELL_conso_tot_moy    As String = "L2"    ' Cellule qui contiendra la consommation moyenne totale pour tous les trajets
Const CELL_plage_donnees    As String = "A" & L_premiere_valeur & ":Q20000" ' Plage de cellules contenant les données
Const CELL_plage_conso_tot  As String = "H" & L_premiere_valeur & ":H"      ' Cellule qui contiendra le kilométrage actuel de la voiture

' V1.5 : MultiVIN
Const G_Nb_VIN_Max = 20                     ' Nb VIN max traité par cette macro
Const G_Nb_Trajets_Max = 20000              ' Nb trajets max par VIN traités par cette macro

Const VERSION As String = "Version du fichier" & vbLf & "v1.6"    ' Version du fchier
Const CELL_ver As String = "D1:D2"     ' Cellule où afficher la version du fichier
'
' Fin de déclaration des constantes
'

'
' Fonction pour lire les données depuis un fichier JSON et les écrire dans le tableau
'
Sub MYP_JSON_Decode()
    Dim jsonText As String
    Dim jsonObject As Object, item As Object, item_item As Object
    Dim i, j, k As Long
    Dim ws As Worksheet
    
    Dim FSO As New FileSystemObject
    Dim JsonTS As TextStream
    
    Dim FichierMYP As Variant   ' nom du fichier myp
    Dim nbTrip, nbTripReel As Integer ' nombre de trajets
    Dim kilometrage As Single   ' Pour stocker le kilométrage total de la voiture.
    Dim conso_totale As Single  ' Pour stocker la consommation totale de tout le tableau
    
' V1.5 : pour MultiVIN
    Dim l_Tab_Vin As Integer                 ' pour boucle sur les vin
    Dim Nb_VIN As Integer                    ' Nb VIN trouvés dans le fichier .myp
    Dim Liste_VIN(G_Nb_VIN_Max, 2) As String ' Tableau interne des VIN trouvés (col 1 = vin, col 2 = retenu)
    Dim VIN_Actuel As String                 ' VIN associé à la boucle des trajets
    Dim VIN_A_Traiter As Boolean
' V1.6 : pour ne pas effacer les données au début
    Dim T_Trajets_existants(G_Nb_Trajets_Max) As String ' Tableau des trajets trouvés existants dans l'onglet
    Dim Derniere_Ligne_remplie As Integer               ' Numéro de la dernière ligne avec des données
    Dim Nb_Trajets As Long
    Dim Trajet_Trouvé As Boolean
    
    Set ws = Worksheets("Trajets-MyPeugeot")
    EcrireValeurFormat cell:=ws.Range(CELL_ver), val:=VERSION, f_size:=10, wrap:=True
    
    JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True
    
    FichierMYP = Application.GetOpenFilename("Fichiers trajets Peugeot App (*.myp),*.myp,Fichiers trajets Citroen App (*.myc),*.myc,Fichiers trajets DS App (*.myd),*.myd")  ' On demande la sélection du fichier
    If FichierMYP = False Then
        MsgBox "Aucun fichier n'a été selectionné !", vbCritical
        Exit Sub
    End If
    Set JsonTS = FSO.OpenTextFile(FichierMYP, ForReading)
    jsonText = JsonTS.ReadAll
    JsonTS.Close
    
    ' Comme le fichier existe, on efface tout
    ' Effacage_Donnees     retiré en 1.6
    
    nbTrip = 0    ' On réinitialise le nombre de trajets
    EcrireValeurFormat cell:=ws.Range(CELL_fichierMYP), val:=FichierMYP, wrap:=True
    Set jsonObject = JsonConverter.ParseJson(jsonText)
    
    ' Recherche de tous les VIN
    Nb_VIN = 0
    For Each item In jsonObject
        If item("vin") <> "" Then
            Nb_VIN = Nb_VIN + 1
            Liste_VIN(Nb_VIN, 1) = item("vin")
            Liste_VIN(Nb_VIN, 2) = False
        End If
    Next
    
    ' Si aucun VIN dans le fichier myp, on sort !
    If Nb_VIN = 0 Then
        MsgBox "Aucun VIN trouvé dans votre fichier .myp. Pb de structure ?", vbCritical
        Exit Sub
    End If
    
    ' Affichage de la forme de choix des VIN à importer, seulement si plus de 1 VIN. Sinon, celui-ci devient défaut
    If Nb_VIN > 1 Then
        ' Remise à zéro de la liste des choix
        For l_Tab_Vin = FormeVIN.FormeVIN_ListeVIN.ListCount - 1 To 0 Step -1
            FormeVIN.FormeVIN_ListeVIN.RemoveItem (l_Tab_Vin)
        Next l_Tab_Vin
        
        ' Ajout des VIN trouvés
        For l_Tab_Vin = 1 To Nb_VIN
          FormeVIN.FormeVIN_ListeVIN.AddItem (Liste_VIN(l_Tab_Vin, 1))
        Next l_Tab_Vin
        
        ' activation de la forme de choix des VIN
        FormeVIN.Show
        
        ' Si bouton annuler = on quitte la procédure
        If FormeVIN.BoutonChoisi.Value = 2 Then
          MsgBox "Vous avez annulé. On quitte !", vbCritical
          Exit Sub
        End If
        ' On parcourt la liste pour récupérer les VIN sélectionnés
        For l_Tab_Vin = 0 To FormeVIN.FormeVIN_ListeVIN.ListCount - 1
          If FormeVIN.FormeVIN_ListeVIN.Selected(l_Tab_Vin) Then
            Liste_VIN(l_Tab_Vin + 1, 2) = True
          End If
        Next
    Else  ' cas 1 seul VIN présent dans le fichier
        Liste_VIN(1, 2) = True
    End If

' V1.6 : stockage dans un tableau interne (que l'on vide d'abord) de tous les trajets déjà dans l'Excel
    For i = 1 To G_Nb_Trajets_Max
        T_Trajets_existants(i) = ""
    Next i
    Derniere_Ligne_remplie = Cells(Columns(1).Cells.Count, 1).End(xlUp).Row
    Nb_Trajets = 0
    For i = L_premiere_valeur To Derniere_Ligne_remplie
        T_Trajets_existants(i - L_premiere_valeur + 1) = ws.Cells(i, C_vin) & ";" & ws.Cells(i, C_id)
        Nb_Trajets = Nb_Trajets + 1
    Next i

    i = L_premiere_valeur + Nb_Trajets   ' On défini un compteur qui sert à se positionner sur la ligne où les données doivent être écrites.
                            ' Le n° de la ligne où on va commencer à écrire les données
                            ' ws.Cells(LIGNE, COLONNE)     Où LIGNE commence à 1
                            '                              Où COLONNE commence à 1 = A
                            ' On va utiliser les constantes définie avant la fonction pour écrire/formater/effacer une cellule
                            
    For Each item In jsonObject
' V1.5 : 1ère vérif à faire : que le VIN du trajet soit dans les VIN à importer
        VIN_Actuel = item("vin")
        VIN_A_Traiter = False
        For l_Tab_Vin = 1 To Nb_VIN
            If Liste_VIN(l_Tab_Vin, 1) = VIN_Actuel Then
                VIN_A_Traiter = Liste_VIN(l_Tab_Vin, 2)
            End If
        Next l_Tab_Vin
      
' V1.5 : on ne traite les trajets QUE SI Vin_A_Traiter est vrai
        If VIN_A_Traiter Then
            For Each item_item In item("trips")   ' Boucle pour récupérer les trajets
                
                Dim MaDate_UNIX_GMT_dep As Long, MaDate_DST_dep As Date  ' Pour convertir la date unix de départ en date excel
                Dim MaDate_UNIX_GMT_arr As Long, MaDate_DST_arr As Date  ' Pour convertir la date unix d'arrivée en date excel
                Dim duree_trajet As Long, duree_trajet_bis As Date
                Dim distance_trajet As Single, conso_trajet As Single, niveau_carb As Single
                                              
' V1.6 : à ne faire que si VIN et Id n'étaient pas déjà présents
                Trajet_Trouvé = False
                For j = 1 To Nb_Trajets
                    If T_Trajets_existants(j) = VIN_Actuel & ";" & item_item("id") Then
                        Trajet_Trouvé = True
                        Exit For
                    End If
                Next j
                If Not Trajet_Trouvé Then
                    With ws.Cells(i, C_vin)
                        .Value = VIN_Actuel   ' On écrit le VIN récupéré
                        .Font.Size = 8
                        .NumberFormat = "General"
                    End With
                
                    EcrireValeurFormat cell:=ws.Cells(i, C_vin), val:=VIN_Actuel, f_size:=8    ' On écrit le VIN récupéré
                    EcrireValeurFormat cell:=ws.Cells(i, C_id), val:=item_item("id")            ' On écrit l'ID récupéré
                    
                    ' Récupération des dates
                    ' On stocke les deux dates (départ et arrivée) car il faut déterminer le temps de parcours
                    ' qui ne doit pas être dépendant d'un éventuelle changement d'heure en cours de route
                    MaDate_UNIX_GMT_dep = item_item("startDateTime")            ' Date de départ
                    MaDate_UNIX_GMT_arr = item_item("endDateTime")              ' Date d'arrivée
                    MaDate_DST_dep = Date_UNIX_To_Date_DST(MaDate_UNIX_GMT_dep)
                    MaDate_DST_arr = Date_UNIX_To_Date_DST(MaDate_UNIX_GMT_arr)
                    
                    EcrireValeurFormat cell:=ws.Cells(i, C_date_dep), val:=MaDate_DST_dep, n_format:="date"
                    EcrireValeurFormat cell:=ws.Cells(i, C_date_arr), val:=MaDate_DST_arr, n_format:="date"
                    
                    ' Calcul de la durée du trajet en cours
                    duree_trajet = MaDate_UNIX_GMT_arr - MaDate_UNIX_GMT_dep
                    duree_trajet_bis = Date_UNIX_To_Date_GMT(MaDate_UNIX_GMT_arr) - Date_UNIX_To_Date_GMT(MaDate_UNIX_GMT_dep)
                    
                    EcrireValeurFormat cell:=ws.Cells(i, C_duree), val:=duree_trajet_bis, n_format:="duree"
                    
                    distance_trajet = item_item("endMileage") - item_item("startMileage")
                    EcrireValeurFormat cell:=ws.Cells(i, C_dist), val:=distance_trajet, n_format:="1"
                    EcrireValeurFormat cell:=ws.Cells(i, C_dist_tot), val:=item_item("endMileage"), n_format:="0"
                    
                    EcrireValeurFormat cell:=ws.Cells(i, C_conso), val:=item_item("consumption"), n_format:="2"
                    conso_trajet = ws.Cells(i, C_conso)
                    'ws.Cells(i, C_conso) = item_item("consumption")
                    'conso_trajet = ws.Cells(i, C_conso)
                    
                    'ws.Cells(i, C_conso).NumberFormatLocal = "0,00"
                    
                    ' Pour le calcul de la consommation moyenne, il faut éviter la division par zéro dans le cas où
                    ' la voiture à tourner à l'arret, la distance parcourue est nulle
                    If distance_trajet <> 0 Then
                        EcrireValeurFormat cell:=ws.Cells(i, C_conso_moy), val:=conso_trajet / distance_trajet * 100, n_format:="1"
                        'ws.Cells(i, C_conso_moy) = conso_trajet / distance_trajet * 100
                        'ws.Cells(i, C_conso_moy).NumberFormatLocal = "0,0"
                    Else
                        EcrireValeurFormat cell:=ws.Cells(i, C_conso_moy), val:="//"
                        'ws.Cells(i, C_conso_moy) = "//"
                        'ws.Cells(i, C_conso_moy).NumberFormat = "General"
                    End If
        
                    EcrireValeurFormat cell:=ws.Cells(i, C_pos_dep_lat), val:=item_item("startPosLatitude")
                    EcrireValeurFormat cell:=ws.Cells(i, C_pos_dep_long), val:=item_item("startPosLongitude")
                                
                    Dim adresse_dep As String, adresse_arr As String
                    adresse_dep = item_item("startPosAddress")
                    EcrireValeurFormat cell:=ws.Cells(i, C_pos_dep_adr), val:=Correction_Adresse(adresse_dep)
                    
                    EcrireValeurFormat cell:=ws.Cells(i, C_pos_arr_lat), val:=item_item("endPosLatitude")
                    EcrireValeurFormat cell:=ws.Cells(i, C_pos_arr_long), val:=item_item("endPosLongitude")
                    
                    adresse_arr = item_item("endPosAddress")
                    EcrireValeurFormat cell:=ws.Cells(i, C_pos_arr_adr), val:=Correction_Adresse(adresse_arr)
                    
                    EcrireValeurFormat cell:=ws.Cells(i, C_niv_carb), val:=item_item("fuelLevel") / 100, n_format:="%"
        
                    EcrireValeurFormat cell:=ws.Cells(i, C_auto), val:=item_item("fuelAutonomy")
                    
                    i = i + 1
                    nbTrip = nbTrip + 1
        
                    kilometrage = item_item("endMileage")   ' On stocke le kilométrage de fin de trajet pour être affiché en tant que kilométrage actuel lors du dernier trajet
' V1.6 : fin du test
                End If
            Next
    
' V .1.5 : fin du test
        End If
    Next
    
' V1.6 : tri final sur colonnes A puis B
    ws.Range("A5:Q5").Select
    ws.Range(Selection, Selection.End(xlDown)).Select
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=Range("A5:A65536"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ws.Sort.SortFields.Add Key:=Range("B5:B65536"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Trajets-MyPeugeot").Sort
        .SetRange Range("A5:Q65536")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ws.Cells(1, 1).Select
    EcrireValeurFormat cell:=ws.Range(CELL_nb_trips), val:=nbTrip, f_size:=12                   ' On écrit le nombre de trajet
    EcrireValeurFormat cell:=ws.Range(CELL_km), val:=kilometrage, n_format:="0", f_size:=12     ' On écrit le kilométrage total de la voiture

    
    ' Calcul de la consommation totale du tableau
    j = 4 + nbTrip      ' valeur qui délimite la dernière cellule de la colonne G contenant une consommation
    'Set MaPlage = ws.Range("H5:H" & j)
    'conso_totale = ws.Application.WorksheetFunction.Sum(ws.Range("H5:H" & j))
    
    ' On écrit la valeur calculée de la consommation totale de tous les trajets
    EcrireValeurFormat cell:=ws.Range(CELL_conso_tot), val:=ws.Application.WorksheetFunction.Sum(ws.Range("H5:H" & j)), n_format:="2", f_size:=12
    
    'ws.Range(CELL_conso_tot) = conso_totale   ' On a maintenant la consommation totale de tous les trajets
    'ws.Range(CELL_conso_tot).NumberFormatLocal = "0,00"
    
    ws.Range(CELL_conso_tot_moy).NumberFormatLocal = "0,0"

End Sub

Private Sub EcrireValeurFormat(cell As Variant, val As Variant, Optional n_format As String = "General", Optional f_size As Integer = 10, Optional wrap As Boolean = False)
    ' Fonction pour écrire une valeur dans une cellule
    ' Arguments obligatoires :  cellule As Variant  <- La cellule ou plage de cellule devant être modifiée
    '                           val As Variant      <- La valeur à écrire dans la cellule/plage
    ' Arguments optionels :     n_format As String = ""         <- Le format NumberFormat, défaut = "Genral"
    '                                                           <- Valeurs = "date" pour format date
    '                                                           <- Valeurs = "duree" pour format durée
    '                                                           <- Valeurs = "0" pour format numérique Local sans virgule
    '                                                           <- Valeurs = "1" pour format numérique Local avec 1 chiffre après la virgule
    '                                                           <- Valeurs = "2" pour format numérique Local avec 2 chiffres après la virgule
    '                           font_size As Integer = 10       <- La taille de caractère, défaut = 10
    '                           wrap As Boolean = False         <- Retour à la ligne dans la cellule, défaut = faux
    With cell
        
        .Value = val
        
        Select Case n_format
            Case "date"
                .NumberFormat = "dd/mm/yy - hh:mm"
            Case "duree"
                .NumberFormat = "h:mm"
            Case "0"
                .NumberFormatLocal = "0"
            Case "1"
                .NumberFormatLocal = "0,0"
            Case "2"
                .NumberFormatLocal = "0,00"
            Case "%"
                .NumberFormat = "0 %"
            Case Else
                NumberFormat = "General"
        End Select
        
        If (f_size <> 10) Then
            .Font.Size = f_size
        Else
            .Font.Size = 10
        End If
        
        If (wrap = True) Then
            .WrapText = True
        Else
            .WrapText = False
        End If
        
    End With
    
End Sub

Sub Effacage_Donnees()
    ' Fonction pour effacer les données avant d'utiliser celles du fichier ouvert
    With Worksheets("Trajets-MyPeugeot")
        .Range(CELL_vin_entete) = ""           ' Le VIN
        .Range(CELL_fichierMYP) = ""           ' Le nom du fichier
        .Range(CELL_nb_trips) = ""  ' Le nombre de trajet et le kilométrage total
        .Range(CELL_km) = ""
        .Range(CELL_plage_donnees) = ""    ' Le grand tableau de valeurs de tous les trajets effectués
        .Range(CELL_conso_tot) = ""           ' La consommation totale sur tous les trajets
    End With
    With Worksheets("Donnees-Recap")
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
