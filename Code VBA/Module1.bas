Attribute VB_Name = "Module1"
' Licence utilis�e :
'                   GNU AFFERO GENERAL PUBLIC LICENSE
'                      Version 3, 19 November 2007
'
' D�p�t GitHub : https://github.com/MilesTEG1/Importation-Excel-Donnees-MyPeugeot
'
' @authors :    MilesTEG1@gmail.com
'               avec les conseils de W13-FP
' @license  AGPL-3.0 (https://www.gnu.org/licenses/agpl-3.0.fr.html)
'
' Suivi des versions
'       - V 1.5 : Multi-VIN
'       - V 1.6 : Ne pas effacer les donn�es au d�but
'       - V 1.7 : Optimisation temps ex�cution
'       - V 1.8 : Tableau crois� dynamique et form information avancement
'       - V 1.9 : Gestion des VIN d�j� connus
'       - V 1.9.1 : Correction de quelques bugs, et am�lioration de la feuille Accueil
'
' Couples de versions d'Excel & OS test�es :
'       - Windows 10 v1909 (18363.752) & Excel pour Office 365 Version 2003 (build 12624.20382)
'       - Windows 10 v1809 & Excel 2016
'       - Windows 10 v1909 & Excel 2019
'
Option Explicit     ' On force � d�clarer les variables
'
' D�inissons quelques constantes qui serviront pour les colonnes/lignes/plages de cellules.
'
' Constantes pour la feuille "Trajets-MyPeugeot"
Const L_Premiere_Valeur As Integer = 3      ' Premi�re ligne � contenir des donn�es (avant ce sont les lignes d'en-t�te
Const C_vin             As Integer = 1      ' COLONNE = 1 (A) -> VIN pour le trajet (il est possible d'avoir plusieurs VIN dans le fichier de donn�e).
Const C_id              As Integer = 2      '         = 2 (B) -> Trip ID#
Const C_date_dep        As Integer = 3      '         = 3 (C) -> Date de D�part (� d�terminer avec une conversion)
Const C_date_arr        As Integer = 4      '         = 4 (D) -> Date de Fin    (� d�terminer avec une conversion)
Const C_duree           As Integer = 5      '         = 5 (E) -> Dur�e du trajet  (� calculer)
Const C_dist            As Integer = 6      '         = 6 (F) -> Distance du trajet en km
Const C_dist_tot        As Integer = 7      '         = 7 (G) -> Distance totale au compteur km
Const C_conso           As Integer = 8      '         = 8 (H) -> Consommation du tajet en L
Const C_conso_moy       As Integer = 9      '         = 9 (I) -> Consommation moyenne en L/100km
Const C_pos_dep_lat     As Integer = 10     '         = 10 (J) -> Position de d�part - Latitude
Const C_pos_dep_long    As Integer = 11     '         = 11 (K) -> Position de d�part - Longitude
Const C_pos_dep_adr     As Integer = 12     '         = 12 (L) -> Position de d�part - Adresse
Const C_pos_arr_lat     As Integer = 13     '         = 13 (M) -> Position d'arriv�e - Latitude
Const C_pos_arr_long    As Integer = 14     '         = 14 (N) -> Position d'arriv�e - Longitude
Const C_pos_arr_adr     As Integer = 15     '         = 15 (O) -> Position d'arriv�e - Adresse
Const C_niv_carb        As Integer = 16     '         = 16 (P) -> Niveau de carburant en %
Const C_auto            As Integer = 17     '         = 17 (Q) -> Autonomie restante en km
'Const CELL_vin_entete   As String = "B" & (L_Premiere_Valeur - 3)    ' Cellule qui contiendra le VIN associ� aux donn�es de l'en-t�te
'Const CELL_nb_trips     As String = "H" & (L_Premiere_Valeur - 3)    ' Cellule qui contiendra le nombre de trajets totale (tous VIN confondus)
Const CELL_fichierMYP   As String = "G" & (L_Premiere_Valeur - 2)  ' Cellule qui contiendra le nom du fichier import�
'Const CELL_km           As String = "H" & (L_Premiere_Valeur - 2)   ' Cellule qui contiendra le kilom�trage actuel de la voiture
'Const CELL_conso_tot    As String = "L" & (L_Premiere_Valeur - 3)   ' Cellule qui contiendra la consommation totale en L
'Const CELL_conso_tot_moy    As String = "L" & (L_Premiere_Valeur - 2)   ' Cellule qui contiendra la consommation moyenne totale pour tous les trajets
Const CELL_plage_donnees    As String = "A" & L_Premiere_Valeur & ":Q40000" ' Plage de cellules contenant les donn�es
'Const CELL_plage_conso_tot  As String = "H" & L_Premiere_Valeur & ":H"      ' Cellule qui contiendra le kilom�trage actuel de la voiture
Const CELL_plage_1ereligne  As String = "A" & L_Premiere_Valeur & ":Q" & L_Premiere_Valeur  ' Toute la 1ere ligne de donn�e pour en faire le tri
Const CELL_plage_max_Donnees As String = "A" & L_Premiere_Valeur & ":Q65536"    ' La plage de donn�es maximale possible
Const CELL_plage_max_COL_vin As String = "A" & L_Premiere_Valeur & ":A65536"    ' La plage de donn�es maximale possible
Const CELL_plage_max_COL_id As String = "B" & L_Premiere_Valeur & ":B65536"    ' La plage de donn�es maximale possible
' La formule qui permet de calculer la consommation moyenne, qui doit �tre : =SI(L1=0;"Non dispo.";L1/(H2-G5)*100)
'Const FORMULE_calcul_Conso_tot_moy As Variant = "=SI(" & CELL_conso_tot & "=0;""Non dispo."";" & CELL_conso_tot & "/(H8-G10)*100)"
' V1.5 : MultiVIN
Const G_Nb_VIN_Max = 20                     ' Nb VIN max trait� par cette macro
Const G_Nb_Trajets_Max = 20000              ' Nb trajets max par VIN trait�s par cette macro


' Constantes pour la feuille Accueil
Const C_entete_ListeVIN        As Integer = 13     ' = 13 (M) Colonne d'ent�te des VIN dans la liste des vins r�cup�r�s, la colonne des descriptions des v�hicules est celle d'� cot� : 13+1 = N
Const L_entete_ListeVIN        As Integer = 3      ' Ligne d'ent�te des VIN dans la liste des vins r�cup�r�s, elle correspond aussi � celles des descriptions des v�hicules
Const VERSION As String = "v1.9.1"    ' Version du fchier
Const CELL_ver As String = "B3"     ' Cellule o� afficher la version du fichier
'
' Fin de d�claration des constantes
'

'
' Fonction pour lire les donn�es depuis un fichier JSON et les �crire dans le tableau
'
Sub MYP_JSON_Decode()
    Dim jsonText As String
    Dim jsonObject As Object, item As Object, item_item As Object
    Dim i, j, k As Long     ' Variables utilis�es pour les compteurs de boucles For
    Dim ws As Worksheet, ws_acc As Worksheet
        
    Dim FSO As New FileSystemObject
    Dim JsonTS As TextStream
    
    Dim FichierMYP As Variant   ' nom du fichier myp
    Dim nbTrip, nbTripReel As Integer ' nombre de trajets
    Dim kilometrage As Single   ' Pour stocker le kilom�trage total de la voiture.
    Dim conso_totale As Single  ' Pour stocker la consommation totale de tout le tableau
    Dim CheminFichier As String
    Dim MaDate_UNIX_GMT_dep As Long, MaDate_DST_dep As Date  ' Pour convertir la date unix de d�part en date excel
    Dim MaDate_UNIX_GMT_arr As Long, MaDate_DST_arr As Date  ' Pour convertir la date unix d'arriv�e en date excel
    Dim duree_trajet As Long, duree_trajet_bis As Date
    Dim distance_trajet As Single, conso_trajet As Single, niveau_carb As Single
    Dim adresse_dep As String, adresse_arr As String
    Dim derniere_val_tab    ' Derni�re ligne qui contiendra un trajet
        
' V1.5 : pour MultiVIN
    Dim l_Tab_Vin As Integer                    ' pour boucle sur les vin
    Dim Nb_VIN As Integer                       ' Nb VIN trouv�s dans le fichier .myp
    Dim Liste_VIN(G_Nb_VIN_Max, 2) As String    ' Tableau interne des VIN trouv�s (col 1 = vin, col 2 = retenu)
    Dim VIN_Actuel As String                    ' VIN associ� � la boucle des trajets
    Dim VIN_A_Traiter As Boolean                ' Pour d�finir le chemin par d�faut o� chercher le fichier de donn�e
' V1.6 : pour ne pas effacer les donn�es au d�but
    Dim T_Trajets_existants(G_Nb_Trajets_Max) As String ' Tableau des trajets trouv�s existants dans l'onglet
    Dim Derniere_Ligne_remplie As Integer               ' Num�ro de la derni�re ligne avec des donn�es
    Dim Nb_Trajets As Long
    Dim Trajet_Trouve As Boolean
' V1.9 : Gestion des VIN
    Dim Infos_VIN() As Variant
    Dim R As Range
    Dim Info_Voiture As String
    Dim Nb_VIN_renseignes As Integer
    Dim i_TCD
    
    Set ws = Worksheets("Trajets-MyPeugeot")
    Set ws_acc = Worksheets("Accueil")
    
    'Sheets("Accueil").Activate
    ws_acc.Activate
    
    EcrireValeurFormat cell:=Range(CELL_ver).Offset(-1, 0), val:="Version du fichier", f_size:=10, wrap:=True
    EcrireValeurFormat cell:=Range(CELL_ver), val:=VERSION, f_size:=16, wrap:=True
' V1.8 : activation de cette feuille pour �tre s�r d'�tre dedans
    Sheets("Trajets-MyPeugeot").Activate
    
    JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True
    
    CheminFichier = ActiveWorkbook.Path & "\"
    ' Il y a une erreur si on travail dans le dossier OneDrive, le CheminFicher est un lien du type https://d.docs.live.net/8f87e4...
    ' Il faut donc v�rifier si le d�but chaine de caract�re n'est pas https://
    If (InStr(1, CheminFichier, "https://", vbTextCompare) <> 1 And InStr(1, CheminFichier, "http://", vbTextCompare) <> 1) Then
        ChDir CheminFichier     ' Le chemin du fichier ne contient pas de lien, on change le dossier d'ouverture
    End If

' V1.7 : optimisation : retrait de la mise � jour de l'affichage (�a acc�l�re sacr�ment le traitement)
    Application.ScreenUpdating = False
    
    FichierMYP = Application.GetOpenFilename("Fichiers trajets Peugeot App (*.myp),*.myp,Fichiers trajets Citroen App (*.myc),*.myc,Fichiers trajets DS App (*.myd),*.myd")  ' On demande la s�lection du fichier
    If FichierMYP = False Then
        MsgBox "Aucun fichier n'a �t� selectionn� !", vbCritical
        ws_acc.Activate
        ws_acc.Range(Cells(1, 1), Cells(1, 1)).Select  ' On s�lection la cellule A1 ne pas avoir tout le tableau s�lectionn�
        Exit Sub
    End If
    FormeEncours.Show
    FormeEncours.TexteEnCours = "Chargement du fichier en cours..."
    FormeEncours.Repaint
    
    Set JsonTS = FSO.OpenTextFile(FichierMYP, ForReading)
    jsonText = JsonTS.ReadAll
    JsonTS.Close
    
    FormeEncours.TexteEnCours = "Fichier charg�, parsing des donn�es ..."
    FormeEncours.Repaint
    
    ' Comme le fichier existe, on efface tout
    ' Effacage_Donnees     retir� en 1.6
    
    nbTrip = 0    ' On r�initialise le nombre de trajets
    EcrireValeurFormat cell:=ws.Range(CELL_fichierMYP), val:=FichierMYP, wrap:=False
    Set jsonObject = JsonConverter.ParseJson(jsonText)

' V1.9.1 : D�port des VINS renseign�s dans la feuille d'accueil
' V1.9 : recherche des VIN renseign�s dans la feuille cach�e
'    Nb_VIN_renseignes = 0
'    Sheets("Donn�es-VIN").Activate
'    Set R = [A1].CurrentRegion
'    Infos_VIN = R
'    Nb_VIN_renseignes = UBound(Infos_VIN)

' V1.9.1 : recherche des VIN renseign�s dans la feuille Acceuil
    Sheets("Accueil").Activate
    ' Il faut tester si la premi�re ligne sous l'ent�te "VIN - Description v�hicule" est vide ou pas
    ' S'il n'y a pas de valeur, alors il n'y a aucun VIN de renseign�s,
    ' Sinon on chercher � prendre tous les VIN �crits.
    If Cells(L_entete_ListeVIN + 1, C_entete_ListeVIN).Value = "" Then
        Set R = Range(Cells(L_entete_ListeVIN, C_entete_ListeVIN), Cells(L_entete_ListeVIN, C_entete_ListeVIN + 1))
    Else
        Set R = Range(Cells(L_entete_ListeVIN, C_entete_ListeVIN), Cells(L_entete_ListeVIN, C_entete_ListeVIN + 1).End(xlDown))
    End If
    Infos_VIN = R
    Nb_VIN_renseignes = UBound(Infos_VIN)
        
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
        MsgBox "Aucun VIN trouv� dans votre fichier .myp. Pb de structure ?", vbCritical
        Exit Sub
    ElseIf Nb_VIN >= G_Nb_VIN_Max Then      ' Par d�faut, on g�re 20 VINs diff�rents
        MsgBox "Trop de VIN d�tect�s. Pb de structure ?", vbCritical
        Exit Sub
    End If
' Affichage de la forme de choix des VIN � importer, seulement si plus de 1 VIN. Sinon, celui-ci devient d�faut
    If Nb_VIN > 1 Then
        ' Remise � z�ro de la liste des choix
        FormeVIN.FormeVIN_ListeVIN.Clear
        ' Ajout des VIN trouv�s
        For l_Tab_Vin = 1 To Nb_VIN
            ' V1.9 - Recherche de l'info v�hicule associ�e au VIN
            Info_Voiture = ""
            For i = 2 To Nb_VIN_renseignes
                If Infos_VIN(i, 1) = Liste_VIN(l_Tab_Vin, 1) Then
                    Info_Voiture = Infos_VIN(i, 2)
                    Exit For
                End If
            Next i
            If Info_Voiture = "" Then    ' VIN n'existant pas, on demande les infos pour ajouter ses infos
                Load FormeInfoVIN
                FormeInfoVIN.NumVIN = "VIN : " & Liste_VIN(l_Tab_Vin, 1)
                FormeInfoVIN.DescriptionVIN = "Ma voiture " & l_Tab_Vin
                FormeInfoVIN.Show
                Info_Voiture = FormeInfoVIN.DescriptionVIN
                FormeInfoVIN.Hide
                Dim toto, tata
                
                ' Il faut tester si la premi�re ligne sous l'ent�te "VIN - Description v�hicule" est vide ou pas
                ' S'il n'y a pas de valeur, alors il n'y a aucun VIN de renseign�s,
                ' Sinon on chercher � prendre tous les VIN �crits.
                If Cells(L_entete_ListeVIN + 1, C_entete_ListeVIN).Value = "" Then
                    Cells(L_entete_ListeVIN + 1, C_entete_ListeVIN) = Liste_VIN(l_Tab_Vin, 1)
                    Cells(L_entete_ListeVIN + 1, C_entete_ListeVIN + 1) = Info_Voiture
                Else
                    Cells(L_entete_ListeVIN, C_entete_ListeVIN).End(xlDown).Offset(1, 0) = Liste_VIN(l_Tab_Vin, 1)
                    Cells(L_entete_ListeVIN, C_entete_ListeVIN + 1).End(xlDown).Offset(1, 0) = Info_Voiture
                End If
                'Cells(Range("M3").End(xlDown).Row + 1, 2) = Info_Voiture
                'Cells(Range("M3").End(xlDown).Row + 1, 1) = Liste_VIN(l_Tab_Vin, 1)
                Nb_VIN_renseignes = Nb_VIN_renseignes + 1
                ReDim Infos_VIN(Nb_VIN_renseignes, 2)
                Infos_VIN(Nb_VIN_renseignes, 1) = Liste_VIN(l_Tab_Vin, 1)
                Infos_VIN(Nb_VIN_renseignes, 2) = Info_Voiture
            End If
            FormeVIN.FormeVIN_ListeVIN.AddItem (Liste_VIN(l_Tab_Vin, 1) & " - " & Info_Voiture)
        Next l_Tab_Vin
        ' activation de la forme de choix des VIN
        FormeVIN.Show
        ' Si bouton annuler = on quitte la proc�dure
        If FormeVIN.BoutonChoisi.Value = 2 Then
            MsgBox "Vous avez annul�. On quitte !", vbCritical
            FormeEncours.Hide
            Exit Sub
        End If
        ' On parcourt la liste pour r�cup�rer les VIN s�lectionn�s
        For l_Tab_Vin = 0 To FormeVIN.FormeVIN_ListeVIN.ListCount - 1
            If FormeVIN.FormeVIN_ListeVIN.Selected(l_Tab_Vin) Then
                Liste_VIN(l_Tab_Vin + 1, 2) = True
            End If
        Next
    Else  ' cas 1 seul VIN pr�sent dans le fichier
        Liste_VIN(1, 2) = True
        ' V1.9 - Recherche de l'info v�hicule associ�e au VIN
        Info_Voiture = ""
        For i = 2 To Nb_VIN_renseignes
            If Infos_VIN(i, 1) = Liste_VIN(1, 1) Then
                Info_Voiture = Infos_VIN(i, 2)
                Exit For
            End If
        Next i
        If Info_Voiture = "" Then    ' VIN n'existant pas, on demande les infos pour ajouter ses infos
            Load FormeInfoVIN
            FormeInfoVIN.NumVIN = "VIN : " & Liste_VIN(1, 1)
            FormeInfoVIN.DescriptionVIN = "Ma voiture"
            FormeInfoVIN.Show
            Info_Voiture = FormeInfoVIN.DescriptionVIN
            
            ' Il faut tester si la premi�re ligne sous l'ent�te "VIN - Description v�hicule" est vide ou pas
            ' S'il n'y a pas de valeur, alors il n'y a aucun VIN de renseign�s,
            ' Sinon on chercher � prendre tous les VIN �crits.
            If Cells(L_entete_ListeVIN + 1, C_entete_ListeVIN).Value = "" Then
                Cells(L_entete_ListeVIN + 1, C_entete_ListeVIN) = Liste_VIN(1, 1)
                Cells(L_entete_ListeVIN + 1, C_entete_ListeVIN + 1) = Info_Voiture
            Else
                Cells(L_entete_ListeVIN, C_entete_ListeVIN).End(xlDown).Offset(1, 0) = Liste_VIN(1, 1)
                Cells(L_entete_ListeVIN, C_entete_ListeVIN + 1).End(xlDown).Offset(1, 0) = Info_Voiture
            End If
            ' Cells(Range("A65536").End(xlUp).Row + 1, 2) = Info_Voiture
            ' Cells(Range("A65536").End(xlUp).Row + 1, 1) = Liste_VIN(1, 1)
            
            Nb_VIN_renseignes = Nb_VIN_renseignes + 1
            ReDim Infos_VIN(Nb_VIN_renseignes, 2)
            Infos_VIN(Nb_VIN_renseignes, 1) = Liste_VIN(1, 1)
            Infos_VIN(Nb_VIN_renseignes, 2) = Info_Voiture
            FormeInfoVIN.Hide
        End If
    End If

    Sheets("Trajets-MyPeugeot").Activate
' V1.8
    FormeEncours.TexteEnCours = "Traitement des donn�es en cours, patience..."
    FormeEncours.Repaint
    
' V1.6 : stockage dans un tableau interne (que l'on vide d'abord) de tous les trajets d�j� dans l'Excel
    For i = 1 To G_Nb_Trajets_Max
        T_Trajets_existants(i) = ""
    Next i
    ' Pour vraiment avoir la derni�re ligne du tableau rempli sans risquer d'effacer une donn�e, il faut r�initialiser les filtres
    ' Sinon la derni�re ligne remplie ne sera que celle affich�e, toutes celles masqu�es au-dessous ne compteront pas...
    Worksheets("Trajets-MyPeugeot").Range("$A$3:$B$150000").AutoFilter Field:=1
    Derniere_Ligne_remplie = ws.Cells(Columns(1).Cells.Count, 1).End(xlUp).Row
    Nb_Trajets = 0
    For i = L_Premiere_Valeur To Derniere_Ligne_remplie
        T_Trajets_existants(i - L_Premiere_Valeur + 1) = ws.Cells(i, C_vin) & ";" & ws.Cells(i, C_id)
        Nb_Trajets = Nb_Trajets + 1
    Next i

    i = L_Premiere_Valeur + Nb_Trajets   ' On d�fini un compteur qui sert � se positionner sur la ligne o� les donn�es doivent �tre �crites.
                        ' Le n� de la ligne o� on va commencer � �crire les donn�es
                        ' ws.Cells(LIGNE, COLONNE)     O� LIGNE commence � 1
                        '                              O� COLONNE commence � 1 = A
                        ' On va utiliser les constantes d�finie avant la fonction pour �crire/formater/effacer une cellule
                        
    For Each item In jsonObject
' V1.5 : 1�re v�rif � faire : que le VIN du trajet soit dans les VIN � importer
        VIN_Actuel = item("vin")
        VIN_A_Traiter = False
        For l_Tab_Vin = 1 To Nb_VIN
            If Liste_VIN(l_Tab_Vin, 1) = VIN_Actuel Then
                VIN_A_Traiter = Liste_VIN(l_Tab_Vin, 2)
            End If
        Next l_Tab_Vin
      
' V1.5 : on ne traite les trajets QUE SI Vin_A_Traiter est vrai
        If VIN_A_Traiter Then
            For Each item_item In item("trips")   ' Boucle pour r�cup�rer les trajets
                nbTrip = nbTrip + 1
' V1.8
                If ((nbTrip Mod 100) = 0) Then
                    FormeEncours.TexteEnCours = "Traitement trajets, " & nbTrip & " analys�s, patience..."
                    FormeEncours.Repaint
                End If

' V1.6 : � ne faire que si VIN et Id n'�taient pas d�j� pr�sents
                Trajet_Trouve = False
                For j = 1 To Nb_Trajets
                    If T_Trajets_existants(j) = VIN_Actuel & ";" & item_item("id") Then
                        Trajet_Trouve = True
                        Exit For
                    End If
                Next j
                If Not Trajet_Trouve Then
                    Cells(i, C_vin).Value = VIN_Actuel      ' On �crit le VIN r�cup�r�
                    Cells(i, C_id).Value = item_item("id")  ' On �crit l'ID r�cup�r�
                    
                    ' R�cup�ration des dates
                    ' On stocke les deux dates (d�part et arriv�e) car il faut d�terminer le temps de parcours
                    ' qui ne doit pas �tre d�pendant d'un �ventuelle changement d'heure en cours de route
                    MaDate_UNIX_GMT_dep = item_item("startDateTime")            ' Date de d�part
                    MaDate_UNIX_GMT_arr = item_item("endDateTime")              ' Date d'arriv�e
                    MaDate_DST_dep = Date_UNIX_To_Date_DST(MaDate_UNIX_GMT_dep)
                    MaDate_DST_arr = Date_UNIX_To_Date_DST(MaDate_UNIX_GMT_arr)
                    
                    Cells(i, C_date_dep).Value = MaDate_DST_dep
                    Cells(i, C_date_arr).Value = MaDate_DST_arr
                    
                    ' Calcul de la dur�e du trajet en cours
                    duree_trajet = MaDate_UNIX_GMT_arr - MaDate_UNIX_GMT_dep
                    duree_trajet_bis = Date_UNIX_To_Date_GMT(MaDate_UNIX_GMT_arr) - Date_UNIX_To_Date_GMT(MaDate_UNIX_GMT_dep)
                    ' V1.9 : en cas de souci entre date d�part et date arriv�e, et donc si duree_trajet est < 0, on met 1 seconde par d�faut
                    If duree_trajet < 0 Then
                        duree_trajet = 1
                        duree_trajet_bis = "00:00:01"
                    End If
                    Cells(i, C_duree).Value = duree_trajet_bis
                    
                    distance_trajet = item_item("endMileage") - item_item("startMileage")
                    Cells(i, C_dist).Value = distance_trajet
                    Cells(i, C_dist_tot).Value = item_item("endMileage")
                    
                    Cells(i, C_conso).Value = item_item("consumption")
                    conso_trajet = Cells(i, C_conso).Value
                    ' Pour le calcul de la consommation moyenne, il faut �viter la division par z�ro dans le cas o�
                    ' la voiture � tourner � l'arret, la distance parcourue est nulle
                    If distance_trajet <> 0 Then
                        Cells(i, C_conso_moy).Value = conso_trajet / distance_trajet * 100
                    Else
                        Cells(i, C_conso_moy).Value = "//"
                    End If
        
                    Cells(i, C_pos_dep_lat).Value = item_item("startPosLatitude")
                    Cells(i, C_pos_dep_long).Value = item_item("startPosLongitude")
                                
                    adresse_dep = item_item("startPosAddress")
                    Cells(i, C_pos_dep_adr).Value = Correction_Adresse(adresse_dep)
                    Cells(i, C_pos_arr_lat).Value = item_item("endPosLatitude")
                    Cells(i, C_pos_arr_long).Value = item_item("endPosLongitude")
                    
                    adresse_arr = item_item("endPosAddress")
                    Cells(i, C_pos_arr_adr).Value = Correction_Adresse(adresse_arr)
                    
                    Cells(i, C_niv_carb).Value = item_item("fuelLevel") / 100
                    Cells(i, C_auto).Value = item_item("fuelAutonomy")
                    
                    i = i + 1
                    ' Il n'est plus utilse de r�cup�rer ces valeurs ici puisqu'elles sont r�cup�rer plus bas en utilisant les range()
                    'kilometrage = item_item("endMileage")   ' On stocke le kilom�trage de fin de trajet pour �tre affich� en tant que kilom�trage actuel lors du dernier trajet
' V1.6 : fin du test
                End If
            Next
    
' V1.5 : fin du test
        End If
    Next
' V1.8
    FormeEncours.TexteEnCours = "Tri des trajets en cours"
    FormeEncours.Repaint
    
' V1.6 : tri final sur colonnes A puis B
    ws.Range(CELL_plage_1ereligne).Select
    ws.Range(Selection, Selection.End(xlDown)).Select
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=Range(CELL_plage_max_COL_vin), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal     ' Range de base "A5:A65536"
    ws.Sort.SortFields.Add Key:=Range(CELL_plage_max_COL_id), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal      ' Range de base "B5:B65536"
    With ws.Sort
        .SetRange Range(CELL_plage_max_Donnees)        ' range par d�faut "A5:Q65536"
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ws.Cells(1, 4).Select   ' On s�lection la cellule de version pour ne pas avoir tout le tableau s�lectionn�
        
' V1.7 : mise en place du formatage par colonne
' Colonne VIN
    Formater_Cellules WS_tmp:=ws, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_vin, f_size:=8
    'ws.Range(ws.Cells(L_Premiere_Valeur, C_vin), ws.Cells(L_Premiere_Valeur, C_vin).End(xlDown)).Font.Size = 8
     
    'Formater_Cellules cell:=ws.Cells(L_Premiere_Valeur, C_vin)
' Colonne Date d�part
    Formater_Cellules WS_tmp:=ws, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_date_dep, n_format:="date"
    'Range(Cells(L_Premiere_Valeur, C_date_dep), Cells(L_Premiere_Valeur, C_date_dep)).Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormat = "dd/mm/yy - hh:mm"
' Colonne Date Arriv�e
    Formater_Cellules WS_tmp:=ws, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_date_arr, n_format:="date"
    'Range(Cells(L_Premiere_Valeur, C_date_arr), Cells(L_Premiere_Valeur, C_date_arr)).Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormat = "dd/mm/yy - hh:mm"
' Colonne Dur�e
    Formater_Cellules WS_tmp:=ws, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_duree, n_format:="duree"
    'Range(Cells(L_Premiere_Valeur, C_duree), Cells(L_Premiere_Valeur, C_duree)).Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormat = "h:mm"
' Colonne Distance
    Formater_Cellules WS_tmp:=ws, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_dist, n_format:="1"
    'Range(Cells(L_Premiere_Valeur, C_dist), Cells(L_Premiere_Valeur, C_dist)).Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormatLocal = "0,0"
' Colonne Distance totale
    Formater_Cellules WS_tmp:=ws, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_dist_tot, n_format:="0"
    'Range(Cells(L_Premiere_Valeur, C_dist_tot), Cells(L_Premiere_Valeur, C_dist_tot)).Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormatLocal = "0"
' Colonne conso
    Formater_Cellules WS_tmp:=ws, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_conso, n_format:="2"
    'Range(Cells(L_Premiere_Valeur, C_conso), Cells(L_Premiere_Valeur, C_conso)).Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormatLocal = "0,00"
' Colonne conso moyenne
    Formater_Cellules WS_tmp:=ws, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_conso_moy, n_format:="1"
    'Range(Cells(L_Premiere_Valeur, C_conso_moy), Cells(L_Premiere_Valeur, C_conso_moy)).Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormatLocal = "0,0"
' Colonne niveau carburant
    Formater_Cellules WS_tmp:=ws, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_niv_carb, n_format:="%"
    'Range(Cells(L_Premiere_Valeur, C_niv_carb), Cells(L_Premiere_Valeur, C_niv_carb)).Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormat = "0 %"
' Colonne adresse D�part
    Formater_Cellules WS_tmp:=ws, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_pos_dep_adr, n_format:="add"
' Colonne adresse Arriv�e
    Formater_Cellules WS_tmp:=ws, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_pos_arr_adr, n_format:="add"
    
   ' On recalcule le nombre de trajets pr�sent dans le tableau : ce nombre tient compte de tous les trajets affich�s, ceux initialement pr�sent avant l'import + ceux import�s.
   ' On compte combien il y a d'ID de trajet
' v1.8    nbTrip = ws.Range(ws.Cells(L_Premiere_Valeur, C_id), ws.Cells(L_Premiere_Valeur, C_id).End(xlDown)).Count
' v1.8    EcrireValeurFormat cell:=ws.Range(CELL_nb_trips), val:=nbTrip, f_size:=12                   ' On �crit le nombre de trajet
    
' v1.8    derniere_val_tab = 4 + nbTrip      ' valeur qui d�limite la derni�re ligne contenant un trajet
    
    ' On r�cup�re le kilom�trage final du dernier trajer vu qu'il n'est plus forc�ment r�cup�r� si on charge un fichier de donn�e o� il n'y a pas de nouveaux trajets
' v1.8    kilometrage = ws.Cells(derniere_val_tab, C_dist_tot).Value
' v1.8    EcrireValeurFormat cell:=ws.Range(CELL_km), val:=kilometrage, n_format:="0", f_size:=12     ' On �crit le kilom�trage total de la voiture
    
    ' Calcul de la consommation totale du tableau
' v1.8    conso_totale = ws.Application.WorksheetFunction.Sum(ws.Range(CELL_plage_conso_tot & derniere_val_tab))
    ' On �crit la valeur calcul�e de la consommation totale de tous les trajets
' v1.8    EcrireValeurFormat cell:=ws.Range(CELL_conso_tot), val:=conso_totale, n_format:="2", f_size:=12
    
' v1.8    ws.Range(CELL_conso_tot_moy).FormulaLocal = FORMULE_calcul_Conso_tot_moy        ' On r��crit la formule au cas-o�... si jamais les cellules sont d�plac�es :)
' v1.8    ws.Range(CELL_conso_tot_moy).NumberFormatLocal = "0,0"
    
' V1.8 : rafraichissement du TCD. On positionne tous les VIN � "coch�" sauf le VIN "vide" par d�faut
    Sheets("Accueil").PivotTables("TCD_VIN").PivotCache.Refresh
    Sheets("Accueil").PivotTables("TCD_VIN").PivotFields("VIN(s)").CurrentPage = "(All)"
' V1.9 : s�lection du 1er vin import� dans le TCD
    With Sheets("Accueil").PivotTables("TCD_VIN").PivotFields("VIN(s)")
        .ClearAllFilters
        For Each i_TCD In .PivotItems
            If i_TCD = Liste_VIN(1, 1) Then
                .CurrentPage = Liste_VIN(1, 1)
                Exit For
            Else
                .CurrentPage = "(blank)"
            End If
        Next
    End With
' V1.8
    FormeEncours.Hide
    Sheets("Trajets-MyPeugeot").Activate
' V1.9 : s�lection du 1er vin import� dans le filtre
    ActiveSheet.Range("$A$2:$B$655007").AutoFilter Field:=1, Criteria1:=Liste_VIN(1, 1)
' V1.7 : optimisation : remise de la mise � jour de l'affichage
    Application.ScreenUpdating = True
   
    ws_acc.Activate
    ws_acc.Range(Cells(1, 1), Cells(1, 1)).Select  ' On s�lection la cellule A1 ne pas avoir tout le tableau s�lectionn�
    
End Sub

Private Sub Formater_Cellules(WS_tmp As Worksheet, ligne_cell As Variant, colonne_cell As Variant, Optional n_format As String = "General", Optional f_size As Integer = 10, Optional wrap As Boolean = False)
    ' Fonction pour �crire formater la valeur dans une cellule ou une plage de cellules
    ' Arguments obligatoires :  WS_tmp As Worksheet     <- La feuille de calcul o� on travail
    '                           ligne_cell As Variant   <- La ligne de la cellule de d�part
    '                           colonne_cell As Variant <- La colonne de la cellule de d�part
    ' Arguments optionels :     n_format As String = "General"  <- Le format NumberFormat, d�faut = "Genral"
    '                                                           <- Valeurs = "date" pour format date
    '                                                           <- Valeurs = "duree" pour format dur�e
    '                                                           <- Valeurs = "0" pour format num�rique Local sans virgule
    '                                                           <- Valeurs = "1" pour format num�rique Local avec 1 chiffre apr�s la virgule
    '                                                           <- Valeurs = "2" pour format num�rique Local avec 2 chiffres apr�s la virgule
    '                                                           <- Valeurs = "add" pour les adresses
    '                           font_size As Integer = 10       <- La taille de caract�re, d�faut = 10
    '                           wrap As Boolean = False         <- Retour � la ligne dans la cellule, d�faut = faux
       
    Select Case n_format
        Case "date"
            WS_tmp.Range(WS_tmp.Cells(ligne_cell, colonne_cell), WS_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).NumberFormat = "dd/mm/yy - hh:mm"
        Case "duree"
            WS_tmp.Range(WS_tmp.Cells(ligne_cell, colonne_cell), WS_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).NumberFormat = "h:mm"
        Case "0"
            WS_tmp.Range(WS_tmp.Cells(ligne_cell, colonne_cell), WS_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).NumberFormatLocal = "0"
        Case "1"
            WS_tmp.Range(WS_tmp.Cells(ligne_cell, colonne_cell), WS_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).NumberFormatLocal = "0,0"
        Case "2"
            WS_tmp.Range(WS_tmp.Cells(ligne_cell, colonne_cell), WS_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).NumberFormatLocal = "0,00"
        Case "%"
            WS_tmp.Range(WS_tmp.Cells(ligne_cell, colonne_cell), WS_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).NumberFormat = "0 %"
        Case Else
            WS_tmp.Range(WS_tmp.Cells(ligne_cell, colonne_cell), WS_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).NumberFormat = "General"
    End Select
    
    If n_format = "add" Then        ' Il faut v�rifier si on est sur un champ adresse car pour l'adresse il faut aligner � gauche xlHAlignLeft
        WS_tmp.Range(WS_tmp.Cells(ligne_cell, colonne_cell), WS_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).HorizontalAlignment = xlHAlignLeft
    Else
        WS_tmp.Range(WS_tmp.Cells(ligne_cell, colonne_cell), WS_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).HorizontalAlignment = xlHAlignCenter
    End If
    ' Dans tous les cas, le VerticalAlignment est � xlVAlignCenter
    WS_tmp.Range(WS_tmp.Cells(ligne_cell, colonne_cell), WS_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).VerticalAlignment = xlVAlignCenter
    
'    If (f_size <> 10) Then
'        WS_tmp.Range(WS_tmp.Cells(ligne_cell, colonne_cell), WS_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).Font.Size = f_size
'    Else
'        WS_tmp.Range(WS_tmp.Cells(ligne_cell, colonne_cell), WS_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).Font.Size = 10
'    End If
    WS_tmp.Range(WS_tmp.Cells(ligne_cell, colonne_cell), WS_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).Font.Size = f_size
    
    If (wrap = True) Then
        WS_tmp.Range(WS_tmp.Cells(ligne_cell, colonne_cell), WS_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).WrapText = True
    Else
        WS_tmp.Range(WS_tmp.Cells(ligne_cell, colonne_cell), WS_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).WrapText = False
    End If
    
End Sub

Private Sub EcrireValeurFormat(cell As Variant, val As Variant, Optional n_format As String = "General", Optional f_size As Integer = 10, Optional wrap As Boolean = False)
    ' Fonction pour �crire une valeur dans une cellule
    ' Arguments obligatoires :  cellule As Variant  <- La cellule ou plage de cellule devant �tre modifi�e
    '                           val As Variant      <- La valeur � �crire dans la cellule/plage
    ' Arguments optionels :     n_format As String = "General"  <- Le format NumberFormat, d�faut = "Genral"
    '                                                           <- Valeurs = "date" pour format date
    '                                                           <- Valeurs = "duree" pour format dur�e
    '                                                           <- Valeurs = "0" pour format num�rique Local sans virgule
    '                                                           <- Valeurs = "1" pour format num�rique Local avec 1 chiffre apr�s la virgule
    '                                                           <- Valeurs = "2" pour format num�rique Local avec 2 chiffres apr�s la virgule
    '                           font_size As Integer = 10       <- La taille de caract�re, d�faut = 10
    '                           wrap As Boolean = False         <- Retour � la ligne dans la cellule, d�faut = faux
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
                .NumberFormat = "General"
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
    ' Fonction pour effacer les donn�es avant d'utiliser celles du fichier ouvert
' 1.9 : il faut d'abord retirer le filtre sur le VIN pour TOUT effacer
    Worksheets("Trajets-MyPeugeot").Range("$A$3:$B$150000").AutoFilter Field:=1
    With Worksheets("Trajets-MyPeugeot")
' 1.8        .Range(CELL_vin_entete) = ""           ' Le VIN
        .Range(CELL_fichierMYP) = ""           ' Le nom du fichier
' 1.8       .Range(CELL_nb_trips) = ""  ' Le nombre de trajet et le kilom�trage total
' 1.8       .Range(CELL_km) = ""
        .Range(CELL_plage_donnees) = ""    ' Le grand tableau de valeurs de tous les trajets effectu�s
' 1.8       .Range(CELL_conso_tot) = ""           ' La consommation totale sur tous les trajets
    End With
    With Worksheets("Accueil")
        .PivotTables("TCD_VIN").PivotCache.Refresh
        .Range("M4:N65000") = ""
    End With
 
    
    
    With Worksheets("Donnees-Recap")
        .Range("A2:D10000") = ""  ' Un tableau r�capitulatif mensuel
        .Range("F4:F53") = 0      ' Le nombre de trajets regroup� en plage de distance
        .Range("H4:H53") = 0      ' La distance des regroupements de trajets
    End With

End Sub
Private Function Date_UNIX_To_Date_GMT(date_unix_GMT As Long) As Date
    ' Fonction qui converti un temps UNIX en date GMT
    '@PARAM {Long} date � convertir, au format UNIX GMT
    '@RETURN {Date} Renvoi la date convertie en date GMT

    Date_UNIX_To_Date_GMT = (date_unix_GMT / 86400) + DateValue("01/01/1970")
End Function
Private Function Date_UNIX_To_Date_DST(date_unix_GMT As Long) As Date
    ' Fonction qui converti un temps UNIX en date avec DST (changement d'heure)
    '@PARAM {Long} date � convertir, au format UNIX GMT
    '@RETURN {Date} Renvoi la date convertie en date GMT
    Dim DST_val As Integer, date_unix_DST As Long
    ' La date ainsi calcul�e ne tient pas compte du passage � l'heure d'�t� ou � l'heure d'hiver
    ' Il faut v�rifier si le jour du mois de cette date est avant ou apr�s le dernier dimanche de mars ou d'octbre.
    ' Rappel :  l'heure passe en GMT +2 au dernier dimanche de mars, c'est l'heure d'�t�
    '           l'heure passe en GMT +1 au dernier dimanche d'octobre, c'est l'heure d'hiver
    ' On va donc devoir ajouter soit 1h=3600s au temps GMT si c'est l'heure d'hiver, soit 2h=7200s au temps GMT si c'est l'heure d'hiver
    ' D�terminons la valeur du facteur DST
    DST_val = DST(date_unix_GMT)
    
    ' Calcul de la nouvelle date avec DST
    date_unix_DST = date_unix_GMT + DST_val * 3600
    ' Conversion en Date de cette date UNIX : la nouvelle Date est maintenant DST
    Date_UNIX_To_Date_DST = Date_UNIX_To_Date_GMT(date_unix_DST)

End Function

Private Function DST(date_unix_GMT As Long) As Integer
    ' Fonction qui d�termine le modificateur d'heure (Day Saving Time) � appliquer � l'heure GMT pour avoir l'heure FR
    ' Ce sera le modificateur d'heure par rapport au temps unix GMT :   1 = pour l'heure d'hiver
    '                                                                   2 = pour l'heure d'�t�
    '@PARAM {Long} date � tester, au format unix GMT
    '@RETURN {Integer} Renvoi un entier permettant la modification de l'heure unix


    ' On d�clare les variables utilis�es pour le jour, le mois, l'ann�e, l'heure en temps GMT
    Dim jour_GMT As Integer, mois_GMT As Integer, annee_GMT As Integer, heure_GMT As Integer, minutes_GMT As Integer
    Dim date_temp As Date
    
    ' On d�clare les variables utilis�es pour le dernier dimanche de mars ou d'octobre : jour/mois/heure
    Dim jour_dD As Integer, mois_dD As Integer, heure_dD As Integer
    Dim num_jour_31 As Integer      ' C'est pour stocker le n� du jour de semaine pour le 31/03/annee ou 31/10/annee
   
    ' On convertir la date unix en date GMT
    date_temp = Date_UNIX_To_Date_GMT(date_unix_GMT)
    
    ' On r�cup�re le jour, le mois, l'ann�e, l'heure en temps GMT
    jour_GMT = Day(date_temp)
    mois_GMT = Month(date_temp)
    annee_GMT = Year(date_temp)
    heure_GMT = Hour(date_temp)
    ' minutes = Minute(date_temp)       ' Inutile ici
     
    Select Case mois_GMT
        Case 1, 2, 11, 12
            ' On est dans le cas o� l'heure d'hiver est appliqu�e : de Novembre � F�vrier
            DST = 1
            
        Case 4 To 9
            ' On est dans le cas o� l'heure d'�t� est appliqu�e : de Avril � Septembre
            DST = 2
            
        Case 3
            ' On est en mars, il faut v�rifier que le jour en question est avant ou apr�s le dernier dimanche de mars
            ' D�termination du n� du jour de la semaine du dernier jour de mars
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
                ' On est apr�s le dernier dimanche du mois, donc en heure d'�t�
                DST = 2
            ElseIf jour_GMT = jour_dD Then
                ' On est le dernier dimanche du mois, donc il faut voir si on est avant 1h du matin ou pas.
                If heure_GMT < 1 Then
                    ' On est encore � l'heure d'hiver
                    DST = 1
                Else
                    ' On est pass� � l'heure d'�t�
                    DST = 2
                End If
            End If
            
        Case 10
            ' On est en octobre, il faut v�rifier que le jour en question est avant ou apr�s le dernier dimanche de mars
            ' D�termination du n� du jour de la semaine du dernier jour d'octobre
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
                ' On est apr�s le dernier dimanche du mois, donc en heure d'�t�
                DST = 2
            ElseIf jour_GMT = jour_dD Then
                ' On est le dernier dimanche du mois, donc il faut voir si on est avant 1h du matin ou pas.
                If heure_GMT < 1 Then
                    ' On est encore � l'heure d'�t�
                    DST = 2
                Else
                    ' On est pass� � l'heure d'hiver
                    DST = 1
                End If
            End If
        End Select
        
End Function

Private Function Correction_Adresse(ByVal adresse As String) As String
    Dim i As Integer
    Dim lettre As String * 2

    adresse = Replace(adresse, vbLf, ", ")   ' On remplace tous les retours � la ligne "\n" par des ", "
    For i = 1 To Len(adresse) - 1
        lettre = Mid(adresse, i, 2)
        Select Case lettre
            Case "è"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2, 100)
            Case "é"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2, 100)
            Case "ê"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2, 100)
            Case "À"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2, 100)
            Case "â"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2, 100)
            Case "œ"   ' oe
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2, 100)
            Case "à"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2, 100)
            Case "ç"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2, 100)
            Case "ï"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2, 100)
            Case "ù"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2, 100)
            Case "ü"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2, 100)
            Case "É"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2, 100)
            Case "È"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2, 100)
        End Select
    Next i
    Correction_Adresse = adresse
End Function

Sub Bouton11_Cliquer()
    Sheets("Accueil").Activate
End Sub

Sub Bouton3_Cliquer()
    Sheets("Trajets-MyPeugeot").Activate
End Sub

