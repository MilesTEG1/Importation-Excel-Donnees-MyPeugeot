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
'       - V 1.9.2 : Correction de quelques bugs
'       - V 1.9.3 : Correction de quelques bugs + Ajout de la derni�re adresse d'arriv�e connue pour le VIN s�lectionn�
'       - V 1.9.4 : Ajout d'une feuille Tutoriel expliquant les diff�rentes fonctions du fichier XLSM.
'       - V 1.9.5 : Ajout d'une d�cimale sur les kilom�tres totaux affich�s (colonne G)
'       - V 2.0 :   Ajout des valeurs non utilis�e dans des colonnes masqu�es en vue de l'exportation des donn�es en format fichier JSON
'                   Mise en place d'une structure de type Dictionnaire pour stocker les colonnes utilis�es pour les donn�es
'                   Ajout de la marque de la voiture (d�tect�e automatiquement avec l'extension du fichier de donn�es fourni)
'                   Ajout de 3 fonctions pour d�terminer un temps UNIX UTC � partir d'une date DST (pour la reconstruction d'un trajet manquant)
'                   Ajout de nouveaux caract�res accentu�s pour la conversion d'adresses
'                   Ajout d'une fonction de correction inverse des adresses en vue de la re-cr�ation d'un fichier de donn�es � partir du tableau excel (afin de tenir compte des �ventuelles modifications d'adresses par les utilisateurs).
'
' Couples de versions d'Excel & OS test�es :
'       - Windows 10 v1909 (18363.752) & Excel pour Office 365 Version 2003 (build 12624.20382 & 12624.20442)
'       - Windows 10 v1809 & Excel 2016
'       - Windows 10 v1909 & Excel 2019
'
Option Explicit     ' On force � d�clarer les variables
'
' D�inissons quelques constantes qui serviront pour les colonnes/lignes/plages de cellules.
'
'
Const VERSION As String = "v2.0 Beta 6"     ' Version du fchier
Const CELL_ver As String = "B3"             ' Cellule o� afficher la version du fichier

' Constantes pour la feuille "Trajets"
Const L_Premiere_Valeur As Integer = 3      ' Premi�re ligne � contenir des donn�es (avant ce sont les lignes d'en-t�te
Const C_vin             As Integer = 1      ' COLONNE = 1 (A) -> VIN pour le trajet (il est possible d'avoir plusieurs VIN dans le fichier de donn�e).
' Const C_id              As Integer = 2      '         = 2 (B) -> Trip ID#
' Const C_date_dep        As Integer = 3      '         = 3 (C) -> Date de D�part (� d�terminer avec une conversion)
' Const C_date_arr        As Integer = 4      '         = 4 (D) -> Date de Fin    (� d�terminer avec une conversion)
Const C_duree           As Integer = 5      '         = 5 (E) -> Dur�e du trajet  (� calculer)
Const C_dist            As Integer = 6      '         = 6 (F) -> Distance du trajet en km
' Const C_dist_tot        As Integer = 7      '         = 7 (G) -> Distance totale au compteur km
' Const C_conso           As Integer = 8      '         = 8 (H) -> Consommation du tajet en L
Const C_conso_moy       As Integer = 9      '         = 9 (I) -> Consommation moyenne en L/100km
' Const C_pos_dep_lat     As Integer = 10     '         = 10 (J) -> Position de d�part - Latitude
' Const C_pos_dep_long    As Integer = 11     '         = 11 (K) -> Position de d�part - Longitude
' Const C_pos_dep_adr     As Integer = 12     '         = 12 (L) -> Position de d�part - Adresse
' Const C_niv_carb        As Integer = 16     '         = 16 (P) -> Niveau de carburant en %
' Const C_auto            As Integer = 17     '         = 17 (Q) -> Autonomie restante en km
Const CELL_fichierMYP   As String = "G" & (L_Premiere_Valeur - 2)  ' Cellule qui contiendra le nom du fichier import�
Const CELL_plage_donnees    As String = "A" & L_Premiere_Valeur & ":AJ65536" ' Plage de cellules contenant les donn�es
Const CELL_plage_1ereligne  As String = "A" & L_Premiere_Valeur & ":AJ" & L_Premiere_Valeur  ' Toute la 1ere ligne de donn�e pour en faire le tri
Const CELL_plage_max_Donnees As String = "A" & L_Premiere_Valeur & ":AJ65536"    ' La plage de donn�es maximale possible
Const CELL_plage_max_COL_vin As String = "A" & L_Premiere_Valeur & ":A65536"    ' La plage de donn�es maximale possible
Const CELL_plage_max_COL_id As String = "B" & L_Premiere_Valeur & ":B65536"    ' La plage de donn�es maximale possible
' V1.5 : MultiVIN
Const G_Nb_Trajets_Max = 20000              ' Nb trajets max par VIN trait�s par cette macro

' Constantes pour la feuille Accueil
Const C_entete_ListeVIN        As Integer = 13      ' = 13 (M) Colonne d'ent�te des VIN dans la liste des vins r�cup�r�s
                                                    ' la colonne des descriptions des v�hicules est celle d'� cot� : 13+1 = N
                                                    ' le colonne de la marque des v�hicule est celle d'� cot� encore : 13+2 = O
Const L_entete_ListeVIN        As Integer = 3      ' Ligne d'ent�te des VIN dans la liste des vins r�cup�r�s, elle correspond aussi � celles des descriptions des v�hicules

'
' Variables g�n�rales
'Public Const C_pos_arr_lat     As Integer = 13     '         = 13 (M) -> Position d'arriv�e - Latitude
'Public Const C_pos_arr_long    As Integer = 14     '         = 14 (N) -> Position d'arriv�e - Longitude
'Public Const C_pos_arr_adr     As Integer = 15     '         = 15 (O) -> Position d'arriv�e - Adresse
Public Const CELL_lat  As String = "H9"
Public Const CELL_long As String = "H8"
Public Const CELL_adr As String = "E8"
Public Const G_Nb_VIN_Max = 20                      ' Nb VIN max trait� par cette macro
Public Macro_en_cours, DicoRempli As Boolean
Public DicoJSON As New Scripting.Dictionary         ' D�crit ton objet, � quoi il sert...
'
' Fin de d�claration des constantes
'

'
' Fonction pour lire les donn�es depuis un fichier JSON et les �crire dans le tableau
'
Sub MYP_JSON_Decode()
    Dim jsonText As String
    Dim jsonObject As Object, item As Object, item_item As Object, item_error As Object, item_item_error As Object
    Dim ws_Trajet As Worksheet, ws_Accueil As Worksheet
    
    Set ws_Trajet = Worksheets("Trajets")
    Set ws_Accueil = Worksheets("Accueil")
    
    Dim i, j As Long     ' Variables utilis�es pour les compteurs de boucles For
    Dim Trouve As Boolean
        
    Dim FSO As New FileSystemObject
    Dim JsonTS As TextStream
    
    Dim FichierMYP As Variant   ' nom du fichier myp
    Dim nbTrip, nbTripReel As Integer ' nombre de trajets
    Dim kilometrage As Single   ' Pour stocker le kilom�trage total de la voiture.
    Dim conso_totale As Single  ' Pour stocker la consommation totale de tout le tableau
    Dim CheminFichier As String
    Dim MaDate_UNIX_UTC_dep As Long, MaDate_DST_dep As Date  ' Pour convertir la date unix de d�part en date excel
    Dim MaDate_UNIX_UTC_arr As Long, MaDate_DST_arr As Date  ' Pour convertir la date unix d'arriv�e en date excel
    Dim duree_trajet As Long, duree_trajet_bis As Date
    Dim distance_trajet As Single, conso_trajet As Single, niveau_carb As Single
    Dim adresse_dep As String, adresse_arr As String, adresse_tmp As String
    Dim derniere_val_tab    ' Derni�re ligne qui contiendra un trajet
        
' V1.5 : pour MultiVIN
    Dim l_Tab_Vin As Integer                    ' pour boucle sur les vin
    Dim Nb_VIN As Integer                       ' Nb VIN trouv�s dans le fichier .myp
    Dim Liste_VIN(G_Nb_VIN_Max, 2) As Variant   ' Tableau interne des VIN trouv�s (col 1 = vin, col 2 = retenu)
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
    Dim Nb_VIN_renseignes, Nb_VIN_Selection, Nb_VIN_Filtre As Integer
    Dim i_TCD, Derniere_Lat, Derniere_Long, Derniere_Adr, Critere(G_Nb_VIN_Max)
' V2.0 : Export JSON => passage des colonnes en dictionaire
    ' Variables pour r�cup�rer les tableaux de code d'erreur pour chaque trajet
    Dim chaine_tmp As String
    Dim nb_item_erreur As Integer, var_tmp_int As Integer
    ' Variable indiquant si le dictionnaire est rempli
    If Not DicoRempli Then
        Call RemplisDicoJSON
    End If
    ' Variable pour la marque de la voiture
    Dim Marque_Voiture As String   ' 3 valeurs possibles : Peugeot si extension de fichier = .myp
                                    '                       Citro�n si extension de fichier = .myc
                                    '                       DS      si extension de fichier = .myd
    
    ' On commence par �crire la valeur de la version du fichier :D
    ws_Accueil.Activate
    EcrireValeurFormat cell:=Range(CELL_ver).Offset(-1, 0), val:="Version du fichier", f_size:=10, wrap:=True
    EcrireValeurFormat cell:=Range(CELL_ver), val:=VERSION, f_size:=16, wrap:=True
    
' V1.8 : activation de cette feuille pour �tre s�r d'�tre dedans
    ws_Trajet.Activate
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
        ws_Accueil.Activate
        Cells(1, 1).Select ' On s�lection la cellule A1 ne pas avoir tout le tableau s�lectionn�
        Exit Sub
    End If
    
        
' V2 : D�termination de la marque du v�hicule
    ' Peugeot si extension de fichier = .myp
    ' Citroen si extension de fichier = .myc
    ' DS      si extension de fichier = .myd
    Select Case Mid(FichierMYP, InStrRev(FichierMYP, ".") + 1)
        Case "myp"  ' Voiture Peugeot
            Marque_Voiture = "Peugeot"
        Case "myc"  ' Voiture Citro�n
            Marque_Voiture = "Citro�n"
        Case "myd"  ' Voiture DS
            Marque_Voiture = "DS"
        Case Else   ' erreur, le fichier n'a pas le bon format...
            ws_Accueil.Activate
            Cells(1, 1).Select ' On s�lection la cellule A1 ne pas avoir tout le tableau s�lectionn�
            MsgBox "Le fichier s�lectionn� n'a pas le bonne extension !" & vbLf & "Merci de changer l'extension du fichier si vous �tes sur de sa validit� pour l'importation des donn�es.", vbCritical
            Exit Sub
    End Select
' Fin d'ajout V2 pour la marque de la voiture
    
    FormeEncours.Show
    FormeEncours.TexteEnCours = "Chargement du fichier en cours..."
    FormeEncours.Repaint
    
    Set JsonTS = FSO.OpenTextFile(FichierMYP, ForReading)
    jsonText = JsonTS.ReadAll
    JsonTS.Close
    
    FormeEncours.TexteEnCours = "Fichier charg�, parsing des donn�es ..."
    FormeEncours.Repaint
   
    nbTrip = 0    ' On r�initialise le nombre de trajets
    EcrireValeurFormat cell:=Range(CELL_fichierMYP), val:=FichierMYP, wrap:=False
    Set jsonObject = JsonConverter.ParseJson(jsonText)


' V1.9.1 : D�port des VINS renseign�s dans la feuille d'accueil
' V1.9.1 : recherche des VIN renseign�s dans la feuille Accueil
    
    ws_Accueil.Activate
    
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
        Application.ScreenUpdating = True
        Exit Sub
    ElseIf Nb_VIN >= G_Nb_VIN_Max Then      ' Par d�faut, on g�re 20 VINs diff�rents
        MsgBox "Trop de VIN d�tect�s. Pb de structure ?", vbCritical
        Application.ScreenUpdating = True
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
                FormeInfoVIN.DescriptionVIN = "Ma voiture " & Nb_VIN_renseignes
                FormeInfoVIN.Show
                Info_Voiture = FormeInfoVIN.DescriptionVIN
                FormeInfoVIN.Hide
                ' Il faut tester si la premi�re ligne sous l'ent�te "VIN - Description v�hicule" est vide ou pas
                ' S'il n'y a pas de valeur, alors il n'y a aucun VIN de renseign�s,
                ' Sinon on chercher � prendre tous les VIN �crits.
                If Cells(L_entete_ListeVIN + 1, C_entete_ListeVIN).Value = "" Then
                    Cells(L_entete_ListeVIN + 1, C_entete_ListeVIN) = Liste_VIN(l_Tab_Vin, 1)
                    Cells(L_entete_ListeVIN + 1, C_entete_ListeVIN + 1) = Info_Voiture
                    Cells(L_entete_ListeVIN + 1, C_entete_ListeVIN + 2) = Marque_Voiture
                Else
                    Cells(L_entete_ListeVIN, C_entete_ListeVIN).End(xlDown).Offset(1, 0) = Liste_VIN(l_Tab_Vin, 1)
                    Cells(L_entete_ListeVIN, C_entete_ListeVIN + 1).End(xlDown).Offset(1, 0) = Info_Voiture
                    Cells(L_entete_ListeVIN, C_entete_ListeVIN + 2).End(xlDown).Offset(1, 0) = Marque_Voiture
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
            FormeInfoVIN.DescriptionVIN = "Ma voiture " & Nb_VIN_renseignes
            FormeInfoVIN.Show
            Info_Voiture = FormeInfoVIN.DescriptionVIN
            
            ' Il faut tester si la premi�re ligne sous l'ent�te "VIN - Description v�hicule" est vide ou pas
            ' S'il n'y a pas de valeur, alors il n'y a aucun VIN de renseign�s,
            ' Sinon on chercher � prendre tous les VIN �crits.
            If Cells(L_entete_ListeVIN + 1, C_entete_ListeVIN).Value = "" Then
                Cells(L_entete_ListeVIN + 1, C_entete_ListeVIN) = Liste_VIN(1, 1)
                Cells(L_entete_ListeVIN + 1, C_entete_ListeVIN + 1) = Info_Voiture
                Cells(L_entete_ListeVIN + 1, C_entete_ListeVIN + 2) = Marque_Voiture
            Else
                Cells(L_entete_ListeVIN, C_entete_ListeVIN).End(xlDown).Offset(1, 0) = Liste_VIN(1, 1)
                Cells(L_entete_ListeVIN, C_entete_ListeVIN + 1).End(xlDown).Offset(1, 0) = Info_Voiture
                Cells(L_entete_ListeVIN, C_entete_ListeVIN + 2).End(xlDown).Offset(1, 0) = Marque_Voiture
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

    ws_Trajet.Activate
' V1.8
    FormeEncours.TexteEnCours = "Traitement des donn�es en cours, patience..."
    FormeEncours.Repaint
    
' V1.6 : stockage dans un tableau interne (que l'on vide d'abord) de tous les trajets d�j� dans l'Excel
    For i = 1 To G_Nb_Trajets_Max
        T_Trajets_existants(i) = ""
    Next i
    ' Pour vraiment avoir la derni�re ligne du tableau rempli sans risquer d'effacer une donn�e, il faut r�initialiser les filtres
    ' Sinon la derni�re ligne remplie ne sera que celle affich�e, toutes celles masqu�es au-dessous ne compteront pas...
    Range("$A$3:$B$150000").AutoFilter Field:=1
    Derniere_Ligne_remplie = Cells(Columns(1).Cells.Count, 1).End(xlUp).Row
    Nb_Trajets = 0
    For i = L_Premiere_Valeur To Derniere_Ligne_remplie
        T_Trajets_existants(i - L_Premiere_Valeur + 1) = Cells(i, C_vin) & ";" & Cells(i, DicoJSON("id"))
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
                If ((nbTrip Mod 100) = 0) Then
                    FormeEncours.TexteEnCours = "Traitement trajets, " & nbTrip & " analys�s, patience..."
                    FormeEncours.Repaint
                End If
                Trajet_Trouve = False
                For j = 1 To Nb_Trajets
                    If T_Trajets_existants(j) = VIN_Actuel & ";" & item_item("id") Then
                        Trajet_Trouve = True
                        Exit For
                    End If
                Next j
                If Not Trajet_Trouve Then
                    Cells(i, C_vin).Value = VIN_Actuel      ' On �crit le VIN r�cup�r�
                    Cells(i, DicoJSON("id")).Value = item_item("id")  ' On �crit l'ID r�cup�r�
                    ' R�cup�ration des dates
                    ' On stocke les deux dates (d�part et arriv�e) car il faut d�terminer le temps de parcours
                    ' qui ne doit pas �tre d�pendant d'un �ventuelle changement d'heure en cours de route
                    MaDate_UNIX_UTC_dep = item_item("startDateTime")            ' Date de d�part
                    MaDate_UNIX_UTC_arr = item_item("endDateTime")              ' Date d'arriv�e
                    MaDate_DST_dep = Date_UNIX_To_Date_DST(MaDate_UNIX_UTC_dep)
                    MaDate_DST_arr = Date_UNIX_To_Date_DST(MaDate_UNIX_UTC_arr)
                    Cells(i, DicoJSON("startDateTimeDATE")).Value = MaDate_DST_dep
                    Cells(i, DicoJSON("endDateTimeDATE")).Value = MaDate_DST_arr
                    ' Calcul de la dur�e du trajet en cours
                    duree_trajet = MaDate_UNIX_UTC_arr - MaDate_UNIX_UTC_dep
                    duree_trajet_bis = Date_UNIX_To_Date_UTC(MaDate_UNIX_UTC_arr) - Date_UNIX_To_Date_UTC(MaDate_UNIX_UTC_dep)
                    ' V1.9 : en cas de souci entre date d�part et date arriv�e, et donc si duree_trajet est < 0, on met 1 seconde par d�faut
                    If duree_trajet < 0 Then
                        duree_trajet = 1
                        duree_trajet_bis = "00:00:01"
                    End If
                    Cells(i, C_duree).Value = duree_trajet_bis
                    
                    distance_trajet = item_item("endMileage") - item_item("startMileage")
                    Cells(i, C_dist).Value = distance_trajet
                    Cells(i, DicoJSON("endMileage")).Value = item_item("endMileage")
                    
                    Cells(i, DicoJSON("consumption")).Value = item_item("consumption")
                    conso_trajet = Cells(i, DicoJSON("consumption")).Value
                    ' Pour le calcul de la consommation moyenne, il faut �viter la division par z�ro dans le cas o�
                    ' la voiture � tourner � l'arret, la distance parcourue est nulle
                    If distance_trajet <> 0 Then
                        Cells(i, C_conso_moy).Value = conso_trajet / distance_trajet * 100
                    Else
                        Cells(i, C_conso_moy).Value = "//"
                    End If
        
                    Cells(i, DicoJSON("startPosLatitude")).Value = item_item("startPosLatitude")
                    Cells(i, DicoJSON("startPosLongitude")).Value = item_item("startPosLongitude")
                                
                    adresse_dep = item_item("startPosAddress")
                    Cells(i, DicoJSON("startPosAddress")).Value = Correction_Adresse(adresse_dep)
                    Cells(i, DicoJSON("endPosLatitude")).Value = item_item("endPosLatitude")
                    Cells(i, DicoJSON("endPosLongitude")).Value = item_item("endPosLongitude")
                    
                    adresse_arr = item_item("endPosAddress")
                    Cells(i, DicoJSON("endPosAddress")).Value = Correction_Adresse(adresse_arr)
                    
                    Cells(i, DicoJSON("fuelLevel")).Value = item_item("fuelLevel") / 100
                    Cells(i, DicoJSON("fuelAutonomy")).Value = item_item("fuelAutonomy")
                    
' V2.0 : On ajoute les donn�es inutilis�es dans les colonnes masqu�es
                    ' On �crit les donn�es non utilis�es pour la future reconstruction du fichier de donn�es
                    Cells(i, DicoJSON("startMileage")).Value = item_item("startMileage")
                    Cells(i, DicoJSON("distance")).Value = item_item("distance")
                    Cells(i, DicoJSON("destLatitude")).Value = item_item("destLatitude")
                    Cells(i, DicoJSON("destLongitude")).Value = item_item("destLongitude")
                    adresse_tmp = item_item("destAddress")
                    Cells(i, DicoJSON("destAddress")).Value = Correction_Adresse(adresse_tmp)
                    Cells(i, DicoJSON("destQuality")).Value = item_item("destQuality")
                    Cells(i, DicoJSON("maintenanceDays")).Value = item_item("maintenanceDays")
                    Cells(i, DicoJSON("maintenanceDistance")).Value = item_item("maintenanceDistance")
                    Cells(i, DicoJSON("maintenancePassed")).Value = item_item("maintenancePassed")
                    Cells(i, DicoJSON("startPosQuality")).Value = item_item("startPosQuality")
                    Cells(i, DicoJSON("endPosQuality")).Value = item_item("endPosQuality")
                    
                    ' R�cup�ration des valeurs d'erreurs dans les tableaux du JSON
                    chaine_tmp = ""
                    nb_item_erreur = item_item("alertsActive").Count    ' On comptabilise le nombre d'erreurs "alertsActive" dans les donn�es
                    For Each item_error In item_item("alertsActive")    ' On r�cup�re toutes les erreurs "alertsActive"
                        If (nb_item_erreur > 1) Then
                            ' Il y a plus d'une erreur, donc on les s�pare par un ;
                            chaine_tmp = chaine_tmp & item_error("code") & ";"
                        Else    ' Derni�re ou unique valeur du tableau, on ne met pas de ;
                            chaine_tmp = chaine_tmp & item_error("code")
                        End If
                        nb_item_erreur = nb_item_erreur - 1
                    Next
                    Cells(i, DicoJSON("alertsActive")).Value = chaine_tmp   ' On �crit la valeur format�e avec des ; dans la cellule ad�quate

                    chaine_tmp = ""
                    nb_item_erreur = item_item("alertsResolved").Count    ' On comptabilise le nombre d'erreurs "alertsResolved" dans les donn�es
                    For Each item_error In item_item("alertsResolved")    ' On r�cup�re toutes les erreurs "alertsResolved"
                       If (nb_item_erreur > 1) Then
                            ' Il y a plus d'une erreur, donc on les s�pare par un ;
                            chaine_tmp = chaine_tmp & item_error("code") & ";"
                        Else    ' Derni�re ou unique valeur du tableau, on ne met pas de ;
                            chaine_tmp = chaine_tmp & item_error("code")
                        End If
                        nb_item_erreur = nb_item_erreur - 1
                    Next
                    Cells(i, DicoJSON("alertsResolved")).Value = chaine_tmp   ' On �crit la valeur format�e avec des ; dans la cellule ad�quate
                    
                    ' R�cup�ration des dates UNIX UTC du fichier de donn�e
                    Cells(i, DicoJSON("startDateTime")).Value = item_item("startDateTime")
                    Cells(i, DicoJSON("endDateTime")).Value = item_item("endDateTime")
                    
'                    ' Ceci est pour tester la correction INVERSE d'adresses. Il se sera pas n�cessaire de l'utiliser ici, mais uniquement lors de la re-cr�ation du fichier de donn�es JSON
'                    Cells(i, 33).Value = Correction_Adresse_INVERSE(Cells(i, DicoJSON("endPosAddress")).Value)
'                    Cells(i, 34).Value = item_item("endPosAddress")
'
'                    If Cells(i, 33).Value = Cells(i, 34).Value Then
'                        Cells(i, 35).Value = True
'                    Else
'                        Cells(i, 35).Value = False
'                    End If
                    
'                    ' Pour tester fonction de conversion de temps DATE DST vers Unix UTC-------------------
'                    ' Devra �tre supprim�e apr�s v�rification
'                    Cells(i, 33).Value = Date_DST_To_Date_UNIX_UTC(Cells(i, DicoJSON("startDateTimeDATE")).Value)
'                    If Cells(i, 33).Value = Cells(i, DicoJSON("startDateTime")).Value Then
'                        Cells(i, 34).Value = True
'                    Else
'                        Cells(i, 34).Value = False
'                    End If
'                    Cells(i, 35).Value = Date_DST_To_Date_UNIX_UTC(Cells(i, DicoJSON("endDateTimeDATE")).Value)
'                    If Cells(i, 35).Value = Cells(i, DicoJSON("endDateTime")).Value Then
'                        Cells(i, 36).Value = True
'                    Else
'                        Cells(i, 36).Value = False
'                    End If
'                    '--------------------------------------------------------------------------------------
' Fin d'ajout pour la V2
                    
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
    Range(CELL_plage_1ereligne).Select
    Range(Selection, Selection.End(xlDown)).Select
    ws_Trajet.Sort.SortFields.Clear
    ws_Trajet.Sort.SortFields.Add Key:=Range(CELL_plage_max_COL_vin), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal     ' Range de base "A5:A65536"
    ws_Trajet.Sort.SortFields.Add Key:=Range(CELL_plage_max_COL_id), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal      ' Range de base "B5:B65536"
    With ws_Trajet.Sort
        .SetRange Range(CELL_plage_max_Donnees)        ' range par d�faut "A5:Q65536"
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Cells(1, 4).Select   ' On s�lection la cellule de version pour ne pas avoir tout le tableau s�lectionn�
' V1.7 : mise en place du formatage par colonne
' Colonne VIN
    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_vin, f_size:=8
' Colonne Date d�part
    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=DicoJSON("startDateTimeDATE"), n_format:="date"
' Colonne Date Arriv�e
    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=DicoJSON("endDateTimeDATE"), n_format:="date"
' Colonne Dur�e
    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_duree, n_format:="duree"
' Colonne Distance
    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_dist, n_format:="1"
' Colonne Distance totale
    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=DicoJSON("endMileage"), n_format:="1"
' Colonne conso
    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=DicoJSON("consumption"), n_format:="2"
' Colonne conso moyenne
    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_conso_moy, n_format:="1"
' Colonne niveau carburant
    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=DicoJSON("fuelLevel"), n_format:="%"
' Colonne adresse D�part
    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=DicoJSON("startPosAddress"), n_format:="add"
' Colonne adresse Arriv�e
    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=DicoJSON("endPosAddress"), n_format:="add"
    
' V1.9.3 : pour indiquer que la macro est en cours
    Macro_en_cours = True

' V1.8 : rafraichissement du TCD. On positionne tous les VIN � "coch�" sauf le VIN "vide" par d�faut
    ws_Accueil.PivotTables("TCD_VIN").PivotCache.Refresh
    ws_Accueil.PivotTables("TCD_VIN").PivotFields("VIN(s)").CurrentPage = "(All)"
' V1.9 : s�lection de tous les vin import�s dans le TCD. On les d�selectionne tous d'abord
    For Each i_TCD In ws_Accueil.PivotTables("TCD_VIN").PivotFields("VIN(s)").PivotItems
        If i_TCD.Name = "(blank)" Then
            i_TCD.Visible = True
        Else
            i_TCD.Visible = False
        End If
    Next
    ws_Accueil.PivotTables("TCD_VIN").PivotFields("VIN(s)").EnableMultiplePageItems = True
    ' V1.9.2 : recherche si le VIN a �t� s�lectionn�, et on compte combien il y en a
    Nb_VIN_Selection = 0
    For i = 1 To UBound(Liste_VIN)
        If Liste_VIN(i, 2) Then
            On Error Resume Next
            ws_Accueil.PivotTables("TCD_VIN").PivotFields("VIN(s)").PivotItems(CStr(Liste_VIN(i, 1))).Visible = True
            On Error GoTo 0
            Nb_VIN_Selection = Nb_VIN_Selection + 1
        End If
    Next i
    ws_Accueil.PivotTables("TCD_VIN").PivotFields("VIN(s)").PivotItems("(blank)").Visible = False
' V1.8
    FormeEncours.Hide
    ws_Trajet.Activate
' V1.9 : s�lection du 1er vin import� dans le filtre
    Nb_VIN_Filtre = 0
    For i = 1 To UBound(Liste_VIN)
        If Liste_VIN(i, 2) Then
            Nb_VIN_Filtre = Nb_VIN_Filtre + 1
            Critere(Nb_VIN_Filtre) = Liste_VIN(i, 1)
            Select Case i
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
    Next
    Cells(1, 1).Select ' On s�lection la cellule A1 ne pas avoir tout le tableau s�lectionn�
' V1.9.3 : D�termination de la derni�re adresse atteinte
    Derniere_Lat = ""
    Derniere_Long = ""
    Derniere_Adr = ""
    If Nb_VIN_Selection = 1 Then
      ' uniquement si un seul vin s�lectionn�
        Range("A2").Select
        Selection.End(xlDown).Select
        Derniere_Lat = Cells(Selection.Row, DicoJSON("endPosLatitude")).Value
        Derniere_Long = Cells(Selection.Row, DicoJSON("endPosLongitude")).Value
        Derniere_Adr = Cells(Selection.Row, DicoJSON("endPosAddress")).Value
    End If
    ws_Accueil.Activate
    Range(CELL_lat) = Derniere_Lat
    Range(CELL_long) = Derniere_Long
    Range(CELL_adr) = Derniere_Adr
' V1.9.3 : pour indiquer que la macro est termin�e
    Macro_en_cours = False
' Ce qui suit ne fonctionne pas, j'ai une erreur sur le .AddItem...
'    Sheets("Accueil").ComboBox1.Clear
'    Sheets("Accueil").ComboBox1.Style = fmStyleDropDownList
'    Sheets("Accueil").ComboBox1.AddItem = "TOTO3"
'    Sheets("Accueil").ComboBox1.AddItem = "TOTO3"
'    Sheets("Accueil").ComboBox1.Clear

    ws_Accueil.Activate
    Cells(1, 1).Select ' On s�lection la cellule A1 ne pas avoir tout le tableau s�lectionn�
' V1.7 : optimisation : remise de la mise � jour de l'affichage
    Application.ScreenUpdating = True
End Sub

Private Sub Formater_Cellules(ws_tmp As Worksheet, ligne_cell As Variant, colonne_cell As Variant, Optional n_format As String = "General", Optional f_size As Integer = 10, Optional wrap As Boolean = False)
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
            ws_tmp.Range(ws_tmp.Cells(ligne_cell, colonne_cell), ws_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).NumberFormat = "dd/mm/yy - hh:mm"
        Case "duree"
            ws_tmp.Range(ws_tmp.Cells(ligne_cell, colonne_cell), ws_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).NumberFormat = "h:mm"
        Case "0"
            ws_tmp.Range(ws_tmp.Cells(ligne_cell, colonne_cell), ws_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).NumberFormatLocal = "0"
        Case "1"
            ws_tmp.Range(ws_tmp.Cells(ligne_cell, colonne_cell), ws_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).NumberFormatLocal = "0,0"
        Case "2"
            ws_tmp.Range(ws_tmp.Cells(ligne_cell, colonne_cell), ws_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).NumberFormatLocal = "0,00"
        Case "%"
            ws_tmp.Range(ws_tmp.Cells(ligne_cell, colonne_cell), ws_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).NumberFormat = "0 %"
        Case Else
            ws_tmp.Range(ws_tmp.Cells(ligne_cell, colonne_cell), ws_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).NumberFormat = "General"
    End Select
    
    If n_format = "add" Then        ' Il faut v�rifier si on est sur un champ adresse car pour l'adresse il faut aligner � gauche xlHAlignLeft
        ws_tmp.Range(ws_tmp.Cells(ligne_cell, colonne_cell), ws_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).HorizontalAlignment = xlHAlignLeft
    Else
        ws_tmp.Range(ws_tmp.Cells(ligne_cell, colonne_cell), ws_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).HorizontalAlignment = xlHAlignCenter
    End If
    ' Dans tous les cas, le VerticalAlignment est � xlVAlignCenter
    ws_tmp.Range(ws_tmp.Cells(ligne_cell, colonne_cell), ws_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).VerticalAlignment = xlVAlignCenter
    ws_tmp.Range(ws_tmp.Cells(ligne_cell, colonne_cell), ws_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).Font.Size = f_size
    
    If (wrap = True) Then
        ws_tmp.Range(ws_tmp.Cells(ligne_cell, colonne_cell), ws_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).WrapText = True
    Else
        ws_tmp.Range(ws_tmp.Cells(ligne_cell, colonne_cell), ws_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).WrapText = False
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
    Worksheets("Trajets").Range("$A$3:$B$150000").AutoFilter Field:=1
    With Worksheets("Trajets")
' 1.8        .Range(CELL_vin_entete) = ""           ' Le VIN
        .Range(CELL_fichierMYP) = ""           ' Le nom du fichier
' 1.8       .Range(CELL_nb_trips) = ""  ' Le nombre de trajet et le kilom�trage total
' 1.8       .Range(CELL_km) = ""
        .Range(CELL_plage_donnees) = ""    ' Le grand tableau de valeurs de tous les trajets effectu�s
' 1.8       .Range(CELL_conso_tot) = ""           ' La consommation totale sur tous les trajets
    End With
    Macro_en_cours = True
    With Worksheets("Accueil")
        .PivotTables("TCD_VIN").PivotCache.Refresh
        .Range("M4:O65000") = ""
    End With
    Macro_en_cours = False
    With Worksheets("Donnees-Recap")
        .Range("A2:D10000") = ""  ' Un tableau r�capitulatif mensuel
        .Range("F4:F53") = 0      ' Le nombre de trajets regroup� en plage de distance
        .Range("H4:H53") = 0      ' La distance des regroupements de trajets
    End With
End Sub

Private Function Date_DST_To_Date_UNIX(date_dst As Date) As Long
    ' Fonction qui convertie une date DST en temps UNIX DST donc avec changement d'heure
    '@PARAM {Long} date � convertir, au format date DST
    '@RETURN {Date} Renvoi la date convertie temps UNIX DST
    Date_DST_To_Date_UNIX = (date_dst - DateValue("01/01/1970")) * 86400
End Function

Private Function Date_DST_To_Date_UNIX_UTC(date_dst As Date) As Long
    ' Fonction qui convertie une date DST en temps UNIX UTC
    '@PARAM {Date} date � convertir, au format date DST
    '@RETURN {Long} Renvoi la date convertie en temps UNIX UTC
    Dim DST_val As Integer, date_unix_DST As Long
    ' La date ainsi calcul�e ne tient pas compte du passage � l'heure d'�t� ou � l'heure d'hiver
    ' Il faut v�rifier si le jour du mois de cette date est avant ou apr�s le dernier dimanche de mars ou d'octbre.
    ' Rappel :  l'heure passe en UTC +2 au dernier dimanche de mars, c'est l'heure d'�t�
    '           l'heure passe en UTC +1 au dernier dimanche d'octobre, c'est l'heure d'hiver
    ' On va donc devoir ajouter soit 1h=3600s au temps UTC si c'est l'heure d'hiver, soit 2h=7200s au temps UTC si c'est l'heure d'hiver
    ' D�terminons la valeur du facteur DST
    DST_val = DST_date(date_dst)
    
    ' Conversion en Date de cette date UNIX : la nouvelle Date est maintenant DST
    Date_DST_To_Date_UNIX_UTC = Date_DST_To_Date_UNIX(date_dst) - DST_val * 3600

End Function

Private Function DST_date(date_dst As Date) As Integer
    ' Fonction qui d�termine le modificateur d'heure (Day Saving Time) appliqu� � la date fournie
    ' Ce sera le modificateur d'heure par rapport au temps unix UTC :   1 = pour l'heure d'hiver
    '                                                                   2 = pour l'heure d'�t�
    '@PARAM {Date} date DST � tester, au format Date
    '@RETURN {Integer} Renvoi un entier permettant la modification de l'heure unix
    
    ' On d�clare les variables utilis�es pour le jour, le mois, l'ann�e, l'heure en temps UTC
    Dim jour_DST As Integer, mois_DST As Integer, annee_DST As Integer, heure_DST As Integer, minutes_DST As Integer
    Dim date_temp As Date
    
    ' On d�clare les variables utilis�es pour le dernier dimanche de mars ou d'octobre : jour/mois/heure
    Dim jour_dD As Integer, mois_dD As Integer, heure_dD As Integer
    Dim num_jour_31 As Integer      ' C'est pour stocker le n� du jour de semaine pour le 31/03/annee ou 31/10/annee
   
    ' On r�cup�re le jour, le mois, l'ann�e, l'heure en date DST
    jour_DST = Day(date_dst)
    mois_DST = Month(date_dst)
    annee_DST = Year(date_dst)
    heure_DST = Hour(date_dst)
    ' minutes = Minute(date_temp)       ' Inutile ici
     
    Select Case mois_DST
        Case 1, 2, 11, 12
            ' On est dans le cas o� l'heure d'hiver est appliqu�e : de Novembre � F�vrier
            DST_date = 1
            
        Case 4 To 9
            ' On est dans le cas o� l'heure d'�t� est appliqu�e : de Avril � Septembre
            DST_date = 2
            
        Case 3
            ' On est en mars, il faut v�rifier que le jour en question est avant ou apr�s le dernier dimanche de mars
            ' D�termination du n� du jour de la semaine du dernier jour de mars
            num_jour_31 = Weekday("31/03/" & annee_DST, vbMonday)
            If num_jour_31 = 7 Then
                ' Le 31 c'est le dimanche
                jour_dD = 7
            Else
                jour_dD = 31 - num_jour_31
            End If
                            
            If jour_DST < jour_dD Then
                ' On est avant le dernier dimanche du mois, donc encore en heure d'hiver
                DST_date = 1
            ElseIf jour_DST > jour_dD Then
                ' On est apr�s le dernier dimanche du mois, donc en heure d'�t�
                DST_date = 2
            ElseIf jour_DST = jour_dD Then
                ' On est le dernier dimanche du mois, donc il faut voir si on est avant 1h UTC donc 2h DST du matin ou pas.
                If heure_DST < 2 Then
                    ' On est encore � l'heure d'hiver
                    DST_date = 1
                Else
                    ' On est pass� � l'heure d'�t�
                    DST_date = 2
                End If
            End If
            
        Case 10
            ' On est en octobre, il faut v�rifier que le jour en question est avant ou apr�s le dernier dimanche de mars
            ' D�termination du n� du jour de la semaine du dernier jour d'octobre
            num_jour_31 = Weekday("31/10/" & annee_DST, vbMonday)
            jour_dD = 31 - num_jour_31
            If num_jour_31 = 7 Then
                ' Le 31 c'est le dimanche, c'est donc lui le dernier dimanche du mois !
                jour_dD = 7
            Else
                ' Le 31 est un autre jour de la semaine, on calcule donc quel sera le jour XX du dernier dimanche du mois
                jour_dD = 31 - num_jour_31
            End If
            
            If jour_DST < jour_dD Then
                ' On est avant le dernier dimanche du mois, donc encore en heure d'hiver
                DST_date = 1
            ElseIf jour_DST > jour_dD Then
                ' On est apr�s le dernier dimanche du mois, donc en heure d'�t�
                DST_date = 2
            ElseIf jour_DST = jour_dD Then
                ' On est le dernier dimanche du mois, donc il faut voir si on est avant 1h UTC donc 3h DST du matin ou pas.
                If heure_DST < 3 Then
                    ' On est encore � l'heure d'�t�
                    DST_date = 2
                Else
                    ' On est pass� � l'heure d'hiver
                    DST_date = 1
                End If
            End If
        End Select
        
End Function
Sub test()
    Dim dateUNIX_retrouvee As Long, dateUNIX_ref As Long, dateDST As Date
    Dim tmp As Long
    ' format UNIX � retrouver / Date Excel  /   Valeur dans la cellule
    '   D�but = 1579799640  /   23/01/20 - 18:14    /   43853.7597222222
    '   fin =   1579802820  /   23/01/20 - 19:07    /   43853.7965277778
    ' Autres valeurs :
    '   D�but = 1587111900  /   17/04/20 - 10:25    /   43938.4340277778
    '   Fin =   1587112380  /   17/04/20 - 10:33    /   43938.4395833333
        
    dateDST = 43938.4340277778
    dateUNIX_ref = 1587111900
    dateUNIX_retrouvee = Date_DST_To_Date_UNIX_UTC(dateDST)
    tmp = dateUNIX_retrouvee - dateUNIX_ref
    Debug.Print "val.ref.....= " & dateUNIX_ref & vbLf & "val.retrouv.= " & dateUNIX_retrouvee & vbLf & "Diff�rence entre retrouv�e et ref : " & tmp & vbLf & "---------"
    
    
End Sub
Private Function Date_UNIX_To_Date_UTC(date_unix_UTC As Long) As Date
    ' Fonction qui converti un temps UNIX en date UTC
    '@PARAM {Long} date � convertir, au format UNIX UTC
    '@RETURN {Date} Renvoi la date convertie en date UTC

    Date_UNIX_To_Date_UTC = (date_unix_UTC / 86400) + DateValue("01/01/1970")
End Function

Private Function Date_UNIX_To_Date_DST(date_unix_UTC As Long) As Date
    ' Fonction qui converti un temps UNIX en date avec DST (changement d'heure)
    '@PARAM {Long} date � convertir, au format UNIX UTC
    '@RETURN {Date} Renvoi la date convertie en date UTC
    Dim DST_val As Integer, date_unix_DST As Long
    ' La date ainsi calcul�e ne tient pas compte du passage � l'heure d'�t� ou � l'heure d'hiver
    ' Il faut v�rifier si le jour du mois de cette date est avant ou apr�s le dernier dimanche de mars ou d'octbre.
    ' Rappel :  l'heure passe en UTC +2 au dernier dimanche de mars, c'est l'heure d'�t�
    '           l'heure passe en UTC +1 au dernier dimanche d'octobre, c'est l'heure d'hiver
    ' On va donc devoir ajouter soit 1h=3600s au temps UTC si c'est l'heure d'hiver, soit 2h=7200s au temps UTC si c'est l'heure d'hiver
    ' D�terminons la valeur du facteur DST
    DST_val = DST(date_unix_UTC)
    
    ' Calcul de la nouvelle date avec DST
    date_unix_DST = date_unix_UTC + DST_val * 3600
    ' Conversion en Date de cette date UNIX : la nouvelle Date est maintenant DST
    Date_UNIX_To_Date_DST = Date_UNIX_To_Date_UTC(date_unix_DST)

End Function

Private Function DST(date_unix_UTC As Long) As Integer
    ' Fonction qui d�termine le modificateur d'heure (Day Saving Time) � appliquer � l'heure UTC pour avoir l'heure FR
    ' Ce sera le modificateur d'heure par rapport au temps unix UTC :   1 = pour l'heure d'hiver
    '                                                                   2 = pour l'heure d'�t�
    '@PARAM {Long} date � tester, au format unix UTC
    '@RETURN {Integer} Renvoi un entier permettant la modification de l'heure unix


    ' On d�clare les variables utilis�es pour le jour, le mois, l'ann�e, l'heure en temps UTC
    Dim jour_UTC As Integer, mois_UTC As Integer, annee_UTC As Integer, heure_UTC As Integer, minutes_UTC As Integer
    Dim date_temp As Date
    
    ' On d�clare les variables utilis�es pour le dernier dimanche de mars ou d'octobre : jour/mois/heure
    Dim jour_dD As Integer, mois_dD As Integer, heure_dD As Integer
    Dim num_jour_31 As Integer      ' C'est pour stocker le n� du jour de semaine pour le 31/03/annee ou 31/10/annee
   
    ' On convertir la date unix en date UTC
    date_temp = Date_UNIX_To_Date_UTC(date_unix_UTC)
    
    ' On r�cup�re le jour, le mois, l'ann�e, l'heure en temps UTC
    jour_UTC = Day(date_temp)
    mois_UTC = Month(date_temp)
    annee_UTC = Year(date_temp)
    heure_UTC = Hour(date_temp)
    ' minutes = Minute(date_temp)       ' Inutile ici
     
    Select Case mois_UTC
        Case 1, 2, 11, 12
            ' On est dans le cas o� l'heure d'hiver est appliqu�e : de Novembre � F�vrier
            DST = 1
            
        Case 4 To 9
            ' On est dans le cas o� l'heure d'�t� est appliqu�e : de Avril � Septembre
            DST = 2
            
        Case 3
            ' On est en mars, il faut v�rifier que le jour en question est avant ou apr�s le dernier dimanche de mars
            ' D�termination du n� du jour de la semaine du dernier jour de mars
            num_jour_31 = Weekday("31/03/" & annee_UTC, vbMonday)
            If num_jour_31 = 7 Then
                ' Le 31 c'est le dimanche
                jour_dD = 7
            Else
                jour_dD = 31 - num_jour_31
            End If
                            
            If jour_UTC < jour_dD Then
                ' On est avant le dernier dimanche du mois, donc encore en heure d'hiver
                DST = 1
            ElseIf jour_UTC > jour_dD Then
                ' On est apr�s le dernier dimanche du mois, donc en heure d'�t�
                DST = 2
            ElseIf jour_UTC = jour_dD Then
                ' On est le dernier dimanche du mois, donc il faut voir si on est avant 1h UTC du matin ou pas.
                If heure_UTC < 1 Then
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
            num_jour_31 = Weekday("31/10/" & annee_UTC, vbMonday)
            jour_dD = 31 - num_jour_31
            If num_jour_31 = 7 Then
                ' Le 31 c'est le dimanche, c'est donc lui le dernier dimanche du mois !
                jour_dD = 7
            Else
                ' Le 31 est un autre jour de la semaine, on calcule donc quel sera le jour XX du dernier dimanche du mois
                jour_dD = 31 - num_jour_31
            End If
            
            If jour_UTC < jour_dD Then
                ' On est avant le dernier dimanche du mois, donc encore en heure d'hiver
                DST = 1
            ElseIf jour_UTC > jour_dD Then
                ' On est apr�s le dernier dimanche du mois, donc en heure d'�t�
                DST = 2
            ElseIf jour_UTC = jour_dD Then
                ' On est le dernier dimanche du mois, donc il faut voir si on est avant 1h UTC du matin ou pas.
                If heure_UTC < 1 Then
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
    Dim addr_tmp As String
    addr_tmp = ""   ' On va mettre les caract�res un par un dans cette variable
    
    adresse = Replace(adresse, vbLf, ", ")   ' On remplace tous les retours � la ligne "\n" par des ", "
    For i = 1 To Len(adresse)
        lettre = Mid(adresse, i, 2)
        Select Case lettre
            Case "è"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "é"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "ê"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "À"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "â"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "œ"   ' oe
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "à"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "ç"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "ï"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "ù"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "ü"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "É"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "È"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "Ê"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "Ê"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "Ë"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "ë"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "Ç"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "Â"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "Ü"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "Û"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "û"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "Ï"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "Î"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "î"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "Ô"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "ô"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "Ö"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)
            Case "ö"   ' �
                adresse = Left(adresse, i - 1) + "�" + Mid(adresse, i + 2)

        End Select
    Next i
    Correction_Adresse = adresse
End Function

Private Function Correction_Adresse_INVERSE(ByVal adresse As String) As String
    Dim i As Integer
    Dim lettre As String * 1
    Dim addr_tmp As String
    addr_tmp = ""   ' On va mettre les caract�res un par un dans cette variable
    adresse = Replace(adresse, ", ", vbLf)   ' On remplace tous les retours � la ligne ", " par des "\n"
    
    For i = 1 To Len(adresse)
        lettre = Mid(adresse, i, 1)
        Select Case lettre                      ' On reconverti toutes les lettres accentu�es trouv�es par leur version UTF8
            Case "�"   ' �
                addr_tmp = addr_tmp + "è"
            Case "�"   ' �
                addr_tmp = addr_tmp + "é"
            Case "�"   ' �
                addr_tmp = addr_tmp + "ê"
            Case "�"   ' �
                addr_tmp = addr_tmp + "À"
            Case "�"   ' �
                addr_tmp = addr_tmp + "â"
            Case "�"   ' oe
                addr_tmp = addr_tmp + "œ"
            Case "�"   ' �
                addr_tmp = addr_tmp + "à"
            Case "�"   ' �
                addr_tmp = addr_tmp + "ç"
            Case "�"   ' �
                addr_tmp = addr_tmp + "ï"
            Case "�"   ' �
                addr_tmp = addr_tmp + "ù"
            Case "�"   ' �
                addr_tmp = addr_tmp + "ü"
            Case "�"   ' �
                addr_tmp = addr_tmp + "É"
            Case "�"   ' �
                addr_tmp = addr_tmp + "È"
            Case "�"   ' �
                addr_tmp = addr_tmp + "Ê"
            Case "�"   ' �
                addr_tmp = addr_tmp + "Ê"
            Case "�"   ' �
                addr_tmp = addr_tmp + "Ë"
            Case "�"   ' �
                addr_tmp = addr_tmp + "ë"
            Case "�"   ' �
                addr_tmp = addr_tmp + "Ç"
            Case "�"   ' �
                addr_tmp = addr_tmp + "Â"
            Case "�"   ' �
                addr_tmp = addr_tmp + "Ü"
            Case "�"   ' �
                addr_tmp = addr_tmp + "Û"
            Case "�"   ' �
                addr_tmp = addr_tmp + "û"
            Case "�"   ' �
                addr_tmp = addr_tmp + "Ï"
            Case "�"   ' �
                addr_tmp = addr_tmp + "Î"
            Case "�"   ' �
                addr_tmp = addr_tmp + "î"
            Case "�"   ' �
                addr_tmp = addr_tmp + "Ô"
            Case "�"   ' �
                addr_tmp = addr_tmp + "ô"
            Case "�"   ' �
                addr_tmp = addr_tmp + "Ö"
            Case "�"   ' �
                addr_tmp = addr_tmp + "ö"
            Case Else   ' pour tous les autres caract�res, on les reopie tel-quel
                addr_tmp = addr_tmp + lettre
        End Select
    Next i

    Correction_Adresse_INVERSE = addr_tmp
End Function

Sub Bouton_Accueil_Cliquer()
' Activer feuille Accueil
    Sheets("Accueil").Activate
End Sub
Sub Bouton_Trajets_Cliquer()
' Activer feuille Trajets
    Sheets("Trajets").Activate
End Sub
Sub Bouton_Tuto_Cliquer()
' Activer la feuille Tutoriel
    Sheets("Tutoriel").Activate
End Sub
Sub RemplisDicoJSON()
    DicoJSON.Add "id", 2                    ' Identificateur trajet
    DicoJSON.Add "startDateTimeDATE", 3     ' Date DST calcul�e � partir du temps UNIX UTC fourni par le fichier de donn�e
    DicoJSON.Add "endDateTimeDATE", 4       ' Date DST calcul�e � partir du temps UNIX UTC fourni par le fichier de donn�e
    DicoJSON.Add "endMileage", 7
    DicoJSON.Add "startMileage", 18         ' Donn�e suppl�mentaire � ajouter dans la partie masqu�e du tableau
    DicoJSON.Add "consumption", 8
    DicoJSON.Add "startPosLatitude", 10
    DicoJSON.Add "startPosLongitude", 11
    DicoJSON.Add "startPosAddress", 12
    DicoJSON.Add "endPosLatitude", 13
    DicoJSON.Add "endPosLongitude", 14
    DicoJSON.Add "endPosAddress", 15
    DicoJSON.Add "fuelLevel", 16
    DicoJSON.Add "fuelAutonomy", 17
' Donn�es non utilis�es du fichier de donn�es charg�
    DicoJSON.Add "distance", 19
    DicoJSON.Add "destLatitude", 20
    DicoJSON.Add "destLongitude", 21
    DicoJSON.Add "destAddress", 22
    DicoJSON.Add "destQuality", 23
    DicoJSON.Add "maintenanceDays", 24
    DicoJSON.Add "maintenanceDistance", 25
    DicoJSON.Add "maintenancePassed", 26
    DicoJSON.Add "startPosQuality", 27
    DicoJSON.Add "endPosQuality", 28
    DicoJSON.Add "alertsActive", 29
    DicoJSON.Add "alertsResolved", 30
    
    DicoJSON.Add "startDateTime", 31    ' Je place ici ces valeurs car ce sont celles fournies par le fichier de donn�es
    DicoJSON.Add "endDateTime", 32      ' Je place ici ces valeurs car ce sont celles fournies par le fichier de donn�es
    DicoRempli = True
End Sub

