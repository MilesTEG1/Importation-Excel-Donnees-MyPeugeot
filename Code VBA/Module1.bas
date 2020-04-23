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
'       - V 1.9.4 : Ajout d'une feuille Tutoriel expliquant les diff�rentes fonctions d'importation du fichier XLSM.
'       - V 1.9.5 : Ajout d'une d�cimale sur les kilom�tres totaux affich�s (colonne G)
'       - V 2.0 :   Ajout des valeurs non utilis�e dans des colonnes masqu�es en vue de l'exportation des donn�es en format fichier JSON
'                   Mise en place d'une structure de type Dictionnaire pour stocker les colonnes utilis�es pour les donn�es
'                   Ajout de la marque de la voiture (d�tect�e automatiquement avec l'extension du fichier de donn�es fourni)
'                   Ajout de 3 fonctions pour d�terminer un temps UNIX UTC � partir d'une date DST (pour la reconstruction d'un trajet manquant)
'                   Ajout de nouveaux caract�res accentu�s pour la conversion d'adresses
'                   Ajout d'une fonction de correction inverse des adresses en vue de la re-cr�ation d'un fichier de donn�es � partir du tableau excel (afin de tenir compte des �ventuelles modifications d'adresses par les utilisateurs).
'                   Ajout d'une proc�dure d'exporation des donn�es dans le format JSON
'                       Il est � noter que lors de l'importation des donn�es, et donc � l'exportation, certaines valeurs perdent quelque peu en pr�cision,
'                       car les nombres ne peuvent avoir plus de 15 chiffres, donc ceux qui en ont plus sont arrondis � 15 chiffres en tout.
'                       C'est la concession � faire pour avoir les coordonn�es GPS sans arrondi, donc les plus pr�cises.
'                       Cela concerne les donn�es suivantes : "consumption" et "distance"
'                   Ajout d'une feuille Tutoriel pour l'exportation
'
' Couples de versions d'Excel & OS test�es :
'       - Windows 10 v1909 (18363.752) & Excel pour Office 365 Version 2003 (16.0 build 12624.20382 & 16.0 build 12624.20442)
'       - Windows 10 v1909 (18363.815) & Excel pour Microsoft 365 Version 2003 (16.0 build 12624.20466)
'       - Windows 10 v1809 & Excel 2016
'       - Windows 10 v1909 & Excel 2019
'
Option Explicit     ' On force � d�clarer les variables
Option Base 1       ' les tableaux commenceront � l'indice 1
'
' D�inissons quelques constantes qui serviront pour les colonnes/lignes/plages de cellules.
'
Const VERSION As String = "v2.0 B�ta 19"     ' Version du fchier
Const CELL_ver As String = "B3"             ' Cellule o� afficher la version du fichier
'Const var_DEBUG As Boolean = True       ' True =    On active un mode DEBUG o� on affiche certaines choses
Const var_DEBUG As Boolean = False      ' False =   On d�sactive un mode DEBUG o� on affiche certaines choses
' Constantes pour la feuille "Trajets"
Const L_Premiere_Valeur As Integer = 3      ' Premi�re ligne � contenir des donn�es (avant ce sont les lignes d'en-t�te
Const C_vin             As Integer = 1      ' COLONNE = 1 (A) -> VIN pour le trajet (il est possible d'avoir plusieurs VIN dans le fichier de donn�e).
Const C_startDate       As Integer = 3      ' Date DST calcul�e � partir du temps UNIX UTC fourni par le fichier de donn�e
Const C_endDate         As Integer = 4      ' Date DST calcul�e � partir du temps UNIX UTC fourni par le fichier de donn�e
Const C_duree           As Integer = 5      '         = 5 (E) -> Dur�e du trajet  (� calculer)
Const C_dist            As Integer = 6      '         = 6 (F) -> Distance du trajet en km
Const C_conso_moy       As Integer = 9      '         = 9 (I) -> Consommation moyenne en L/100km
Const C_marque          As Integer = 33     '         = 33 (AG) -> Marque de la voiture
Const CELL_fichierMYP   As String = "G" & (L_Premiere_Valeur - 2)  ' Cellule qui contiendra le nom du fichier import�
Const CELL_plage_donnees    As String = "A" & L_Premiere_Valeur & ":AJ65536"    ' Plage de cellules contenant les donn�es
Const CELL_plage_1ereligne  As String = "A" & L_Premiere_Valeur & ":AJ" & L_Premiere_Valeur  ' Toute la 1ere ligne de donn�e pour en faire le tri
Const CELL_plage_max_Donnees As String = "A" & L_Premiere_Valeur & ":AJ65536"   ' La plage de donn�es maximale possible
Const CELL_plage_max_COL_vin As String = "A" & L_Premiere_Valeur & ":A65536"    ' La plage de donn�es maximale possible
Const CELL_plage_max_COL_id As String = "B" & L_Premiere_Valeur & ":B65536"     ' La plage de donn�es maximale possible
'V2.B18 : Constante d�finissant les colonnes masqu�es
Const COLs_Masquees     As String = "R:AI"  ' Colonne R = 18 -/- Colonne AI = 35
' V1.5 : MultiVIN
Const G_Nb_Trajets_Max = 20000              ' Nb trajets max par VIN trait�s par cette macro

' Constantes pour la feuille Accueil
Const C_entete_ListeVIN        As Integer = 13      ' = 13 (M) Colonne d'ent�te des VIN dans la liste des vins r�cup�r�s
                                                    ' la colonne des descriptions des v�hicules est celle d'� cot� : 13+1 = N
                                                    ' le colonne de la marque des v�hicule est celle d'� cot� encore : 13+2 = O
Const L_entete_ListeVIN        As Integer = 3       ' Ligne d'ent�te des VIN dans la liste des vins r�cup�r�s, elle correspond aussi � celles des descriptions des v�hicules

Const Nb_marques As Integer = 3
'
' Variables g�n�rales
Public Const CELL_lat  As String = "H9"
Public Const CELL_long As String = "H8"
Public Const CELL_adr As String = "E8"
Public Const G_Nb_VIN_Max = 20                      ' Nb VIN max trait� par cette macro
Public Macro_en_cours, DicoRempli As Boolean
Public DicoJSON As New Scripting.Dictionary         ' Pour la correspondance terme JSON - Colonne Excel
Public Const Tableau_Marques As String = "Peugeot|Citro�n|DS"
'
' Fin de d�claration des constantes
'

'
' Fonction pour lire les donn�es depuis un fichier JSON et les �crire dans le tableau
'
Sub MYP_JSON_Decode()
    Dim jsonText As String
    Dim jsonObject As Object, item As Object, item_item As Object, item_error As Object, item_item_error As Object
    'JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True
    ' En fait cette option pose des probl�mes avec les nombres d�cimaux qui ont plus que 15 chiffres en tout...
    
    Dim ws_Trajet As Worksheet, ws_Accueil As Worksheet
    Set ws_Trajet = Worksheets("Trajets")
    Set ws_Accueil = Worksheets("Accueil")
    
    Dim i, j As Long     ' Variables utilis�es pour les compteurs de boucles For
    Dim Trouve As Boolean
    
    Dim fso As New FileSystemObject
    Dim JsonTS As TextStream
    Dim FichierMYP As Variant   ' nom du fichier myp
    
    Dim nbTrip, nbTripReel As Integer ' nombre de trajets
    Dim CheminFichier As String
    Dim MaDate_UNIX_UTC_dep As Long, MaDate_DST_dep As Date  ' Pour convertir la date unix de d�part en date excel
    Dim MaDate_UNIX_UTC_arr As Long, MaDate_DST_arr As Date  ' Pour convertir la date unix d'arriv�e en date excel
    Dim duree_trajet As Long, duree_trajet_bis As Date
    Dim distance_trajet, conso_trajet As Double
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
    Dim Infos_VIN(), Infos_VIN_Excel() As Variant
    Dim R As Range
    Dim Info_Voiture As String
    Dim Nb_VIN_renseignes, Nb_VIN_Selection, Nb_VIN_Filtre As Integer
    Dim i_TCD, Derniere_Lat, Derniere_Long, Derniere_Adr, Critere(G_Nb_VIN_Max)
    ' Variable pour la marque de la voiture
    Dim Marque_Voiture, Marque_Voiture_Apriori As String    ' 3 valeurs possibles : Peugeot si extension de fichier = .myp
                                                            '                       Citro�n si extension de fichier = .myc
                                                            '                       DS      si extension de fichier = .myd
' V2.0 : Export JSON => passage des colonnes en dictionaire
    ' Variables pour r�cup�rer les tableaux de code d'erreur pour chaque trajet
    Dim chaine_tmp As String
    Dim nb_item_erreur As Integer, var_tmp_int As Integer
    

    ' Variable indiquant si le dictionnaire est rempli
    If Not DicoRempli Then
        Call RemplisDicoJSON
    End If
    ' On commence par �crire la valeur de la version du fichier :D
    ws_Accueil.Activate
    EcrireValeurFormat cell:=Range(CELL_ver).Offset(-1, 0), val:="Version du fichier", f_size:=10, wrap:=True
    EcrireValeurFormat cell:=Range(CELL_ver), val:=VERSION, f_size:=16, wrap:=True
    ws_Trajet.Activate
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
        Application.ScreenUpdating = True
        Exit Sub
    End If
    Select Case Mid(FichierMYP, InStrRev(FichierMYP, ".") + 1)
        Case "myp"  ' Voiture Peugeot
            Marque_Voiture_Apriori = Split(Tableau_Marques, "|")(0)
        Case "myc"  ' Voiture Citro�n
            Marque_Voiture_Apriori = Split(Tableau_Marques, "|")(1)
        Case "myd"  ' Voiture DS
            Marque_Voiture_Apriori = Split(Tableau_Marques, "|")(2)
        Case Else   ' erreur, le fichier n'a pas le bon format...
            ws_Accueil.Activate
            Cells(1, 1).Select ' On s�lection la cellule A1 ne pas avoir tout le tableau s�lectionn�
            MsgBox "Le fichier s�lectionn� n'a pas le bonne extension !" & vbLf & _
                   "Merci de changer l'extension du fichier si vous �tes sur de sa validit� pour l'importation des donn�es.", vbCritical
            Exit Sub
    End Select
    
    FormeEncours.Show
    FormeEncours.TexteEnCours = "Chargement du fichier en cours..."
    FormeEncours.Repaint
    
    Set JsonTS = fso.OpenTextFile(FichierMYP, ForReading)
    jsonText = JsonTS.ReadAll
    JsonTS.Close
    FormeEncours.TexteEnCours = "Fichier charg�, parsing des donn�es ..."
    FormeEncours.Repaint
    nbTrip = 0    ' On r�initialise le nombre de trajets
    EcrireValeurFormat cell:=Range(CELL_fichierMYP), val:=FichierMYP, wrap:=False
    Set jsonObject = JsonConverter.ParseJson(jsonText)
    
    ws_Accueil.Activate
   
    ' Il faut tester si la premi�re ligne sous l'ent�te "VIN - Description v�hicule" est vide ou pas
    ' S'il n'y a pas de valeur, alors il n'y a aucun VIN de renseign�s,
    ' Sinon on chercher � prendre tous les VIN �crits.
    If Cells(L_entete_ListeVIN + 1, C_entete_ListeVIN).Value = "" Then
        Set R = Range(Cells(L_entete_ListeVIN, C_entete_ListeVIN), Cells(L_entete_ListeVIN, C_entete_ListeVIN + 2))
    Else
        Set R = Range(Cells(L_entete_ListeVIN, C_entete_ListeVIN), Cells(L_entete_ListeVIN, C_entete_ListeVIN + 2).End(xlDown))
    End If
    Infos_VIN_Excel = R
    Set R = Nothing
    Nb_VIN_renseignes = UBound(Infos_VIN_Excel)
' V 2.0 : remplissage du tableau des infos VIN
    For i = 2 To Nb_VIN_renseignes
        ReDim Preserve Infos_VIN(3, i - 1)
        For j = 1 To 3    ' 3 colonnes : VIN, description, Marque
            Infos_VIN(j, i - 1) = Infos_VIN_Excel(i, j)
        Next j
    Next i
    If Nb_VIN_renseignes = 1 Then
        Nb_VIN_renseignes = 0
    Else
        Nb_VIN_renseignes = i - 2
    End If
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
        Unload FormeEncours
        Application.ScreenUpdating = True
        Exit Sub
    ElseIf Nb_VIN >= G_Nb_VIN_Max Then      ' Par d�faut, on g�re 20 VINs diff�rents
        MsgBox "Trop de VIN d�tect�s. Pb de structure ?", vbCritical
        Unload FormeEncours
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
            For i = 1 To Nb_VIN_renseignes
                If Infos_VIN(1, i) = Liste_VIN(l_Tab_Vin, 1) Then
                    Info_Voiture = Infos_VIN(2, i)
                    Marque_Voiture = Infos_VIN(3, i)
                    Exit For
                End If
            Next i
            If Info_Voiture = "" Then    ' VIN n'existant pas, on demande les infos pour ajouter ses infos
                Load FormeInfoVIN
                FormeInfoVIN.ListeMarque.Clear
                For i = 1 To Nb_marques
                    FormeInfoVIN.ListeMarque.AddItem Split(Tableau_Marques, "|")(i - 1)
                Next i
                FormeInfoVIN.ListeMarque.Value = Marque_Voiture_Apriori
                FormeInfoVIN.NumVIN = "VIN : " & Liste_VIN(l_Tab_Vin, 1)
                Nb_VIN_renseignes = Nb_VIN_renseignes + 1
                FormeInfoVIN.DescriptionVIN = "Ma voiture " & Nb_VIN_renseignes
                FormeInfoVIN.Show
                Info_Voiture = FormeInfoVIN.DescriptionVIN
                Marque_Voiture = FormeInfoVIN.ListeMarque.Value
                FormeInfoVIN.Hide
                Unload FormeInfoVIN
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
                ReDim Preserve Infos_VIN(3, Nb_VIN_renseignes)
                Infos_VIN(1, Nb_VIN_renseignes) = Liste_VIN(l_Tab_Vin, 1)
                Infos_VIN(2, Nb_VIN_renseignes) = Info_Voiture
                Infos_VIN(3, Nb_VIN_renseignes) = Marque_Voiture
            End If
            FormeVIN.FormeVIN_ListeVIN.AddItem (Liste_VIN(l_Tab_Vin, 1) & " - " & Info_Voiture & " - " & Marque_Voiture)
        Next l_Tab_Vin
        ' activation de la forme de choix des VIN
        FormeVIN.Show
        ' Si bouton annuler = on quitte la proc�dure
        If FormeVIN.BoutonChoisi.Value = 2 Then
            MsgBox "Vous avez annul�. On quitte !", vbCritical
            FormeEncours.Hide
            Unload FormeEncours
            Unload FormeVIN
            Application.ScreenUpdating = True
            Exit Sub
        End If
        ' On parcourt la liste pour r�cup�rer les VIN s�lectionn�s
        For l_Tab_Vin = 0 To FormeVIN.FormeVIN_ListeVIN.ListCount - 1
            If FormeVIN.FormeVIN_ListeVIN.Selected(l_Tab_Vin) Then
                Liste_VIN(l_Tab_Vin + 1, 2) = True
            End If
        Next
        Unload FormeVIN
    Else  ' cas 1 seul VIN pr�sent dans le fichier
        Liste_VIN(1, 2) = True
        ' V1.9 - Recherche de l'info v�hicule associ�e au VIN
        Info_Voiture = ""
        For i = 1 To Nb_VIN_renseignes
            If Infos_VIN(1, i) = Liste_VIN(1, 1) Then
                Info_Voiture = Infos_VIN(2, i)
                Marque_Voiture = Infos_VIN(3, i)
                Exit For
            End If
        Next i
        If Info_Voiture = "" Then    ' VIN n'existant pas, on demande les infos pour ajouter ses infos
            Load FormeInfoVIN
            FormeInfoVIN.ListeMarque.Clear
            For i = 1 To Nb_marques
                FormeInfoVIN.ListeMarque.AddItem Split(Tableau_Marques, "|")(i - 1)
            Next i
            FormeInfoVIN.ListeMarque.Value = Marque_Voiture_Apriori
            FormeInfoVIN.NumVIN = "VIN : " & Liste_VIN(1, 1)
            Nb_VIN_renseignes = Nb_VIN_renseignes + 1
            FormeInfoVIN.DescriptionVIN = "Ma voiture " & Nb_VIN_renseignes
            FormeInfoVIN.Show
            Info_Voiture = FormeInfoVIN.DescriptionVIN
            Marque_Voiture = FormeInfoVIN.ListeMarque.Value
            FormeInfoVIN.Hide
            Unload FormeInfoVIN
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
            ReDim Preserve Infos_VIN(3, Nb_VIN_renseignes)
            Infos_VIN(1, Nb_VIN_renseignes) = Liste_VIN(1, 1)
            Infos_VIN(2, Nb_VIN_renseignes) = Info_Voiture
            Infos_VIN(3, Nb_VIN_renseignes) = Marque_Voiture
        End If
    End If
    ws_Trajet.Activate
    FormeEncours.TexteEnCours = "Traitement des donn�es en cours, patience..."
    FormeEncours.Repaint
' V1.6 : stockage dans un tableau interne (que l'on vide d'abord) de tous les trajets d�j� dans l'Excel
    For i = 1 To G_Nb_Trajets_Max
        T_Trajets_existants(i) = ""
    Next i
    ' Pour vraiment avoir la derni�re ligne du tableau rempli sans risquer d'effacer une donn�e, il faut r�initialiser les filtres
    ' Sinon la derni�re ligne remplie ne sera que celle affich�e, toutes celles masqu�es au-dessous ne compteront pas...
'V2.B18 : On d�masque les colonnes masqu�es
    If var_DEBUG Then
        ws_Trajet.Range(COLs_Masquees).EntireColumn.Hidden = False
    Else
        ws_Trajet.Range(COLs_Masquees).EntireColumn.Hidden = True
    End If
    Range("$A$3:$B$150000").AutoFilter Field:=1
    Derniere_Ligne_remplie = Cells(Columns(1).Cells.Count, 1).End(xlUp).Row
    Nb_Trajets = 0
    For i = L_Premiere_Valeur To Derniere_Ligne_remplie
        T_Trajets_existants(i - L_Premiere_Valeur + 1) = Cells(i, C_vin) & ";" & Cells(i, DicoJSON("id"))
        Nb_Trajets = Nb_Trajets + 1
    Next i
    
    ' On met le format texte sur toutes les colonnes qui vont contenir des nombres d�cimaux avec beaucoup de chiffre...
    ' Sinon les valeurs seront arrondie/tronqu�es � 15 chiffres en tout...
    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=DicoJSON("startPosLatitude"), n_format:="txt"
    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=DicoJSON("endPosLatitude"), n_format:="txt"
    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=DicoJSON("startPosLongitude"), n_format:="txt"
    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=DicoJSON("endPosLongitude"), n_format:="txt"
    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=DicoJSON("destLatitude"), n_format:="txt"
    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=DicoJSON("destLongitude"), n_format:="txt"
' V1.7 : mise en place du formatage par colonne
' Colonne VIN
    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_vin, f_size:=8
' Colonne Date d�part
    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_startDate, n_format:="date"
' Colonne Date Arriv�e
    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_endDate, n_format:="date"
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
    
    i = L_Premiere_Valeur + Nb_Trajets   ' On d�fini un compteur qui sert � se positionner sur la ligne o� les donn�es doivent �tre �crites.
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
' V 2.0 : on recherche la marque de la voiture correspondant � ce VIN
            For j = 1 To Nb_VIN_renseignes
                If Infos_VIN(1, j) = VIN_Actuel Then
                    Marque_Voiture = Infos_VIN(3, j)
                End If
            Next j
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
                    Cells(i, C_startDate).Value = MaDate_DST_dep
                    Cells(i, C_endDate).Value = MaDate_DST_arr
                    ' Calcul de la dur�e du trajet en cours
                    duree_trajet = MaDate_UNIX_UTC_arr - MaDate_UNIX_UTC_dep
                    duree_trajet_bis = Date_UNIX_To_Date_UTC(MaDate_UNIX_UTC_arr) - Date_UNIX_To_Date_UTC(MaDate_UNIX_UTC_dep)
                    ' V1.9 : en cas de souci entre date d�part et date arriv�e, et donc si duree_trajet est < 0, on met 1 seconde par d�faut
                    If duree_trajet < 0 Then
                        duree_trajet = 1
                        duree_trajet_bis = "00:00:01"
                    End If
                    Cells(i, C_duree).Value = duree_trajet_bis
' V2.0 : on prend la donn�e de distance si elle existe, sinon on la calcule
                    distance_trajet = item_item("distance")
                    If distance_trajet = 0 Then
                        distance_trajet = item_item("endMileage") - item_item("startMileage")
                    Else ' on a une valeur
                        distance_trajet = CDbl(Replace(item_item("distance"), ".", ","))
                    End If
                    Cells(i, C_dist).Value = distance_trajet
                    Cells(i, DicoJSON("endMileage")).Value = item_item("endMileage")
                    
                    Cells(i, DicoJSON("consumption")).Value = CDbl(Replace(item_item("consumption"), ".", ","))
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
' V2.0 : on met une distance calcul�e si n�cessaire
                    If item_item("distance") = 0 Then
                        Cells(i, DicoJSON("distance")).Value = distance_trajet
                    End If
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
                    ' Marque de la voiture
                    Cells(i, C_marque).Value = Marque_Voiture
                    Cells(i, DicoJSON("startPosAltitude")).Value = item_item("startPosAltitude")
                    Cells(i, DicoJSON("endPosAltitude")).Value = item_item("endPosAltitude")
                    
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
                End If
            Next
        End If
    Next
    FormeEncours.TexteEnCours = "Tri des trajets en cours"
    FormeEncours.Repaint
    
' V1.6 : tri final sur colonnes A puis B
    Range(CELL_plage_1ereligne).Select
    Range(Selection, Selection.End(xlDown)).Select
    ws_Trajet.Sort.SortFields.Clear
    ws_Trajet.Sort.SortFields.Add Key:=Range(CELL_plage_max_COL_vin), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal     ' Range de base "A5:A65536"
    ws_Trajet.Sort.SortFields.Add Key:=Range(CELL_plage_max_COL_id), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal      ' Range de base "B5:B65536"
    With ws_Trajet.Sort
        .SetRange Range(CELL_plage_max_Donnees)        ' range par d�faut "A5:AJ65536"
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Cells(1, 4).Select   ' On s�lection la cellule de version pour ne pas avoir tout le tableau s�lectionn�
    
' V2.beta15 : d�placement du formatage avant le remplissage
'' V1.7 : mise en place du formatage par colonne
'' Colonne VIN
'    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_vin, f_size:=8
'' Colonne Date d�part
'    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_startDate, n_format:="date"
'' Colonne Date Arriv�e
'    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_endDate, n_format:="date"
'' Colonne Dur�e
'    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_duree, n_format:="duree"
'' Colonne Distance
'    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_dist, n_format:="1"
'' Colonne Distance totale
'    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=DicoJSON("endMileage"), n_format:="1"
'' Colonne conso
'    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=DicoJSON("consumption"), n_format:="2"
'' Colonne conso moyenne
'    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=C_conso_moy, n_format:="1"
'' Colonne niveau carburant
'    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=DicoJSON("fuelLevel"), n_format:="%"
'' Colonne adresse D�part
'    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=DicoJSON("startPosAddress"), n_format:="add"
'' Colonne adresse Arriv�e
'    Formater_Cellules ws_tmp:=ws_Trajet, ligne_cell:=L_Premiere_Valeur, colonne_cell:=DicoJSON("endPosAddress"), n_format:="add"
    
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
    ws_Accueil.Activate
    Cells(1, 1).Select ' On s�lection la cellule A1 ne pas avoir tout le tableau s�lectionn�
' V1.7 : optimisation : remise de la mise � jour de l'affichage
    Application.ScreenUpdating = True
    
End Sub     'Fin de la proc�dure d'importation des donn�es JSON
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
    '                                                           <- Valeurs = "txt" pour du texte (utilis� pour les coordonn�es)
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
        Case "txt", "add"
            ws_tmp.Range(ws_tmp.Cells(ligne_cell, colonne_cell), ws_tmp.Cells(ligne_cell, colonne_cell).End(xlDown)).NumberFormat = "@"
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
    ' On commence par �crire la valeur de la version du fichier :D
    Worksheets("Accueil").Activate
    EcrireValeurFormat cell:=Range(CELL_ver).Offset(-1, 0), val:="Version du fichier", f_size:=10, wrap:=True
    EcrireValeurFormat cell:=Range(CELL_ver), val:=VERSION, f_size:=16, wrap:=True
' 1.9 : il faut d'abord retirer le filtre sur le VIN pour TOUT effacer
    Worksheets("Trajets").Range("$A$3:$B$150000").AutoFilter Field:=1
    With Worksheets("Trajets")
        .Range(CELL_fichierMYP) = ""           ' Le nom du fichier
        .Range(CELL_plage_donnees) = ""    ' Le grand tableau de valeurs de tous les trajets effectu�s
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
Private Sub AfficheMasque_Colonnes_masquees()
' Pour afficher ou masquer les colonnes masqu�es
    If Worksheets("Trajets").Range(COLs_Masquees).EntireColumn.Hidden = False Then
        Worksheets("Trajets").Range(COLs_Masquees).EntireColumn.Hidden = True
    Else
        Worksheets("Trajets").Range(COLs_Masquees).EntireColumn.Hidden = False
    End If
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
    
    adresse = Replace(adresse, vbCrLf, ", ")   ' On remplace tous les retours � la ligne "\n" par des ", "
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
    adresse = Replace(adresse, ", ", "\n")   ' On remplace tous les retours � la ligne ", " par des "\n"
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
Sub Bouton_Tuto_Import_Cliquer()
' Activer la feuille Tutoriel d'Importation
    Sheets("Tuto-Import").Activate
End Sub
Sub Bouton_Tuto_Export_Cliquer()
' Activer la feuille Tutoriel d'Exportation
    Sheets("Tuto-Export").Activate
End Sub
Sub RemplisDicoJSON()
    DicoJSON.Add "id", 2                    ' Identificateur trajet
    DicoJSON.Add "endMileage", 7
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
    DicoJSON.Add "startMileage", 18         ' Donn�e suppl�mentaire � ajouter dans la partie masqu�e du tableau
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
    ' Attention la colonne 33 est r�serv�e pour la marque.
    DicoJSON.Add "startPosAltitude", 34 ' Valeur pr�sente dans certains fichier de donn�es
    DicoJSON.Add "endPosAltitude", 35   ' Valeur pr�sente dans certains fichier de donn�es
    DicoRempli = True
End Sub

Sub TEST_PROC()
    Dim ws_Trajet As Worksheet, ws_Accueil As Worksheet
    Set ws_Trajet = Worksheets("Trajets")
    Set ws_Accueil = Worksheets("Accueil")
    Dim tmp As Variant
    Dim i As Integer
    
    tmp = Split(ws_Trajet.Range(CELL_fichierMYP).Value, "\")
    Left(tmp(UBound(tmp)), Len(tmp(UBound(tmp))) - 4)
    
    i = Len(tmp(UBound(tmp)))
    
    MsgBox "Nom du fichier ouvert : " & Left(tmp(UBound(tmp)), Len(tmp(UBound(tmp))) - 4), vbOKOnly + vbInformation
    
End Sub



'
' Fonction pour �crire les donn�es trajets dans un fichier JSON : V 2.0
'
Sub MYP_JSON_Encode()
    Dim ws_Trajet As Worksheet, ws_Accueil As Worksheet
    Set ws_Trajet = Worksheets("Trajets")
    Set ws_Accueil = Worksheets("Accueil")
    
    Dim R As Range
    Dim Tableau_Trajets(), Infos_VIN(), Tableau_VIN_Filtre(), Liste_VIN() As Variant
    Dim Nb_VIN_renseignes, Nb_VIN_Filtre, Nb_VIN_Selection, Nb_Alerts As Integer
    Dim i, j, k, Nb_Trajets_ecrits, Nb_Trajets_ecrits_boucle As Long
    Dim Boucle_Filtre, Marque, NumFichier, InfoJSON, Valeur
    Dim Dico_Marque, Dico_VIN As Object
    Dim CheminFichier, Marque_Voiture, NomFichier, Extension, Info_Voiture As String
    Dim Nom_Fichier_Defaut As Variant      ' Pour la proposition du nom du fichier � �crire
    Dim l_Tab_Vin, l_tab_Marque As Integer                    ' pour boucle sur les vin ou les marques
    Dim fso As New FileSystemObject
    Dim Fini As Boolean
    Dim FichierMYP As Variant   ' nom du fichier myp
    Dim VIN_Actuel As String                    ' VIN associ� � la boucle des trajets
    Dim Reponse As Integer
    Dim virgule_O_N As String   ' Variable qui contiendra ou non une virgule, en fonction de si on est sur le dernier item ou pas...
    ' On commence par �crire la valeur de la version du fichier :D
    ws_Accueil.Activate
    EcrireValeurFormat cell:=Range(CELL_ver).Offset(-1, 0), val:="Version du fichier", f_size:=10, wrap:=True
    EcrireValeurFormat cell:=Range(CELL_ver), val:=VERSION, f_size:=16, wrap:=True
    
'V2.B18 : On d�masque les colonnes masqu�es
    If var_DEBUG Then
        ws_Trajet.Range(COLs_Masquees).EntireColumn.Hidden = False
    Else
        ws_Trajet.Range(COLs_Masquees).EntireColumn.Hidden = True
    End If
    
' D�but macro
    ' Variable indiquant si le dictionnaire est rempli
    If Not DicoRempli Then
        Call RemplisDicoJSON
    End If
    Application.ScreenUpdating = False
    FormeEncours.Show
    FormeEncours.TexteEnCours = "Recherche des informations v�hicules existantes ..."
    FormeEncours.Repaint
    
' 1�re �tape : mettre dans un tableau toutes les informations v�hicules existantes
    If Cells(L_entete_ListeVIN + 1, C_entete_ListeVIN).Value = "" Then
        MsgBox "Aucune donn�e v�hicule existante, pb de structure ou de donn�es dans l'Excel. Abort", vbCritical
        FormeEncours.Hide
        Exit Sub
    Else
' Ajouter ici le tri des trajets !!!!!! A FAIRE
        Set R = Range(Cells(L_entete_ListeVIN, C_entete_ListeVIN), Cells(L_entete_ListeVIN, C_entete_ListeVIN + 2).End(xlDown))
    End If
    Infos_VIN = R
    Nb_VIN_renseignes = UBound(Infos_VIN)

' 2�me �tape : sauvegarder les valeurs du filtre actuel des trajets
    ws_Trajet.Activate
    If Not (ActiveSheet.AutoFilterMode) Then
        Nb_VIN_Filtre = 0
    ElseIf Not (ActiveSheet.AutoFilter.Filters(1).On) Then
        Nb_VIN_Filtre = 0
    Else
        Nb_VIN_Filtre = ActiveSheet.AutoFilter.Filters(1).Count
    End If
    Select Case Nb_VIN_Filtre
        Case 0
        Case 1
            ReDim Tableau_VIN_Filtre(1)
            Tableau_VIN_Filtre(1) = Right(ActiveSheet.AutoFilter.Filters(1).Criteria1, Len(ActiveSheet.AutoFilter.Filters(1).Criteria1) - 1)
        Case 2
            ReDim Tableau_VIN_Filtre(2)
            Tableau_VIN_Filtre(1) = Right(ActiveSheet.AutoFilter.Filters(1).Criteria1, Len(ActiveSheet.AutoFilter.Filters(1).Criteria1) - 1)
            Tableau_VIN_Filtre(2) = Right(ActiveSheet.AutoFilter.Filters(1).Criteria2, Len(ActiveSheet.AutoFilter.Filters(1).Criteria2) - 1)
        Case Else
            Tableau_VIN_Filtre = ActiveSheet.AutoFilter.Filters(1).Criteria1
    End Select

' 3�me �tape : retrait de tous les filtres
    Range("$A$3:$B$150000").AutoFilter Field:=1

' 4�me �tape : mise en m�moire des donn�es trajets
    FormeEncours.TexteEnCours = "Recherche des trajets existants ..."
    FormeEncours.Repaint
    Set R = Range(Cells(L_Premiere_Valeur, 1), Cells(L_Premiere_Valeur, DicoJSON("endPosAltitude")).End(xlDown))
    Tableau_Trajets = R
    Set R = Nothing     ' lib�ration m�moire
    
' 5�me �tape : demander pour quelle marque on veut exporter en fonction des marques pr�sentes dans les trajets
' Evidemment, si une seule marque, on force celle-l� !
    FormeEncours.TexteEnCours = "Choix de la marque � traiter ..."
    FormeEncours.Repaint
    Load FormeMarque
    FormeMarque.ListeMarque.Clear
    Set Dico_Marque = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(Tableau_Trajets)
        If (Not Dico_Marque.Exists(Tableau_Trajets(i, C_marque))) And (Len(Tableau_Trajets(i, C_marque)) > 1) Then
            Dico_Marque.Add Tableau_Trajets(i, C_marque), Tableau_Trajets(i, C_marque)
            FormeMarque.ListeMarque.AddItem Tableau_Trajets(i, C_marque)
        End If
    Next
    If Dico_Marque.Count > 1 Then
        FormeMarque.Show
        ' Si bouton annuler = on quitte la proc�dure
        If FormeMarque.BoutonChoisi.Value = 2 Then
            MsgBox "Vous avez annul�. On quitte !", vbCritical
            FormeEncours.Hide
            Unload FormeMarque
            Set Dico_Marque = Nothing
            Exit Sub
        End If
        ' On parcourt la liste pour r�cup�rer la marque s�lectionn�e
        For l_tab_Marque = 0 To FormeMarque.ListeMarque.ListCount - 1
            If FormeMarque.ListeMarque.Selected(l_tab_Marque) Then
                Marque_Voiture = FormeMarque.ListeMarque.List(l_tab_Marque)
            End If
        Next
        FormeMarque.Hide
    Else
        Marque_Voiture = Tableau_Trajets(L_Premiere_Valeur, C_marque)
    End If
    Unload FormeMarque
    Set Dico_Marque = Nothing

' 6�me �tape : demander les VIN de cette marque pour lesquelles ont veut exporter en fonction des VIN pr�sents
    FormeEncours.TexteEnCours = "Choix de(s) VIN(s) � exporter ..."
    FormeEncours.Repaint
    Load FormeVINExport
    FormeVINExport.ListeVIN.Clear
    Set Dico_VIN = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(Tableau_Trajets)
        If (Not Dico_VIN.Exists(Tableau_Trajets(i, C_vin))) And (Tableau_Trajets(i, C_marque) = Marque_Voiture) Then
            Dico_VIN.Add Tableau_Trajets(i, C_vin), Tableau_Trajets(i, C_vin)
            Info_Voiture = ""
            For j = 1 To UBound(Infos_VIN)
                If Infos_VIN(j, 1) = Tableau_Trajets(i, C_vin) Then
                    Info_Voiture = Infos_VIN(j, 2)
                End If
            Next j
            FormeVINExport.ListeVIN.AddItem Tableau_Trajets(i, C_vin) & " - " & Info_Voiture
        End If
    Next
    If Dico_VIN.Count > 1 Then
        FormeVINExport.Show
        ' Si bouton annuler = on quitte la proc�dure
        If FormeVINExport.BoutonChoisi.Value = 2 Then
            MsgBox "Vous avez annul�. On quitte !", vbCritical
            FormeEncours.Hide
            Unload FormeVINExport
            Set Dico_VIN = Nothing
            Exit Sub
        End If
        ' On parcourt la liste pour r�cup�rer les VIN s�lectionn�s
        Nb_VIN_Selection = 0
        For l_Tab_Vin = 0 To FormeVINExport.ListeVIN.ListCount - 1
            If FormeVINExport.ListeVIN.Selected(l_Tab_Vin) Then
                Nb_VIN_Selection = Nb_VIN_Selection + 1
                ReDim Preserve Liste_VIN(Nb_VIN_Selection)
                Liste_VIN(Nb_VIN_Selection) = Split(FormeVINExport.ListeVIN.List(l_Tab_Vin), " -")(0)
            End If
        Next
        FormeVINExport.Hide
    Else
        ReDim Liste_VIN(1)
        Nb_VIN_Selection = 1
        Liste_VIN(1) = Tableau_Trajets(L_Premiere_Valeur, C_vin)
    End If
    Unload FormeVINExport
    Set Dico_VIN = Nothing
    
' 7�me �tape : demande du nom de fichier en sortie, on impose l'extension
    Select Case Marque_Voiture
        Case Split(Tableau_Marques, "|")(0)  ' Voiture Peugeot
            Extension = ".myp"
        Case Split(Tableau_Marques, "|")(1)  ' Voiture Citro�n
            Extension = ".myc"
        Case Split(Tableau_Marques, "|")(2)  ' Voiture DS
            Extension = ".myd"
        Case Else   ' erreur, marque inconnue !
            ws_Accueil.Activate
            Cells(1, 1).Select ' On s�lection la cellule A1 ne pas avoir tout le tableau s�lectionn�
            MsgBox "La marque semble inconnue !" & vbLf & "Merci de v�rifier les donn�es des trajets dans l'Excel.", vbCritical
            FormeEncours.Hide
            Exit Sub
    End Select
    
    ' On r�cup�re le chemin du fichier excel ouvert
    CheminFichier = ActiveWorkbook.Path & "\"
    ' Il y a une erreur si on travail dans le dossier OneDrive, le CheminFicher est un lien du type https://d.docs.live.net/8f87e4...
    ' Il faut donc v�rifier si le d�but chaine de caract�re n'est pas https://
    If (InStr(1, CheminFichier, "https://", vbTextCompare) <> 1 And InStr(1, CheminFichier, "http://", vbTextCompare) <> 1) Then
        ChDir CheminFichier     ' Le chemin du fichier ne contient pas de lien, on change le dossier d'ouverture
    End If
    
    ' Ce qui suit va permettre d'obtenir un nom par d�faut pour le fichier d'exportation, libre � l'utilisateur d'accepter ce nom ou pas
    ' On va d�terminer le nom du fichier qui a �t� import� en dernier (le seul qui est �crit)
    Nom_Fichier_Defaut = Split(ws_Trajet.Range(CELL_fichierMYP).Value, "\")
    Nom_Fichier_Defaut = "EXPORT-JSON-Excel - " & Left(Nom_Fichier_Defaut(UBound(Nom_Fichier_Defaut)), Len(Nom_Fichier_Defaut(UBound(Nom_Fichier_Defaut))) - 4)
    
    Fini = False
    While Not Fini
        'NomFichier = InputBox("Indiquez le nom du fichier � cr�er (sans extension)", "Choix du nom du fichier")
        FichierMYP = Application.GetSaveAsFilename(InitialFileName:=Nom_Fichier_Defaut, fileFilter:="Fichiers trajets " & Marque_Voiture & " App " & "(*" & Extension & "), *" & Extension)
        Set fso = CreateObject("Scripting.FileSystemObject")
        If FichierMYP = False Then
            MsgBox "Aucun fichier n'a �t� selectionn� !", vbCritical
            FormeEncours.Hide
            ws_Accueil.Activate
            Cells(1, 1).Select ' On s�lection la cellule A1 ne pas avoir tout le tableau s�lectionn�
            Exit Sub
        End If
        If fso.FileExists(FichierMYP) Then
            Reponse = MsgBox("Le fichier s�lectionn� existe d�j�. Voulez-vous vraiment le remplacer ?" & vbLf & vbLf & FichierMYP, vbYesNoCancel + vbExclamation, "Attention...")
            If Reponse = vbYes Then  'Si le bouton Oui est cliqu� ...
                Fini = True
            ElseIf (Reponse = vbCancel) Then
                ' On annule tout, et on stoppe la proc�dure.
                MsgBox "La proc�dure d'exportation a �t� annul�e.", vbCritical
                FormeEncours.Hide
                ws_Accueil.Activate
                Cells(1, 1).Select ' On s�lection la cellule A1 ne pas avoir tout le tableau s�lectionn�
                Exit Sub
            End If
        Else
            Fini = True
        End If
    Wend

' 8�me �tape : �criture des donn�es
    FormeEncours.TexteEnCours = "Ecriture du fichier en sortie, patience ..."
    FormeEncours.Repaint
    NumFichier = FreeFile
    Open FichierMYP For Output As #NumFichier
    Print #NumFichier, "["
    ' Boucle sur l'ensemble des VIN s�lectionn�s
    Nb_Trajets_ecrits = 0
    For i = 1 To Nb_VIN_Selection
        VIN_Actuel = Liste_VIN(i)
        If i > 1 Then
            Print #NumFichier, "      }"    ' fin du VIN pr�c�dent
            Print #NumFichier, "    ]"    ' fin du VIN pr�c�dent
            Print #NumFichier, "  },"    ' fin du VIN pr�c�dent
        End If
        Print #NumFichier, "  {"
        Print #NumFichier, "    " & Chr(34) & "vin" & Chr(34) & ": " & Chr(34) & VIN_Actuel & Chr(34) & ","
        Print #NumFichier, "    " & Chr(34) & "trips" & Chr(34) & ": ["
        ' Boucle sur l'ensemble des trajets des VIN s�lectionn�s
        Nb_Trajets_ecrits_boucle = 0
        For j = 1 To UBound(Tableau_Trajets())
            If Tableau_Trajets(j, C_vin) = VIN_Actuel Then     ' VIN en cours
                Nb_Trajets_ecrits = Nb_Trajets_ecrits + 1
                Nb_Trajets_ecrits_boucle = Nb_Trajets_ecrits_boucle + 1
                If ((Nb_Trajets_ecrits Mod 100) = 0) Then
                    FormeEncours.TexteEnCours = "Traitement trajets, " & Nb_Trajets_ecrits & " �crits, patience..."
                    FormeEncours.Repaint
                End If
                If Nb_Trajets_ecrits_boucle > 1 Then
                    Print #NumFichier, "      },"    ' fin du trajet pr�c�dent
                End If
                Print #NumFichier, "      {"
                k = 0 ' Compteur pour mat�rialiser l'item en cours d'�criture
                virgule_O_N = ""
                For Each InfoJSON In DicoJSON.Keys    ' �criture de toutes les cl�s JSON pour un trajet
                ' ce qu'il faut �crire doit �tre transcod�, suivant les cas
                    k = k + 1 ' On passe � l'item suivant (Rappel, initialement k=0, donc on aura k=1 juste apr�s).
                    ' Test pour savoir si on est sur le dernier item ou non
                    If (k < DicoJSON.Count) Then    ' On n'est pas encore sur le dernier item, donc on met une virgule apr�s l'item �crit
                        virgule_O_N = ","
                    Else
                        virgule_O_N = ""            ' On est sur le dernier item, et il ne faut pas mettre de virgule apr�s l'item �crit
                    End If
                
                    Select Case InfoJSON
                        Case "startPosAddress", "endPosAddress", "destAddress"
                            Print #NumFichier, "        " & Chr(34) & InfoJSON & Chr(34) & ": " & Chr(34) & Correction_Adresse_INVERSE(Tableau_Trajets(j, DicoJSON(InfoJSON))) & Chr(34) & virgule_O_N
                        Case "fuelLevel"
                            Print #NumFichier, "        " & Chr(34) & InfoJSON & Chr(34) & ": " & Round(Tableau_Trajets(j, DicoJSON(InfoJSON)) * 100) & virgule_O_N
                        Case "maintenancePassed"
                            If Tableau_Trajets(j, DicoJSON(InfoJSON)) Then
                                Print #NumFichier, "        " & Chr(34) & InfoJSON & Chr(34) & ": true" & virgule_O_N
                            Else
                                Print #NumFichier, "        " & Chr(34) & InfoJSON & Chr(34) & ": false" & virgule_O_N
                            End If
                        Case "alertsActive", "alertsResolved"
                            Print #NumFichier, "        " & Chr(34) & InfoJSON & Chr(34) & ": ["
                            Nb_Alerts = 0
                            For Each Valeur In Split(Tableau_Trajets(j, DicoJSON(InfoJSON)), ";")
                                Nb_Alerts = Nb_Alerts + 1
                                If Nb_Alerts > 1 Then
                                    Print #NumFichier, "          },"    ' fin des alertes pr�c�dentes
                                End If
                                Print #NumFichier, "          {"
                                Print #NumFichier, "            " & Chr(34) & "code" & Chr(34) & ": " & Chr(34) & Valeur & Chr(34)
                            Next
                            If Nb_Alerts > 0 Then
                                Print #NumFichier, "          }"
                            End If
                            Print #NumFichier, "        ]" & virgule_O_N
                        Case "endMileage", "consumption", "startPosLatitude", "startPosLongitude", "endPosLatitude", "endPosLongitude", "startMileage", "distance"   ' transformer , en .
                            Print #NumFichier, "        " & Chr(34) & InfoJSON & Chr(34) & ": " & Replace(Tableau_Trajets(j, DicoJSON(InfoJSON)), ",", ".") & virgule_O_N
                        Case "destLatitude", "destLongitude", "startPosAltitude", "endPosAltitude"   ' transformer , en . et champs potentiellement vides
                            If Tableau_Trajets(j, DicoJSON(InfoJSON)) <> "" Then
                                Print #NumFichier, "        " & Chr(34) & InfoJSON & Chr(34) & ": " & Replace(Tableau_Trajets(j, DicoJSON(InfoJSON)), ",", ".") & virgule_O_N
                            End If
                        Case "destQuality", "startPosQuality", "endPosQuality", "maintenanceDays"  ' champs potentiellement vides
                            ' Il faut tester si les champs sont vides ou pas. S'ils sont vides, on �crit rien, s'il y a quelque chose on l'�crit.
                            If (Tableau_Trajets(j, DicoJSON(InfoJSON)) <> "") Then
                                Print #NumFichier, "        " & Chr(34) & InfoJSON & Chr(34) & ": " & Tableau_Trajets(j, DicoJSON(InfoJSON)) & virgule_O_N
                            End If
                            
                        Case Else
                            Print #NumFichier, "        " & Chr(34) & InfoJSON & Chr(34) & ": " & Tableau_Trajets(j, DicoJSON(InfoJSON)) & virgule_O_N
                    End Select
                Next
            End If
        Next j
    Next i
    Print #NumFichier, "      }"    ' fin du VIN pr�c�dent
    Print #NumFichier, "    ]"    ' fin du VIN pr�c�dent
    Print #NumFichier, "  }"    ' fin du VIN pr�c�dent
    Print #NumFichier, "]"
'Fermeture
    Close #NumFichier
    MsgBox "Fichier " & FichierMYP & " sauvegard�", vbOKOnly + vbInformation
    
' 9�me �tape : remise en place du filtre sur les trajets
    For i = 1 To Nb_VIN_Filtre
        Select Case i
            Case 1
                ActiveSheet.Range("$A$2:$B$65000").AutoFilter Field:=1, Criteria1:=Tableau_VIN_Filtre(1)
            Case 2
                ActiveSheet.Range("$A$2:$B$65000").AutoFilter Field:=1, Criteria1:=Tableau_VIN_Filtre(1), Operator:=xlOr, Criteria2:=Tableau_VIN_Filtre(2)
            Case Else
                ActiveSheet.Range("$A$2:$B$65000").AutoFilter Field:=1, Criteria1:=Tableau_VIN_Filtre, Operator:=xlFilterValues
        End Select
    Next
    FormeEncours.Hide
    ws_Accueil.Activate
    Cells(1, 1).Select ' On s�lection la cellule A1 ne pas avoir tout le tableau s�lectionn�
    Application.ScreenUpdating = True
End Sub



