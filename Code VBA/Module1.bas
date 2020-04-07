Attribute VB_Name = "Module1"
' Licence utilis�e :
'                   GNU AFFERO GENERAL PUBLIC LICENSE
'                      Version 3, 19 November 2007
'
' D�p�t GitHub : https://github.com/MilesTEG1/Importation-Excel-Donnees-MyPeugeot/blob/master/LICENSE.md
'
' @author MilesTEG1@gmail.com
' @license  AGPL-3.0 (https://www.gnu.org/licenses/agpl-3.0.fr.html)



'
' D�inissons quelques constantes qui serviront pour les colonnes/lignes/plages de cellules.
'
Const L_premiere_valeur As Integer = 5      ' Premi�re ligne � contenir des donn�es (avant ce sont les lignes d'en-t�te
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
Const CELL_vin_entete   As String = "B1"    ' Cellule qui contiendra le VIN associ� aux donn�es de l'en-t�te
Const CELL_nb_trips     As String = "H1"    ' Cellule qui contiendra le nombre de trajets totale (tous VIN confondus)
Const CELL_fichierMYP   As String = "O1"    ' Cellule qui contiendra le nom du fichier import�
Const CELL_km           As String = "H2"    ' Cellule qui contiendra le kilom�trage actuel de la voiture
Const CELL_conso_tot    As String = "L1"    ' Cellule qui contiendra la consommation totale en L
Const CELL_conso_tot_moy    As String = "L2"    ' Cellule qui contiendra la consommation moyenne totale pour tous les trajets
Const CELL_plage_donnees    As String = "A" & L_premiere_valeur & ":Q20000" ' Plage de cellules contenant les donn�es
Const CELL_plage_conso_tot  As String = "H" & L_premiere_valeur & ":H"      ' Cellule qui contiendra le kilom�trage actuel de la voiture

Const VERSION As String = "Version du fichier" & vbLf & "v1.4"    ' Version du fchier
Const CELL_ver As String = "D1:D2"     ' Cellule o� afficher la version du fichier
'
' Fin de d�claration des constantes
'

'
' Fonction pour lire les donn�es depuis un fichier JSON et les �crire dans le tableau
'
Sub MYP_JSON_Decode()
    Dim jsonText As String
    Dim jsonObject As Object, item As Object, item_item As Object
    Dim i As Long
    Dim ws As Worksheet
    
    Dim FSO As New FileSystemObject
    Dim JsonTS As TextStream
    
    Dim FichierMYP As Variant   ' nom du fichier myp
    Dim nbTrip, nbTripReel As Integer ' nombre de trajets
    Dim kilometrage As Single   ' Pour stocker le kilom�trage total de la voiture.
    Dim conso_totale As Single  ' Pour stocker la consommation totale de tout le tableau
    
    Set ws = Worksheets("Trajets-MyPeugeot")
    EcrireValeurFormat cell:=ws.Range(CELL_ver), val:=VERSION, f_size:=10, wrap:=True
    
    JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True
    
    FichierMYP = Application.GetOpenFilename("Fichiers trajets Peugeot App (*.myp),*.myp,Fichiers trajets Citroen App (*.myc),*.myc,Fichiers trajets DS App (*.myd),*.myd")  ' On demande la s�lection du fichier
    If FichierMYP = False Then
        MsgBox "Aucun fichier n'a �t� selectionn� !"
        Exit Sub
    Else
        Set JsonTS = FSO.OpenTextFile(FichierMYP, ForReading)
        jsonText = JsonTS.ReadAll
        JsonTS.Close
    
        ' Comme le fichier existe, on efface tout
        Effacage_Donnees
        
        nbTrip = 0    ' On r�initialise le nombre de trajets
    
        EcrireValeurFormat cell:=ws.Range(CELL_fichierMYP), val:=FichierMYP, wrap:=True
        
    End If
    
  
    Set jsonObject = JsonConverter.ParseJson(jsonText)
    
        
    i = L_premiere_valeur   ' On d�fini un compteur qui sert � se positionner sur la ligne o� les donn�es doivent �tre �crites.
                            ' Le n� de la ligne o� on va commencer � �crire les donn�es
                            ' ws.Cells(LIGNE, COLONNE)     O� LIGNE commence � 1
                            '                              O� COLONNE commence � 1 = A
                            ' On va utiliser les constantes d�finie avant la fonction pour �crire/formater/effacer une cellule
                            
                            
    For Each item In jsonObject
        
        For Each item_item In item("trips")   ' Boucle pour r�cup�rer les trajets
            
            Dim MaDate_UNIX_GMT_dep As Long, MaDate_DST_dep As Date  ' Pour convertir la date unix de d�part en date excel
            Dim MaDate_UNIX_GMT_arr As Long, MaDate_DST_arr As Date  ' Pour convertir la date unix d'arriv�e en date excel
            Dim duree_trajet As Long, duree_trajet_bis As Date
            Dim distance_trajet As Single, conso_trajet As Single, niveau_carb As Single
                                              
'            With ws.Cells(i, C_vin)
'                .Value = item("vin")   ' On �crit le VIN r�cup�r�
'                .Font.Size = 8
'                .NumberFormat = "General"
'            End With
            
            EcrireValeurFormat cell:=ws.Cells(i, C_vin), val:=item("vin"), f_size:=8    ' On �crit le VIN r�cup�r�
            EcrireValeurFormat cell:=ws.Cells(i, C_id), val:=item_item("id")            ' On �crit l'ID r�cup�r�
            
            ' R�cup�ration des dates
            ' On stocke les deux dates (d�part et arriv�e) car il faut d�terminer le temps de parcours
            ' qui ne doit pas �tre d�pendant d'un �ventuelle changement d'heure en cours de route
            MaDate_UNIX_GMT_dep = item_item("startDateTime")            ' Date de d�part
            MaDate_UNIX_GMT_arr = item_item("endDateTime")              ' Date d'arriv�e
            MaDate_DST_dep = Date_UNIX_To_Date_DST(MaDate_UNIX_GMT_dep)
            MaDate_DST_arr = Date_UNIX_To_Date_DST(MaDate_UNIX_GMT_arr)
            
            EcrireValeurFormat cell:=ws.Cells(i, C_date_dep), val:=MaDate_DST_dep, n_format:="date"
            EcrireValeurFormat cell:=ws.Cells(i, C_date_arr), val:=MaDate_DST_arr, n_format:="date"
            
            ' Calcul de la dur�e du trajet en cours
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
            
            ' Pour le calcul de la consommation moyenne, il faut �viter la division par z�ro dans le cas o�
            ' la voiture � tourner � l'arret, la distance parcourue est nulle
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

            kilometrage = item_item("endMileage")   ' On stocke le kilom�trage de fin de trajet pour �tre affich� en tant que kilom�trage actuel lors du dernier trajet
            
        Next
    

            
    Next
    
    EcrireValeurFormat cell:=ws.Range(CELL_nb_trips), val:=nbTrip, f_size:=12                   ' On �crit le nombre de trajet
    EcrireValeurFormat cell:=ws.Range(CELL_km), val:=kilometrage, n_format:="0", f_size:=12     ' On �crit le kilom�trage total de la voiture

    
    ' Calcul de la consommation totale du tableau
    Dim j As Integer
    j = 4 + nbTrip      ' valeur qui d�limite la derni�re cellule de la colonne G contenant une consommation
    'Set MaPlage = ws.Range("H5:H" & j)
    'conso_totale = ws.Application.WorksheetFunction.Sum(ws.Range("H5:H" & j))
    
    ' On �crit la valeur calcul�e de la consommation totale de tous les trajets
    EcrireValeurFormat cell:=ws.Range(CELL_conso_tot), val:=ws.Application.WorksheetFunction.Sum(ws.Range("H5:H" & j)), n_format:="2", f_size:=12
    
    'ws.Range(CELL_conso_tot) = conso_totale   ' On a maintenant la consommation totale de tous les trajets
    'ws.Range(CELL_conso_tot).NumberFormatLocal = "0,00"
    
    ws.Range(CELL_conso_tot_moy).NumberFormatLocal = "0,0"

End Sub

Private Sub EcrireValeurFormat(cell As Variant, val As Variant, Optional n_format As String = "General", Optional f_size As Integer = 10, Optional wrap As Boolean = False)
    ' Fonction pour �crire une valeur dans une cellule
    ' Arguments obligatoires :  cellule As Variant  <- La cellule ou plage de cellule devant �tre modifi�e
    '                           val As Variant      <- La valeur � �crire dans la cellule/plage
    ' Arguments optionels :     n_format As String = ""         <- Le format NumberFormat, d�faut = "Genral"
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
    ' Fonction pour effacer les donn�es avant d'utiliser celles du fichier ouvert
    With Worksheets("Trajets-MyPeugeot")
        .Range(CELL_vin_entete) = ""           ' Le VIN
        .Range(CELL_fichierMYP) = ""           ' Le nom du fichier
        .Range(CELL_nb_trips) = ""  ' Le nombre de trajet et le kilom�trage total
        .Range(CELL_km) = ""
        .Range(CELL_plage_donnees) = ""    ' Le grand tableau de valeurs de tous les trajets effectu�s
        .Range(CELL_conso_tot) = ""           ' La consommation totale sur tous les trajets
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


Sub test()

    
    Dim MaDate_UNIX_GMT As Long, MaDate_GMT As Date, MaDate_DST As Date  ' Pour convertir une date unix en date excel

    
    ' pour les essais = 25/10/2020 � 0h50min00s  ==  1603587000 UTC-UNIX <- c'est cens� �tre encore l'heure d'�t�
    ' pour les essais = 29/03/2020 � 0h50min00s  ==  1585443000 UTC-UNIX <- c'est cens� �tre encore l'heure d'hiver
    MaDate_UNIX_GMT = 1603590600
    'MaDate_GMT = (MaDate_UNIX_GMT / 86400) + DateValue("01/01/1970")

    MaDate_DST = Date_UNIX_To_Date_DST(MaDate_UNIX_GMT)
    
    MsgBox "Date UNIX GMT = " & MaDate_GMT & vbCrLf & "Date DST = " & MaDate_DST

End Sub


















