# Importation-Excel-Donnees-MyPeugeot

[:arrow_right: To-do Liste](https://github.com/MilesTEG1/Importation-Excel-Donnees-MyPeugeot/blob/master/ReadME.md#to-do-priorit%C3%A9)

## Objectif

Il s'agit d'un fichier excel avec macros (XLSM) qui permet de récupérer les données des applications MyPeugeot, MyCitroën et MyDS et d'en faire des statistiques.

## Idée

Mon projet est une variante assez proche (j'ai piqué plein d'idées) à celui-ci :
[Trajets myp de MYPEUGEOT APP sous Excel toutes versions](https://www.forum-peugeot.com/Forum/threads/trajets-myp-de-mypeugeot-app-sous-excel-toutes-versions.9456/)  
J'ai commencé [à en parler ici](https://www.forum-peugeot.com/Forum/threads/fichier-excel-macros-pour-r%C3%A9cup%C3%A9rer-les-trajets-de-lapplication-mypeugeot.119785/).  

Il faut exporter ses données MyPeugeot en utilisant cet option :  
![Option à utiliser](https://raw.githubusercontent.com/MilesTEG1/Importation-Excel-Donnees-MyPeugeot/master/images/Option%20pour%20exporter%20les%20trajets%20dans%20l'app%20MyPeugeot.png)  
Ça va envoyer un email avec un fichier .myp.

## Pourquoi refaire un fichier qui fonctionne ?

Et bien, parce que pour moi il ne fonctionne pas. Les données du fichier .myp exportés pour chaque trajets ne sont pas dans le même ordre que celui pour lequel a été conçu le fichier original. Je ne sais pas pourquoi ce n'est pas le même ordre, mais toujours est-il que ça rend l'exploitation impossible. Mais comme le fichier de données .myp est dans un format JSON, il est possible de faire autrement qu'avec une structure figée, il suffit de "parser" les données pour les récupérer.  

## Outils nécessaires

Pour celà je me suis aidé de cette bibliothèque de fonctions : [VBA-JSON-2.3.1](https://github.com/VBA-tools/VBA-JSON)  

Le fichier utilisé est normalement déjà inclus dans le fichier excel.
Cependant il faut activer une référence : **Microsoft Scripting Runtime** afin d'ouvrir le fichier et l'utiliser.  
![Référence à ajouter VBA](https://raw.githubusercontent.com/MilesTEG1/Importation-Excel-Donnees-MyPeugeot/master/images/R%C3%A9f%C3%A9rence%20%C3%A0%20Ajouter%20au%20projet%20VBA.png)

---

## Ce qui fonctionne (à peu près)

* Import des fichiers trajets
* Prise en compte de fichiers avec plusieurs VIN
* Sélection des VIN affichés via un filtre Excel
* Calcul des moyennes et autres informations par VIN ou un ensemble de VIN (filtre sur tableau croisé dynamique Excel)
* Ajout des nouveaux trajets, sans remise à 0 initiale
* Utilisation d'une page d'accueil listant les VINs importés avec une correspondance d'un véhicule
* La page d'accueil contient le TCD et les boutons pour lancer les macros

![Feuille d'accueil](https://raw.githubusercontent.com/MilesTEG1/Importation-Excel-Donnees-MyPeugeot/master/images/Feuille%20d'accueil.png)

---

# To-do (priorité)

- [x] Renommer la feuille "Trajets-MyPeugeot" en "Trajets"
- [ ] Recruter des testeurs de l'app MyCytroen et MyDS
- [x] Tenir compte des noms déjà présent de véhicule pour proposer un nom par défaut de type "Ma voiture 1" ; "Ma voiture 2" ; etc ...
- [ ] Augmenter la taille du nombre de VIN possible (tableau [M;N])

## Next (ensuite)

- [ ] Refaire tous les graphiques
- [ ] Reconstruire certaines infos manquante en cas de trajets manquants :
  - [ ] Adresse de départ du tajet i = adresse d'arrivée du trajet i-1
  - [ ] Adresse de départ du tajet i = adresse de départ du trajet i+1
  - [ ] Kilométrage de départ du tajet i = Kilométrage d'arrivée du trajet i-1
  - [ ] Kilométrage d'arrivée du tajet i = Kilométrage de départ du trajet i+1
  - [ ] idem pour le volume de carburant
- [ ] Importer tous les champs non-importés ?
- [ ] Faire une version qui exporte les données dans un fichier JSON pour MyPeugeot, MyCytroen, MyDS
- [ ] Nettoyer le code de tous les commentaires inutiles de code inutilisé quand le code sera stabilisé
