# Importation-Excel-Donnees-MyPeugeot
Importation des données issues de l'application MyPeugeot dans un tableau Excel (Présence de macros)

Fichier excel avec macros qui permet de récupérer les données de l'application MyPeugeot et d'en faire des statistiques.

Mon projet est une variante assez proche (j'ai piqué plein d'idées) à celui-ci :
[Trajets myp de MYPEUGEOT APP sous Excel toutes versions](https://www.forum-peugeot.com/Forum/threads/trajets-myp-de-mypeugeot-app-sous-excel-toutes-versions.9456/)
 : 
J'ai commencé [à en parler ici](https://www.forum-peugeot.com/Forum/threads/fichier-excel-macros-pour-r%C3%A9cup%C3%A9rer-les-trajets-de-lapplication-mypeugeot.119785/)


Il faut exporter ses données MyPeugeot en utilisant cet option :
![Option à utiliser](https://raw.githubusercontent.com/MilesTEG1/Importation-Excel-Donnees-MyPeugeot/master/images/Option%20pour%20exporter%20les%20trajets%20dans%20l'app%20MyPeugeot.png)
upload_2020-4-5_18-58-20.png
Ça va envoyer un email avec un fichier .myp.

Pourquoi refaire un fichier qui fonctionne ?
Et bien parce que pour moi il ne fonctionne pas. Les données du fichier .myp exportés pour chaque trajets ne sont pas dans le même ordre que celui pour lequel a été conçu le fichier de vr34.
Je ne sais pas pourquoi ce n'est pas le même ordre, mais toujours est-il que ça rend l'exploitation impossible.
Mais comme le fichier de données .myp est dans un format JSON, il est possible de faire autrement qu'avec une structure figée, il suffit de "parser" les données pour les récupérer.
Pour celà je me suis aidé de cette bibliothèque de fonctions :
[VBA-JSON-2.3.1](https://github.com/VBA-tools/VBA-JSON)

Le fichier utilisé est normalement déjà inclus dans le fichier excel.
Cependant il faut activer une référence : Microsoft Scripting Runtime afin d'ouvrir le fichier et l'utiliser.
![Référence à ajouter VBA](https://raw.githubusercontent.com/MilesTEG1/Importation-Excel-Donnees-MyPeugeot/master/images/R%C3%A9f%C3%A9rence%20%C3%A0%20Ajouter%20au%20projet%20VBA.png)



Je précise que pour le moment ne fonctionne que ceci :
- l'importation des données dans la feuille de calcul Trajets-MyPeugeot
- les quelques calculs fait dedans.
