6 points je pense pour la suite :
•	Quand tu charges les données json, est-ce que tu importes TOUS les champs ?
•	Petite amélioration possible avec ton "nom par défaut" des infos véhicule : en effet, si je charge mon fichier 3 vin, j'ai "ma voiture 1" jusqu'à ma "voiture 3". Si ensuite j'importe un nouveau fichier (avec vin non encore connu), il me repropose "ma voiture 1" alors qu'on devrait en être à "ma voiture 4" ==> tenir compte du nombre de VIN dans la liste des VIN avec infos voiture
•	J'augmenterais la taille du nombre de VIN dans ton tableau en M:N, ou à tout le moins, prévoir de faire des insertions de cellules pour rajouter des lignes au cas oùn tout en maintenant tes bordures. Ca je sais faire, mais n'ai pas le temps today
•	Je prévois une version avec EXPORT des données au format JSON. Très facile à faire, et cela permettrait si on corrige des données de pouvoir exporter et remettre ça dans le smartphone, avec sélection de la marque (car on ne peut pas importer un VIN Peugeot dans MyCitroën, sauf erreur de ma part (j'avoue ne pas avoir testé !)
•	Il faudrait, je pense, renommer la feuille "Trajets-MyPeugeot" en "Trajets" tout court puisque ça marche aussi pour Citroën et/ou DS
•	Partie graphique et donnees-recap à refaire/màj
