
# To-do (priorité)

- [ ] Recruter de nouveaux testeurs (notemment de l'app MyCytroen et MyDS)
- [ ] Augmenter la taille du nombre de VIN possible (tableau [M;N])


## Next (ensuite)

- [ ] Refaire tous les graphiques
- [ ] Reconstruire certaines infos manquante en cas de trajets manquants :
  - [X] Adresse de départ du tajet i = adresse d'arrivée du trajet i-1
  - [X] Adresse de départ du tajet i = adresse de départ du trajet i+1
  - [ ] Kilométrage de départ du tajet i = Kilométrage d'arrivée du trajet i-1
  - [ ] Kilométrage d'arrivée du tajet i = Kilométrage de départ du trajet i+1
  - [X] idem pour le volume de carburant
- [ ] Nettoyer le code de tous les commentaires inutiles de code inutilisé quand le code sera stabilisé


# Ce qui a été réalisé
- [X] Import des fichiers trajets
- [X] Prise en compte de fichiers avec plusieurs VIN
- [X] Sélection des VIN affichés via un filtre Excel
- [X] Calcul des moyennes et autres informations par VIN ou un ensemble de VIN (filtre sur tableau croisé dynamique Excel)
- [X] Ajout des nouveaux trajets, sans remise à 0 initiale
- [X] Utilisation d'une page d'accueil listant les VINs importés avec une correspondance d'un véhicule
- [X] La page d'accueil contient le TCD et les boutons pour lancer les macros
- [X] Affichage de la dernière position connue pour le VIN sélectionné (uniquement un VIN sélectionné)
- [X] Affichage avec une décimale des kilométrages
- [X] Faire une version qui exporte les données dans un fichier JSON pour MyPeugeot, MyCytroen, MyDS
- [x] Renommer la feuille "Trajets-MyPeugeot" en "Trajets"
- [x] Tenir compte des noms déjà présent de véhicule pour proposer un nom par défaut de type "Ma voiture 1" ; "Ma voiture 2" ; etc ...
- [X] Ajouter l'adresse du dernier lieu connu correspondant au dernier trajet fait pour le VIN sélectionné
- [X] Détermination de la marque de la voiture
- [X] Importer tous les champs non-importés
- [X] Rendre compatible avec Excel macOS !!

