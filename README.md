# Projet : Logiciel de validation
#### *Auteur : Hakim AOUDIA*
#### *Langage : C# (Windows Forms), VBA*
#### *Date : 2022*

<p align="center">
  <img src="https://user-images.githubusercontent.com/106891439/222987208-4381a267-deca-41da-bb1b-b743b7e9c7a5.png" width="40%" height="40%">
</p>

Avec le logiciel, on passe du fichier .csv (à gauche) à un fichier Excel (à droite) avec des boutons pouvant lancer des scripts VBA.

<p align="center">
  <img src="https://user-images.githubusercontent.com/106891439/233512076-b0a7b886-41cd-4e4f-bb58-22448b4d8a85.png" width="40%" height="40%">
  <img src="https://user-images.githubusercontent.com/106891439/233512085-24a08db9-046b-4e15-9179-3a7874e86c9a.png" width="40%" height="40%">
</p>

## Présentation
Le logiciel est un logiciel de traitement de feuilles Excel. Le but ici est d'aider à simplifier des tâches redondantes pour éviter les erreurs humaines et accélérer le processus de validation. 

## Utilisation

### Lots
Les lots donnent des informations sur le lot utilisé pour chaque série. Cela varie selon le protocole utilisé (Perkin, pentalyse...).

### Chemins
L'utilisateur doit indiquer ou il faut importer les futurs fichiers .xlsm (chemin V1).
Les chemins Penta et Perkin donnent le chemin de destination des exportations au format .csv.

### Importé un fichier CSV
Pour importer un fichier csv sur lequel on va traiter les informations, il y a deux possibilités.

#### Glissez déposer
Soit faire un glissez-déposez avec le fichier en question
<img src="https://user-images.githubusercontent.com/106891439/222988010-6bfd4633-e097-4429-a619-bdc04e225a58.png" width="25%" height="25%">

#### Bouton ouvrir
Soit ouvrir le répertoire puis sélectionner le fichier.
<img src="https://user-images.githubusercontent.com/106891439/222988059-a6afa701-b83a-409e-a574-cdb4e9021e93.png" width="25%" height="25%">

## Partie excel (VBA)
Le fichier Excel s'ouvre automatiquement dès que la conversions csv -> xlsm se termine. Le fichier .xlsm ouvert possède 4 boutons.
Chaque bouton appelle du code VBA déjà importé avec le logiciel.

### Trier
- Tri dans l'ordre croissant la colonne Target
- Tri dans l'ordre croissant la colonne Well
- Tri dans l'ordre croissant la colonne Content
- Colorie un damier de 5 lignes par 5 lignes en alternant deux couleurs différentes.

- Tri chaque gène selon une valeur de CT choisis.

### Exporter
On demande si l'utilisateur veut importer le fichier .xlsm en .csv dans le dossier qu'il a précédemment indiqué lors de l'ouverture du logiciel.

### Mise en page
- Tri en sorte de récupérer seulement les patients positifs (avec un Sarscov <= 35 sans les NaN).
- Compte le nombre de positifs et écrit le nombre dans la cellule A18
- Change la taille des colonnes. (pour préparer l'impression)

### Imprimer
- Imprime la feuille avec les patients triés qui sera récupéré par les techniciens.
