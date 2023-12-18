# Documentation

Ce script Python est conçu pour générer et organiser des données fictives pour une compagnie d'assurance. 

## Description des parties principales du script

### Génération de données et enregistrement dans Excel

Le script commence par générer différentes données pour l'assurance et la réassurance à l'aide de la classe `DataGenerator`. Ces données comprennent les identifiants, les contrats, l'historique des réclamations et les paramètres de tarification de la réassurance.

Ensuite, il utilise la classe `ExcelGenerator` pour créer et enregistrer ces données dans un fichier Excel.

Ces étapes sont répétées pour différentes sections de l'assurance, comme les données démographiques, les données sur les réclamations, les données sur les primes, les indices de marché, les détails des clients et les détails des polices.

### Structuration en arborescence des fichiers Excel

Après la génération de tous les fichiers Excel, le script crée une structure de dossier pour organiser les fichiers. Les fichiers sont ensuite déplacés dans les dossiers appropriés selon leur nature.

### Ajout des clés étrangères

Ensuite, le script ajoute des clés étrangères aux fichiers Excel. Cela est fait en utilisant la classe `ExcelModifier`. Les clés étrangères sont ajoutées pour faciliter l'interrelation entre les différents ensembles de données.

### Génération du graphe à partir des fichiers Excel

Enfin, le script génère un graphe à partir des fichiers Excel à l'aide de la classe `GraphGenerator`. Ce graphe peut aider à comprendre les relations entre les différents ensembles de données et peut être utilisé pour effectuer des analyses de données plus complexes.

## Comment exécuter le script

Pour exécuter le script, assurez-vous d'avoir toutes les dépendances nécessaires installées et exécutez simplement le script à partir de la ligne de commande : python main.py

Assurez-vous que le répertoire de travail courant est le même que celui où se trouve le script avant de l'exécuter.

## Dépendances

- pandas
- openpyxl
- networkx
- matplotlib

Installez ces dépendances avec pip en utilisant la commande suivante :

pip install pandas openpyxl networkx matplotlib