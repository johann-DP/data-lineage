# graph_generator.py

Ce fichier contient la classe `GraphGenerator`, qui fournit des méthodes pour générer des graphes à partir de fichiers Excel et des formules contenues dans les cellules de ces fichiers.

## Classe `GraphGenerator`

### `__init__`

La méthode d'initialisation crée une instance de la classe `GraphGenerator`.

### `extract_column_references`

Extrait les références de colonne à partir d'une formule de cellule et les traduit en noms d'en-têtes de colonne.

### `walk_files_and_generate_graph`

Parcoure les fichiers dans un répertoire donné et génère un graphe en fonction des formules contenues dans les cellules de ces fichiers.

### `add_foreign_keys`

Ajoute les clés étrangères au graphe.

### `save_graph`

Sauvegarde le graphe au format GEXF.

### `create_subgraphs_and_visualize`

Crée des sous-graphes et les visualise. Il parcourt les dossiers et les fichiers, génère des graphes sans les clés étrangères, obtient les sous-graphes et les filtre en fonction de certaines conditions. Ensuite, il dessine et affiche les sous-graphes filtrés.
