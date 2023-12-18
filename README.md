# Data Lineage Project

## Description

Ce projet vise à fournir le data lineage pour une compagnie d'assurance, avec un focus sur le département actuariel. Il inclut la création de fichiers Excel simulés, l'établissement d'un graphe représentant toutes les variables présentes dans les différents onglets de tous les fichiers Excel fictifs, la génération d'une visualisation des données du graphe sous forme de page web en utilisant plotly/Dash, et la génération d'un rapport en deux parties : un fichier Excel qui donne la piste d'audit complète des indicateurs calculés dans les fichiers de simulation et un fichier .pptx qui montre cette piste d'audit de manière plus adaptée pour la communication entre les utilisateurs et les professionnels des métiers de la compagnie d'assurance.

## Installation
...

## Structure du Projet

La structure du répertoire est la suivante :

- Data_Lineage
	- data_lineage_project
		- .gitignore
		- bin
		- data
			- data_lineage_report.pptx
			- Insurance_Company
				- Actuarial
					- Pricing_and_Modelling
						- Capital_Modeling.xlsx
						- Pricing.xlsx
					- RiskManagement.xlsx
					- Solvency.xlsx
				- Claims_and_Policy
					- Claim_Data.xlsx
					- Policy_Data.xlsx
					- Reserving.xlsx
				- Customer_Relations
					- Agent_Data.xlsx
					- Customer_Data.xlsx
					- Product_Data.xlsx
				- Finance
					- Market_Data.xlsx
					- Reinsurance.xlsx
			- lineage_graph (2).gexf
			- lineage_graph.gexf
		- docs
		- lib
		- README.md
		- requirements.txt
		- results
		- scripts
			- lineage_dataviz.py
			- lineage_graph.py
			- lineage_reporting.py
		- setup.py
		- src
			- generation
				- data_generator.py
				- excel_generator.py
				- \__init__.py
				- \__pycache__
					- data_generator.cpython-39.pyc
					- excel_generator.cpython-39.pyc
			- graph_generator.py
			- main.py
			- \__pycache__
				- graph_generator.cpython-39.pyc
		- tests

## Usage

### Simulation de fichiers Excel d'actuariat

Le projet commence par la simulation de plusieurs fichiers Excel qui représentent les différentes facettes de l'activité d'une compagnie d'assurance. Ces fichiers sont stockés dans le dossier `data/Insurance_Company`. Chaque fichier représente un département ou une fonction spécifique au sein de la compagnie, comme l'actuariat, la gestion des sinistres et des polices, la relation client, et la finance. Chaque fichier contient plusieurs onglets qui représentent différents aspects des fonctions respectives.

### Établissement du graphe de lignage des données

Une fois les fichiers Excel simulés, le programme établit un graphe représentant l'ensemble des variables présentes dans les différents onglets de tous les fichiers. Ce graphe permet de visualiser les relations entre les différentes variables et de comprendre comment elles interagissent.

### Visualisation des données avec Plotly/Dash

La visualisation des données est réalisée avec Plotly/Dash et est accessible via une page web. Cette page web comprend plusieurs fonctionnalités pour explorer le graphe de lignage des données, notamment un menu déroulant permettant de sélectionner les indicateurs à auditer et un autre menu déroulant permettant de sélectionner les arêtes à afficher. Ces arêtes peuvent représenter les liens entre les variables d'un même fichier ou les relations entre les fichiers.

### Reporting

Le projet produit également un rapport en deux parties. La première partie est un fichier Excel qui donne l'intégralité de la piste d'audit des indicateurs calculés dans les fichiers de simulation. La deuxième partie est un fichier .pptx qui présente la piste d'audit de manière plus visuelle et adaptée à la communication entre les utilisateurs et les professionnels des métiers de l'assurance.

Pour utiliser le projet, il suffit de lancer le script `main.py` qui se charge d'exécuter toutes les étapes du projet, de la simulation des fichiers Excel à la génération du rapport.


## A date...

... Seules les deux premières parties sont réalisées.


## Documentations

Les parties relatives aux fichiers .py sont données dans 'docs'
