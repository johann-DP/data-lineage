# excel_generator.py

Ce fichier contient deux classes, `ExcelGenerator` et `ExcelModifier`, qui fournissent des méthodes pour générer et modifier des fichiers Excel à partir de DataFrames pandas.

## Classe `ExcelGenerator`

### `__init__`

La méthode d'initialisation crée une instance de la classe `ExcelGenerator`.

### `append_data_to_excel`

Cette méthode prend un DataFrame pandas, le nom d'une feuille de calcul et un classeur Excel comme arguments, et ajoute le DataFrame à la feuille de calcul spécifiée dans le classeur.

### `generate_xls_and_folder`

Cette méthode génère un dossier avec le fichier Excel à l'intérieur.

### `append_calculation_sheet`

Cette méthode ajoute un nouvel onglet avec des formules de calcul à un classeur Excel.

### `append_risk_management_calculation_sheet`

Ajoute un nouvel onglet avec des formules de calcul à un classeur Excel pour le RiskManagement.

### `append_solvency_calculation_sheet`

Ajoute un nouvel onglet avec des formules de calcul à un classeur Excel pour le Solvency.

### `append_capital_modeling_calculations_sheet`

Ajoute un nouvel onglet pour les calculs de modélisation du capital.

### `reopen_and_save_workbook`

Cette méthode rouvre et enregistre un classeur existant.

### `append_reinsurance_pricing_calculations_sheet`

Cette méthode ajoute un nouvel onglet avec des formules de calcul de tarification de réassurance à un classeur Excel.

### `append_market_indices_to_excel`

Cette méthode ajoute les indices de marché à un classeur Excel.

### `append_customer_details_to_excel`

Cette méthode ajoute les détails des clients à un classeur Excel.

### `append_policy_details_to_excel`

Cette méthode ajoute les détails de la politique à un classeur Excel.

### `append_actuarial_calculation_sheet`

Cette méthode ajoute un nouvel onglet avec des formules de calcul actuarielles à un classeur Excel.

## Classe `ExcelModifier`

### `__init__`

La méthode d'initialisation crée une instance de la classe `ExcelModifier`.

### `add_foreign_keys`

Cette méthode ajoute des clés étrangères à toutes les feuilles de calcul d'un fichier Excel.
