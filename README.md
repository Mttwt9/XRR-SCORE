# Macro VBA pour fichier XRR SCORE 
*📝 Création des fichiers d'import des inscrits au format XRR (XML) vers l'outil SCORE de la FFVoile*

Le fichier .vb contient le code de la macro permettant de générer le fichier XRR pour importer les inscrits dans une compétition.
Le fichier .xlsx est un modèle de tableau Excel avec les colonnes disposés dans l'ordre prévu des constantes de la macro (COL_xxx).

Pour utiliser la macro, il convient de vérifier les contantes et modifier eventuellement l'ordre des colonnes. Si certaines des colonnes n'existent pas dans le fichier source, il faudra indiquer l'index d'un colonne vide ou bien commenter la définition des variables concernées ET remplacer la valeur des setAttribute afférants (`ws.Cells(i, COL_XXX).Value` -> `""`).



![GitHub License](https://img.shields.io/github/license/Mttwt9/XRR-SCORE?style=flat-square&color=blue)
