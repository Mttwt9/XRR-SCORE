
# XRR-SCORE

*Génération automatique de fichiers d'import XRR v2.0 (XML) pour l'outil SCORE de la FFVoile à partir d'une feuille Excel.*

[![GitHub License](https://img.shields.io/github/license/Mttwt9/XRR-SCORE?style=flat-square&color=blue)](https://github.com/Mttwt9/XRR-SCORE/blob/main/LICENSE)

---

## 📖 Introduction

Ce projet propose une macro en VBA permettant de générer un fichier XML au format XRR, compatible avec l'outil SCORE de la FFVoile, à partir d'une liste d'inscrits provenant d'un fichier Excel.

## 🛠️ Prérequis

- Microsoft Excel (avec macros activées)
- Connaissances de base en VBA recommandées

## 📦 Contenu du dépôt

- `SailingXML.bas` : module VBA (version 2025-09-29 -- XRR v2.0)
- `Template_Source.xlsx` : modèle Excel avec l'ordre des colonnes attendu
- `README.md` : ce fichier
- `LICENSE` : licence GNU GPL v3

> [!Note]
> La macro gère à la fois les inscriptions en temps compensé (attribut OsirisGuest) et en temps réel (attribut Ranking) au format XRR v2.0.
> L'encodage du fichier est en Windows-1252 (CP1252).

## 🚀 Installation

1. Télécharger [![GitHub Release](https://img.shields.io/github/v/release/Mttwt9/XRR-SCORE?style=flat-square&label=lastRelease&color=magenta)](https://github.com/Mttwt9/XRR-SCORE/releases/latest) ou cloner ce dépôt.
2. Ouvrir le fichier `Template_Source.xlsx` et entrer les inscrits selon les colonnes prévues.
Toutes les colonnes ne sont pas obligatoires. Pour les licenciés FFVoile, la licence suffit (mais pour plus de lisibilité, il est recommandé de fournir aussi le nom).
> [!TIP]
> Les dates de naissance doivent être saisies dans un format DATE d'Excel : la macro convertira automatiquement ce format en `yyyy-mm-jj` dans le fichier XML généré (format attendu par SCORE).
> 
> La seule colonne OBLIGATOIRE est la colonne *Numéro compétition*. Mais le but de cet automatisme est de donner le plus d'information possible dans le fichier d'import pour avoir le moins d'informations à saisir ! 
3. Ouvrir l'éditeur VBA (Alt+F11) dans Excel.
4. Importer le module `SailingXML.bas` dans le projet VBA (menu Fichier > Importer un fichier...).

## 📝 Utilisation
> [!IMPORTANT]
> La macro utilise la feuille active pour lire les inscrits : assurez-vous d'être sur la bonne feuille avant d'exécuter la macro.

1. Depuis le fichier Excel avec les inscrits.
2. Exécuter la macro `CreateSailingXML` depuis l'éditeur VBA ou depuis le bouton *Macro* du ruban si l'onglet *Développeur* est visible.
3. Le fichier XML sera généré et enregistré sur le bureau de l'utilisateur courant avec la date du jour dans le nom (ex : `%USERPROFILE%\Desktop\SailingXRR_2025-08-25.xml`).
4. Importer ce fichier dans SCORE.

## ⚙️ Personnalisation

- Les constantes `COL_xxx` définissent les index des colonnes. Si le fichier source diffère du modèle, il convient de modifier leurs valeurs.
- Si une colonne n'existe pas, indiquer l'index d'une colonne vide ou adapter le code. Pour cela :
> - Commenter les lignes de définition des constantes (`Dim COL_xxx`)
> - Corriger la création des attributs afférents aux constantes commentées en remplaçant `ws.Cells(i, COL_xxx).Value` par `""`.
*Le plus simple étant de respecter le fichier modèle !*

## 📂 Exemple de résultat

```xml
<SailingXRR>
    <Person PersonID="123-P1" FamilyName="..." ... />
    <Person PersonID="123-P2" FamilyName="..." ... />
    <Boat BoatID="123-B1" SailNumber="..." ... />
    <Event CoID="123">
        <Team TeamID="123-T1" BoatID="123-B1" Cat="..." NOC="...">
            <Crew PersonID="123-P1" Position="S" />
            <Crew PersonID="123-P2" Position="C" />
        </Team>
    </Event>
</SailingXRR>
```

## 📚 Références

- [Documentation SCORE FFVoile](https://arbitrage.ffvoile.fr/logiciel-de-classement/)
- [Format XRR SCORE](https://arbitrage.ffvoile.fr/media/tuxghvae/xrr_inscriptions_documentation.zip)

## 📝 Licence

Ce projet est sous licence [GNU GPL v3](LICENSE).

## 🙋 Support / Contact

Pour toute question, suggestion ou bug, ouvrez une issue sur GitHub ou contactez [Mttwt9](https://github.com/Mttwt9).
