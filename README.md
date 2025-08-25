
# XRR-SCORE

*Génération automatique de fichiers d'import XRR (XML) pour l'outil SCORE de la FFVoile à partir d'une feuille Excel.*

[![GitHub License](https://img.shields.io/github/license/Mttwt9/XRR-SCORE?style=flat-square&color=blue)](https://github.com/Mttwt9/XRR-SCORE/blob/main/LICENSE)

---

## 📖 Introduction

Ce projet propose une macro en VBA permettant de générer un fichier XML au format XRR, compatible avec l'outil SCORE de la FFVoile, à partir d'une liste d'inscrits provenant d'un fichier Excel.

## 🛠️ Prérequis

- Microsoft Excel (avec macros activées)
- Connaissances de base en VBA recommandées

## 📦 Contenu du dépôt

- `SailingXML_TC.bas` : module VBA contenant le code source
- `Template_Source_TC.xlsx` : modèle Excel avec l'ordre des colonnes attendu
- `README.md` : ce fichier
- `LICENSE` : licence GNU GPL v3

> [!Note]
> Les fichiers suffixés par TC sont prévus pour des imports en temps compensé.
[À venir] Les fichiers avec le suffixe TR sont des dérivés des TC sans les attributs uniquement nécessaires au TC.

## 🚀 Installation

1. Télécharger [![GitHub Release](https://img.shields.io/github/v/release/Mttwt9/XRR-SCORE?style=flat-square&label=lastRelease&color=magenta)](https://github.com/Mttwt9/XRR-SCORE/releases/latest) ou cloner ce dépôt.
2. Ouvrir le fichier `Template_Source_TC.xlsx` et entrer les inscrits selon les colonnes prévues.
3. Ouvrir l'éditeur VBA (Alt+F11) dans Excel.
4. Importer le module `SailingXML_TC.bas` dans le projet VBA (menu Fichier > Importer un fichier...).

## 📝 Utilisation
> [!IMPORTANT]
> La macro utilise la feuille active pour lire les inscrits : assurez-vous d'être sur la bonne feuille avant d'exécuter la macro.

1. Ouvrir le fichier Excel avec les inscrits.
2. Exécuter la macro `CreateSailingXML`.
3. Le fichier XML sera généré et enregistré sur le bureau de l'utilisateur courant avec la date du jour dans le nom (ex : `%USERPROFILE%\Desktop\SailingXRR_2025-08-25.xml`).
4. Importer ce fichier dans SCORE.

## ⚙️ Personnalisation

- Les constantes `COL_xxx` définissent les index des colonnes. Si le fichier source diffère du modèle, il convient de modifier leurs valeurs.
- Si une colonne n'existe pas, indiquer l'index d'une colonne vide ou adapter le code :
> - Commenter les lignes de définition des constantes (`Dim COL_xxx`)
> - Corriger la création des attributs afférents aux constantes commentées en remplaçant `ws.Cells(i, COL_xxx).Value` par `""`.

## 📂 Exemple de résultat

```xml
<SailingXRR>
	<Person PersonID="123-P1" FamilyName="..." ... />
	<Person PersonID="123-P2" FamilyName="..." ... />
	<Boat BoatID="123-B1" SailNumber="..." ... />
	<Event CoID="123">
		<Team TeamID="123-T1" BoatID="123-B1" ...>
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
