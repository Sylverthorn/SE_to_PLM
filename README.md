# SE_to_PLM

Script pour extraire la structure PDM/PLM d'un assemblage Solid Edge et l'exporter vers un fichier Excel.

## Installation

1. Activer l'environnement virtuel (si nécessaire):
```bash
python -m venv venv
venv\Scripts\activate
```

2. Installer les dépendances:
```bash
pip install -r requirements.txt
```

## Utilisation

Lancer le script et fournir le chemin du fichier d'assemblage Solid Edge (.asm):
```bash
python extract_relational_links.py
```

Le script vous demandera d'entrer le chemin du fichier `.asm` principal.

Le fichier Excel résultant sera sauvegardé dans le même dossier que le fichier `.asm` sous le nom `liens_relationnels.xlsx`.

## Format de sortie

Le fichier Excel contient les colonnes suivantes selon la logique PDM/PLM:
- **Level**: Niveau de profondeur dans la hiérarchie (0 = racine, 1 = enfant direct, etc.)
- **Relationship**: Type de lien sémantique:
  - `ComposedOf`: Lien d'assemblage physique (l'assemblage contient le composant)
  - `Drawing`: Lien de documentation (le plan met en plan le composant 3D)
- **Class**: Typologie du fichier:
  - `PART_A`: Pièce unitaire (.par ou .psm)
  - `SUB_ASSY_A`: Sous-assemblage (.asm ou .pwd)
  - `CAD_DRAWING_A`: Mise en plan (.dft)
- **quantite**: Nombre d'occurrences du composant
- **ref_utilisat**: Nom de référence (nom du fichier sans extension)
- **version**: Version du document
- **revision**: Révision du document

## Fonctionnalités

- Extraction récursive de la structure d'assemblage
- Identification automatique des fichiers de dessin associés (.dft)
- Filtrage des pièces exclues de la nomenclature (IncludeInBom)
- Classification correcte selon les types de documents Solid Edge (Part, Assembly, Draft, SheetMetal, Weldment)

## Prérequis

- Solid Edge installé sur la machine
- Windows (l'automatisation COM fonctionne uniquement sur Windows)
- Python 3.8 ou supérieur