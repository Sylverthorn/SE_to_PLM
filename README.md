# SE_to_PLM - Extracteur de données Solid Edge pour PLM

Outil d'extraction de données depuis Solid Edge vers un format Excel compatible avec les systèmes PLM (Product Lifecycle Management).

## 📋 Description

Ce projet permet d'extraire la structure complète d'un assemblage Solid Edge (fichiers `.asm`) et de générer un fichier Excel contenant :
- L'arborescence des composants 3D (pièces et sous-assemblages)
- Les plans 2D associés (fichiers `.dft`)
- Les métadonnées (désignation, révision)
- Les relations entre composants
- Les quantités de chaque pièce

L'outil scanne automatiquement le projet entier pour associer les plans correspondants aux pièces 3D.

## 🚀 Installation

### Prérequis

- **Windows** (l'outil utilise Solid Edge via COM)
- **Solid Edge** installé sur la machine
- **Python 3.8+**

### Installation automatique

Double-cliquez simplement sur `RUN.vbs` pour :
1. Créer automatiquement l'environnement virtuel
2. Installer les dépendances
3. Lancer l'interface graphique

### Installation manuelle

```bash
# Créer l'environnement virtuel
python -m venv venv

# Activer l'environnement
venv\Scripts\activate

# Installer les dépendances
pip install -r requirements.txt
```

## 📦 Dépendances

- `pywin32>=306` - Communication avec Solid Edge via COM
- `openpyxl>=3.1.2` - Génération des fichiers Excel
- `PyQt5>=5.15.9` - Interface graphique

## 💻 Utilisation

### Mode Graphique (Recommandé)

Lancez l'interface graphique en exécutant :

```bash
python gui_se_to_plm.py
```

Ou double-cliquez sur `RUN.vbs`.

**Interface :**
1. Cliquez sur "Parcourir..." pour sélectionner votre fichier `.asm` principal
2. Le nom de sortie est généré automatiquement (modifiable)
3. Cliquez sur "Lancer l'extraction"
4. Suivez la progression dans la console
5. Le fichier Excel est généré dans `Documents/Exports_PLM/`

### Mode Ligne de Commande

```bash
python se_to_plm.py
```

Le script vous demandera de sélectionner le fichier ASM via une boîte de dialogue.

## 📊 Structure du fichier Excel généré

Le fichier Excel contient les colonnes suivantes :

| Colonne | Description |
|---------|-------------|
| Level | Niveau d'imbrication dans l'assemblage |
| Relationship | Type de relation (ComposedOf, Drawing) |
| ordre | Numéro d'ordre |
| quantite | Quantité de la pièce |
| repere | Repère (vide) |
| localisation | Localisation (vide) |
| Fichier_Ref | Nom du fichier |
| Class | Classe PLM (Projet, ASM, PART, Plan_A) |
| ref_utilisat | Référence utilisateur |
| version | Version (vide) |
| revision | Révision depuis Solid Edge |
| designation | Désignation depuis Solid Edge |
| dia_se | Diamètre SE (vide) |
| Attachments | Chemin complet du fichier |

## 🔧 Fonctionnement

1. **Indexation des plans** : Scan récursif du dossier projet pour trouver tous les fichiers `.dft`
2. **Connexion Solid Edge** : Ouverture de l'assemblage principal via COM
3. **Traversée de l'assemblage** : Exploration récursive des occurrences
4. **Extraction métadonnées** : Récupération titre et révision depuis les propriétés Solid Edge
5. **Association plans** : Lien automatique entre pièces 3D et leurs plans 2D
6. **Génération Excel** : Création du fichier avec mise en forme

## 📁 Structure du projet

```
SE_to_PLM/
├── se_to_plm.py          # Script principal (mode CLI)
├── gui_se_to_plm.py      # Interface graphique PyQt5
├── requirements.txt      # Dépendances Python
├── RUN.vbs              # Lanceur automatique
└── README.md            # Ce fichier
```

## 🎯 Classes PLM détectées

- **Projet** : Assemblage racine
- **ASM** : Sous-assemblages (fichiers `.asm`)
- **PART** : Pièces (fichiers `.par`, `.psm`)
- **Plan_A** : Plans 2D (fichiers `.dft`)

## ⚠️ Notes importantes

- Solid Edge doit être installé et fonctionnel
- L'assemblage est ouvert en arrière-plan (`Visible = False`)
- Les fichiers inactifs ou manquants sont ignorés
- Le dossier de sortie par défaut est `Documents/Exports_PLM/`

## 🐛 Dépannage

**Solid Edge ne s'ouvre pas :**
- Vérifiez que Solid Edge est correctement installé
- Assurez-vous qu'aucune instance de Solid Edge n'est bloquée

**Erreur COM :**
- Redémarrez Solid Edge
- Vérifiez les permissions COM

**Plans non associés :**
- Vérifiez que les plans ont le même nom de base que les pièces 3D
- Assurez-vous que les plans sont dans le dossier du projet ou ses sous-dossiers

## 📝 Licence

Ce projet est fourni tel quel pour l'extraction de données Solid Edge vers des systèmes PLM.

## 🤝 Contribution

Les contributions sont les bienvenues pour améliorer l'outil et ajouter de nouvelles fonctionnalités.
