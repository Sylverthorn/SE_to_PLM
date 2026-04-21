"""
Solid Edge → PDM/PLM BOM Extractor
=======================================================
# Auteur : Korichi Yanis
# Date de création : 20 Avril 2026
# Rôle du script : Parcourt récursivement un assemblage Solid Edge (.asm) et génère un fichier
Excel à importer dans le PLM Andros.
# v2 : Support mode dossier (batch) — traite chaque .asm racine séparément

Prérequis :
  - Windows + Solid Edge installé
  - pip install pywin32 openpyxl

Usage :
  python se_to_plm_full.py
"""
import win32com.client
import openpyxl
import os
import time
import tkinter as tk
from tkinter import filedialog
from openpyxl.styles import Font, PatternFill, Alignment

def demander_fichier_asm():
    """Ouvre une fenêtre pour sélectionner l'assemblage principal"""
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title="Sélectionnez l'assemblage principal (.asm)", 
        filetypes=[("Assemblage Solid Edge", "*.asm")]
    )

def indexer_les_plans_projet_entier(chemin_asm_initial):
    """
    Scanne le dossier de l'assemblage ET son dossier parent (et tous leurs sous-dossiers)
    pour trouver 100% des plans du projet.
    """
    index = {}
    if not chemin_asm_initial: 
        return index
        
    # 1. On trouve le dossier où est l'assemblage
    dossier_asm = os.path.dirname(chemin_asm_initial)
    # 2. On remonte d'un niveau pour englober tout le projet (Racine)
    racine_projet = os.path.dirname(dossier_asm)

    print(f"--- Indexation globale des plans (.dft) ---")
    print(f"Zone scannée : {racine_projet} (et tous ses sous-dossiers)")
    
    for dossier, _, fichiers in os.walk(racine_projet):
        for fichier in fichiers:
            if fichier.lower().endswith('.dft'):
                nom_base = os.path.splitext(fichier)[0].lower()
                index[nom_base] = os.path.join(dossier, fichier)
                
    print(f"-> {len(index)} plan(s) trouvé(s) et mis en mémoire.")
    return index

def lancer_extraction_plm():
    try:
        # --- 1. SÉLECTION ET INDEXATION ---
        chemin_asm = demander_fichier_asm()
        if not chemin_asm: 
            print("Opération annulée.")
            return

        index_plans = indexer_les_plans_projet_entier(chemin_asm)

        # --- 2. CONNEXION SOLID EDGE ---
        print("\nConnexion à Solid Edge...")
        app = win32com.client.dynamic.Dispatch("SolidEdge.Application")
        app.Visible = True 
        
        print(f"Ouverture de l'assemblage...")
        doc_racine = app.Documents.Open(chemin_asm)
        
        print("Attente de chargement de l'arbre (3 sec)...")
        time.sleep(3) 

        # --- 3. PRÉPARATION DES DONNÉES ---
        lignes_excel = []
        compteur_ordre = 1
        stats = {"3d": 0, "2d": 0}

        def ajouter_ligne(niveau, relation, nom_fichier):
            nonlocal compteur_ordre
            lignes_excel.append([niveau, relation, compteur_ordre, 1, "", "", nom_fichier])
            compteur_ordre += 1

        def traiter_document_principal(doc_obj, niveau_actuel):
            nonlocal stats
            try: nom_fichier = os.path.basename(doc_obj.FullName)
            except: nom_fichier = os.path.basename(doc_obj.Name)
                
            nom_pur = os.path.splitext(nom_fichier)[0].lower()
            if nom_pur in index_plans:
                nom_dft = os.path.basename(index_plans[nom_pur])
                ajouter_ligne(niveau_actuel + 1, "Drawing", nom_dft)
                stats["2d"] += 1

        def explorer_occurrences(occurrences, niveau):
            nonlocal stats, compteur_ordre
            if occurrences is None: return
            
            # Regroupement des pièces identiques du même niveau
            dict_occ = {}
            for i in range(1, occurrences.Count + 1):
                try:
                    occ = occurrences.Item(i)
                    nom_reel = ""
                    
                    # Tentative de récupérer le vrai chemin, sinon Fallback sur le nom de l'arbre
                    try:
                        path_reel = occ.OccurrenceDocument.FullName
                        nom_reel = os.path.basename(path_reel)
                    except:
                        nom_reel = occ.Name.split(':')[0]

                    if nom_reel not in dict_occ:
                        dict_occ[nom_reel] = {"qte": 1, "obj": occ}
                    else:
                        dict_occ[nom_reel]["qte"] += 1
                except: 
                    continue # Ignore les pièces corrompues ou virtuelles

            for nom, data in dict_occ.items():
                stats["3d"] += 1
                
                # Ajout de la pièce 3D
                lignes_excel.append([niveau, "ComposedOf", compteur_ordre, data["qte"], "", "", nom])
                compteur_ordre += 1
                
                # Recherche et ajout du plan (Drawing)
                nom_sans_ext = os.path.splitext(nom)[0].lower()
                if nom_sans_ext in index_plans:
                    ajouter_ligne(niveau + 1, "Drawing", os.path.basename(index_plans[nom_sans_ext]))
                    stats["2d"] += 1
                
                # Descente dans le sous-assemblage (Récursivité)
                if data["obj"].Subassembly:
                    try: 
                        explorer_occurrences(data["obj"].OccurrenceDocument.Occurrences, niveau + 1)
                    except: 
                        pass

        # --- 4. EXÉCUTION DU SCAN ---
        print("\n--- DEBUT DE L'ANALYSE DE L'ARBRE ---")
        try: nom_racine = os.path.basename(doc_racine.FullName)
        except: nom_racine = os.path.basename(doc_racine.Name)
            
        # Ajout du Niveau 0 (Assemblage principal) et de son plan
        ajouter_ligne(0, "", nom_racine)
        traiter_document_principal(doc_racine, 0)
        
        # Exploration de tout le reste
        explorer_occurrences(doc_racine.Occurrences, 1)

        # --- 5. GÉNÉRATION DU FICHIER EXCEL ---
        print("\nGénération du fichier Excel...")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Import PLM"
        ws.append(["Level", "Relationship", "ordre", "quantite", "repere", "localisation", "Fichier_Ref"])
        
        for l in lignes_excel: ws.append(l)

        # Style des entêtes
        header_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="left")

        # Ajustement de la largeur des colonnes
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try: max_length = max(max_length, len(str(cell.value)))
                except: pass
            ws.column_dimensions[column].width = max_length + 2
        
        # Sauvegarde
        nom_out = f"Nomenclature_{os.path.splitext(os.path.basename(chemin_asm))[0]}.xlsx"
        wb.save(nom_out)
        
        # --- 6. RAPPORT FINAL ---
        print(f"\n" + "="*30)
        print(f"       RAPPORT FINAL")
        print(f"+" + "="*28)
        print(f" ✅ Articles 3D traités : {stats['3d']}")
        print(f" ✅ Plans 2D couplés  : {stats['2d']}")
        print(f" ✅ Fichier généré    : {nom_out}")
        print(f"+" + "="*28)

    except Exception as e:
        print(f"\n❌ Erreur critique : {e}")

if __name__ == "__main__":
    lancer_extraction_plm()
