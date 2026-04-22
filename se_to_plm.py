import win32com.client
import openpyxl
import os
import time
import tkinter as tk
from tkinter import filedialog
from openpyxl.styles import Font, PatternFill, Alignment

def demander_fichier_asm():
    """Ouvre l'explorateur pour choisir le fichier ASM."""
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title="Sélectionnez l'assemblage principal (.asm)", 
        filetypes=[("Assemblage Solid Edge", "*.asm")]
    )

def indexer_les_plans_projet_entier(chemin_asm_initial):
    """Parcourt le dossier et les sous-dossiers pour trouver tous les plans .dft."""
    index = {}
    if not chemin_asm_initial: return index
    
    dossier_asm = os.path.dirname(chemin_asm_initial)
    racine_projet = os.path.dirname(dossier_asm)

    print(f"--- Indexation globale des plans (.dft) ---")
    print(f"Scan en cours : {racine_projet}")
    
    for dossier, _, fichiers in os.walk(racine_projet):
        for fichier in fichiers:
            if fichier.lower().endswith('.dft'):
                nom_base = os.path.splitext(fichier)[0].lower()
                index[nom_base] = os.path.join(dossier, fichier)
                
    print(f"-> {len(index)} plan(s) détecté(s).")
    return index

def extraire_metadonnees(doc_obj):
    """Récupère le titre et la révision depuis Solid Edge."""
    meta = {"designation": "", "revision": ""}
    try:
        # Propriétés standard de Solid Edge
        meta["designation"] = doc_obj.SummaryInformation.Title
        meta["revision"] = doc_obj.SummaryInformation.RevisionNumber
    except:
        pass
    return meta

def determiner_classe(nom_fichier, est_projet=False):
    """Détermine la classe PLM en fonction de l'extension du fichier."""
    if est_projet: return "Projet"
    ext = os.path.splitext(nom_fichier)[1].lower()
    if ext == '.asm': return "ASM"
    if ext in ['.par', '.psm']: return "PART"
    if ext == '.dft': return "Plan_A"
    return "Folder"

def lancer_extraction_plm():
    try:
        # On sélectionne le fichier et on indexe les plans
        chemin_asm = demander_fichier_asm()
        if not chemin_asm: return

        index_plans = indexer_les_plans_projet_entier(chemin_asm)

        # On se connecte à Solid Edge
        print("\nOuverture de Solid Edge...")
        app = win32com.client.dynamic.Dispatch("SolidEdge.Application")
        app.Visible = False 
        doc_racine = app.Documents.Open(chemin_asm)
        time.sleep(3) 

        lignes_excel = []
        compteur_ordre = 1
        stats = {"3d": 0, "2d": 0}

        def ajouter_ligne(niveau, relation, nom_fichier, chemin_complet, classe, qte=1):
            nonlocal compteur_ordre
            
            # On essaie de lire les propriétés si le fichier est accessible
            designation = ""
            rev = ""
            ref_util = os.path.splitext(nom_fichier)[0]
            
            if chemin_complet and os.path.exists(chemin_complet):
                try:
                    # Pour les documents déjà ouverts
                    # (Note: lire chaque .par ralentit, on peut se limiter au niveau 0)
                    pass 
                except: pass

            lignes_excel.append([
                niveau, relation, compteur_ordre, qte, "", "", nom_fichier, 
                classe, ref_util, "", rev, designation, "", chemin_complet
            ])
            compteur_ordre += 1

        def explorer_occurrences(occurrences, niveau):
            nonlocal stats, compteur_ordre
            if occurrences is None: return
            
            dict_occ = {}
            for i in range(1, occurrences.Count + 1):
                try:
                    occ = occurrences.Item(i)
                    path_reel = ""
                    nom_reel = ""
                    try:
                        path_reel = occ.OccurrenceDocument.FullName
                        nom_reel = os.path.basename(path_reel)
                    except:
                        nom_reel = occ.Name.split(':')[0]
                        # Au cas où la pièce est inactive, on essaie de trouver le chemin
                        # (Optionnel: demande plus de ressources)

                    if nom_reel not in dict_occ:
                        dict_occ[nom_reel] = {"qte": 1, "obj": occ, "chemin": path_reel}
                    else:
                        dict_occ[nom_reel]["qte"] += 1
                except: continue

            for nom, data in dict_occ.items():
                stats["3d"] += 1
                classe_3d = determiner_classe(nom)
                
                # On récupère les métadonnées si possible
                meta = {"designation": "", "revision": ""}
                try:
                    meta = extraire_metadonnees(data["obj"].OccurrenceDocument)
                except: pass

                # On ajoute la ligne 3D
                lignes_excel.append([
                    niveau, "ComposedOf", compteur_ordre, data["qte"], "", "", nom, 
                    classe_3d, os.path.splitext(nom)[0], "", meta["revision"], meta["designation"], "", data["chemin"]
                ])
                compteur_ordre += 1
                
                # On ajoute la ligne du plan si ça existe
                nom_sans_ext = os.path.splitext(nom)[0].lower()
                if nom_sans_ext in index_plans:
                    chemin_dft = index_plans[nom_sans_ext]
                    ajouter_ligne(niveau + 1, "Drawing", os.path.basename(chemin_dft), chemin_dft, "Plan_A")
                    stats["2d"] += 1
                
                if data["obj"].Subassembly:
                    try: explorer_occurrences(data["obj"].OccurrenceDocument.Occurrences, niveau + 1)
                    except: pass

        # On lance l'analyse
        print("\nAnalyse de la structure...")
        
        # On traite le fichier racine
        meta_root = extraire_metadonnees(doc_racine)
        nom_root = os.path.basename(doc_racine.FullName)
        lignes_excel.append([
            0, "", compteur_ordre, 1, "", "", nom_root,
            "Projet", os.path.splitext(nom_root)[0], "", meta_root["revision"], meta_root["designation"], "", doc_racine.FullName
        ])
        compteur_ordre += 1

        # On regarde si y a un plan pour la racine
        nom_root_pur = os.path.splitext(nom_root)[0].lower()
        if nom_root_pur in index_plans:
            path_dft_root = index_plans[nom_root_pur]
            ajouter_ligne(1, "Drawing", os.path.basename(path_dft_root), path_dft_root, "Plan_A")
            stats["2d"] += 1

        explorer_occurrences(doc_racine.Occurrences, 1)

        # On crée le fichier Excel
        print("Génération du fichier...")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Import PLM"
        
        headers = [
            "Level", "Relationship", "ordre", "quantite", "repere", "localisation", "Fichier_Ref",
            "Class", "ref_utilisat", "version", "revision", "designation", "dia_se", "Attachments"
        ]
        ws.append(headers)
        
        header_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
        orange_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
        
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = header_fill if cell.column < 14 else orange_fill
            cell.alignment = Alignment(horizontal="left")

        for l in lignes_excel: ws.append(l)

        for col in ws.columns:
            max_length = 0
            for cell in col:
                try: max_length = max(max_length, len(str(cell.value)))
                except: pass
            ws.column_dimensions[col[0].column_letter].width = max_length + 2
        
        nom_out = f"Export_PLM_{int(time.time())}.xlsx"
        wb.save(nom_out)
        print(f"\nFichier généré : {nom_out}")
        print(f"3D: {stats['3d']} | Plans: {stats['2d']}")

    except Exception as e:
        print(f"\nErreur : {e}")

if __name__ == "__main__":
    lancer_extraction_plm()
