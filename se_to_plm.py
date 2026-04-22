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
    """Récupère le titre, la version et force la révision à 1 depuis Solid Edge."""
    meta = {"designation": "", "revision": "1", "version": "-"}
    try:
        meta["designation"] = doc_obj.SummaryInformation.Title
        try:
            # Récupération de la version
            version_val = doc_obj.ProjectInformation.Version
            if version_val and str(version_val).strip() != "":
                meta["version"] = str(version_val).strip()
        except:
            pass
    except:
        pass
    return meta

def determiner_classe(nom_fichier, est_projet=False):
    """Détermine la classe PLM en fonction de l'extension du fichier."""
    if est_projet: return "SUB_ASSY_A"
    ext = os.path.splitext(nom_fichier)[1].lower()
    if ext == '.asm': return "SUB_ASSY_A"
    if ext in ['.par', '.psm']: return "PART_A"
    if ext == '.dft': return "CAD_DRAWING_A"
    return "Folder"

def lancer_extraction_plm():
    try:
        chemin_asm = demander_fichier_asm()
        if not chemin_asm: return

        index_plans = indexer_les_plans_projet_entier(chemin_asm)

        print("\nOuverture de Solid Edge...")
        app = win32com.client.dynamic.Dispatch("SolidEdge.Application")
        app.Visible = False 
        doc_racine = app.Documents.Open(chemin_asm)
        time.sleep(3) 

        lignes_excel = []
        compteur_ordre = 1
        # Liste pour stocker les couples (DFT, 3D) pour les ajouter à la fin
        liste_plans_a_rajouter = [] 
        stats = {"3d": 0, "2d": 0}

        def ajouter_ligne(niveau, relation, nom_fichier, chemin_complet, classe, qte=1, rev="1", desig="", ver="-"):
            nonlocal compteur_ordre
            ref_util = os.path.splitext(nom_fichier)[0]
            
            lignes_excel.append([
                niveau, relation, compteur_ordre, qte, "", nom_fichier, 
                classe, ref_util, ver, rev, desig, "", chemin_complet
            ])
            compteur_ordre += 1

        def explorer_occurrences(occurrences, niveau):
            nonlocal stats
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

                    if nom_reel not in dict_occ:
                        dict_occ[nom_reel] = {"qte": 1, "obj": occ, "chemin": path_reel}
                    else:
                        dict_occ[nom_reel]["qte"] += 1
                except: continue

            for nom, data in dict_occ.items():
                stats["3d"] += 1
                classe_3d = determiner_classe(nom)
                meta = {"designation": "", "revision": "1", "version": "-"}
                try:
                    meta = extraire_metadonnees(data["obj"].OccurrenceDocument)
                except: pass

                # Ajout de la pièce/sous-assemblage 3D dans l'arbre principal
                ajouter_ligne(niveau, "ComposedOf", nom, data["chemin"], classe_3d, data["qte"], meta["revision"], meta["designation"], meta["version"])
                
                # Vérification si un plan existe (mais on NE l'ajoute PAS dans l'arbre principal)
                nom_sans_ext = os.path.splitext(nom)[0].lower()
                if nom_sans_ext in index_plans:
                    chemin_dft = index_plans[nom_sans_ext]
                    nom_dft = os.path.basename(chemin_dft)
                    
                    # On stocke l'information pour générer la structure inversée à la fin
                    liste_plans_a_rajouter.append({
                        "dft_nom": nom_dft, 
                        "dft_path": chemin_dft,
                        "src_nom": nom,
                        "src_path": data["chemin"],
                        "src_classe": classe_3d,
                        "src_rev": meta["revision"],
                        "src_desig": meta["designation"],
                        "src_ver": meta["version"]
                    })
                    stats["2d"] += 1
                
                if data["obj"].Subassembly:
                    try: explorer_occurrences(data["obj"].OccurrenceDocument.Occurrences, niveau + 1)
                    except: pass

        print("\nAnalyse de la structure...")
        
        # Racine du projet
        meta_root = extraire_metadonnees(doc_racine)
        nom_root = os.path.basename(doc_racine.FullName)
        ajouter_ligne(0, "", nom_root, doc_racine.FullName, "SUB_ASSY_A", 1, meta_root["revision"], meta_root["designation"], meta_root["version"])

        # Plan de la racine
        nom_root_pur = os.path.splitext(nom_root)[0].lower()
        if nom_root_pur in index_plans:
            path_dft_root = index_plans[nom_root_pur]
            nom_dft_root = os.path.basename(path_dft_root)
            liste_plans_a_rajouter.append({
                "dft_nom": nom_dft_root, "dft_path": path_dft_root,
                "src_nom": nom_root, "src_path": doc_racine.FullName,
                "src_classe": "SUB_ASSY_A", "src_rev": meta_root["revision"], "src_desig": meta_root["designation"],
                "src_ver": meta_root["version"]
            })
            stats["2d"] += 1

        # Exploration récursive de l'arbre 3D
        explorer_occurrences(doc_racine.Occurrences, 1)

        # --- A LA FIN : AJOUT DES PLANS EN LEVEL 0 ET LEURS 3D EN LEVEL 1 ---
        print("Ajout des plans (DFT) à la fin du fichier...")
        plans_deja_traites = set()

        for item in liste_plans_a_rajouter:
            if item["dft_path"] not in plans_deja_traites:
                # Le plan (DFT) est le Parent (Level 0)
                ajouter_ligne(0, "", item["dft_nom"], item["dft_path"], "CAD_DRAWING_A")
                # Le fichier 3D associé devient l'enfant (Level 1)
                ajouter_ligne(1, "Drawing", item["src_nom"], item["src_path"], item["src_classe"], 1, item["src_rev"], item["src_desig"], item["src_ver"])
                plans_deja_traites.add(item["dft_path"])

        # Génération Excel
        print("Génération du fichier Excel...")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Import PLM"
        
        headers = ["Level", "Relationship", "ordre", "quantite", "repere", "Fichier_Ref", "Class", "ref_utilisat", "version", "revision", "designation", "dia_se", "Attachments"]
        ws.append(headers)
        
        header_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
        orange_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
        
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = header_fill if cell.column < 13 else orange_fill
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