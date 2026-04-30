import win32com.client
import openpyxl
import os
import time
import pythoncom
import tkinter as tk
from tkinter import filedialog
from openpyxl.styles import Font, PatternFill, Alignment
from concurrent.futures import ThreadPoolExecutor, as_completed
from functools import lru_cache
import threading

def demander_fichier_asm():
    """Ouvre l'explorateur pour choisir le fichier ASM."""
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title="Sélectionnez l'assemblage principal (.asm)", 
        filetypes=[("Assemblage Solid Edge", "*.asm")]
    )

def _scanner_dossier_thread_safe(args):
    """Fonction worker pour scanner un dossier (thread-safe)."""
    chemin_dossier, depth, max_depth = args
    plans_trouves = []
    sous_dossiers = []
    dossiers_count = 1
    
    try:
        with os.scandir(chemin_dossier) as it:
            for entry in it:
                if entry.is_file() and entry.name.lower().endswith('.dft'):
                    nom_base = os.path.splitext(entry.name)[0].lower()
                    plans_trouves.append((nom_base, entry.path))
                elif entry.is_dir() and depth < max_depth:
                    sous_dossiers.append((entry.path, depth + 1))
    except (PermissionError, OSError):
        pass
    
    return plans_trouves, sous_dossiers, dossiers_count

def indexer_les_plans_projet_entier(chemin_asm_initial, dossier_dft=None, mode_recherche="les_deux", max_depth=3, callback_progress=None):
    """Parcourt le dossier et les sous-dossiers pour trouver tous les plans .dft.
    Version multithreadée avec ThreadPoolExecutor pour I/O parallèles.
    
    Modes de recherche:
    - "arborescence": cherche uniquement dans l'arborescence remontée depuis l'ASM
    - "dossier_specifique": cherche uniquement dans le dossier spécifique
    - "les_deux": cherche dans les deux (comportement par défaut)
    
    Args:
        callback_progress: Fonction appelée avec (dossiers_scannes, plans_trouves, total_a_faire)
    """
    index = {}
    if not chemin_asm_initial:
        return index

    print(f"--- Indexation des plans (.dft) [Mode: {mode_recherche}] ---")

    dossiers_a_scanner = []

    # Mode arborescence ou les_deux: ajouter l'arborescence depuis l'ASM
    if mode_recherche in ["arborescence", "les_deux"]:
        racine_projet = chemin_asm_initial
        for _ in range(2):
            parent = os.path.dirname(racine_projet)
            if not parent or parent == racine_projet:
                break
            racine_projet = parent
        
        print(f"Dossier racine : {racine_projet}")
        dossiers_a_scanner.append((racine_projet, 0))

    # Mode dossier_specifique ou les_deux: ajouter le dossier spécifique
    if mode_recherche in ["dossier_specifique", "les_deux"]:
        if dossier_dft and os.path.exists(dossier_dft):
            print(f"Dossier spécifique : {dossier_dft}")
            dossiers_a_scanner.append((dossier_dft, 0))
        elif mode_recherche == "dossier_specifique":
            print("AVERTISSEMENT: Aucun dossier spécifique n'a été sélectionné!")
            return index

    print(f"Profondeur max : {max_depth} niveaux")

    dossiers_traites = 0
    pile = dossiers_a_scanner.copy()
    
    # Utiliser ThreadPoolExecutor pour paralléliser le scan
    with ThreadPoolExecutor(max_workers=8) as executor:
        while pile:
            # Prendre un batch de dossiers à traiter
            batch = []
            while pile and len(batch) < 16:
                batch.append(pile.pop(0) + (max_depth,))
            
            if not batch:
                break
            
            # Soumettre les tâches en parallèle
            futures = {executor.submit(_scanner_dossier_thread_safe, args): args for args in batch}
            
            # Collecter les résultats
            for future in as_completed(futures):
                try:
                    plans_trouves, sous_dossiers, count = future.result()
                    
                    # Ajouter les plans trouvés
                    for nom_base, chemin in plans_trouves:
                        index[nom_base] = chemin
                    
                    # Ajouter les sous-dossiers à la pile
                    pile.extend(sous_dossiers)
                    dossiers_traites += count
                    
                    # Callback de progression
                    if callback_progress and dossiers_traites % 50 == 0:
                        callback_progress(dossiers_traites, len(index), len(pile) + dossiers_traites)
                        
                except Exception as e:
                    print(f"  Erreur scan: {e}")
            
            if dossiers_traites % 100 == 0:
                print(f"  ... {dossiers_traites} dossiers scannés, {len(index)} plans trouvés")

    print(f"-> {len(index)} plan(s) détecté(s) dans {dossiers_traites} dossiers")
    return index

class MetadataCache:
    """Cache thread-safe pour les métadonnées des documents."""
    def __init__(self, max_size=1000):
        self._cache = {}
        self._lock = threading.Lock()
        self._max_size = max_size
        self._access_order = []
    
    def get(self, key):
        with self._lock:
            if key in self._cache:
                # Mettre à jour l'ordre d'accès (LRU)
                self._access_order.remove(key)
                self._access_order.append(key)
                return self._cache[key]
            return None
    
    def set(self, key, value):
        with self._lock:
            if key in self._cache:
                self._access_order.remove(key)
            elif len(self._cache) >= self._max_size:
                # Éviction LRU
                oldest = self._access_order.pop(0)
                del self._cache[oldest]
            
            self._cache[key] = value
            self._access_order.append(key)
    
    def clear(self):
        with self._lock:
            self._cache.clear()
            self._access_order.clear()

# Instance globale du cache
g_metadata_cache = MetadataCache()

def lister_proprietes(doc_obj):
    """Liste toutes les propriétés disponibles pour le débogage."""
    try:
        print(f"  --- Propriétés disponibles ---")
        print(f"  Type de document: {type(doc_obj)}")
        
        # Lister les attributs principaux
        attrs = [attr for attr in dir(doc_obj) if not attr.startswith('_')]
        print(f"  Attributs: {attrs[:20]}...")  # Limiter l'affichage
        
        # Essayer Properties (au lieu de PropertySets)
        if hasattr(doc_obj, 'Properties'):
            print(f"  Properties disponible")
            try:
                for prop_set in doc_obj.Properties:
                    nom_set = prop_set.Name if hasattr(prop_set, 'Name') else "Sans nom"
                    print(f"    PropertySet: {nom_set}")
                    if nom_set == "Custom":
                        print(f"      Propriétés Custom:")
                        for prop in prop_set:
                            nom = prop.Name if hasattr(prop, 'Name') else "Sans nom"
                            valeur = prop.Value if hasattr(prop, 'Value') else ""
                            print(f"        - {nom}: {valeur}")
            except Exception as e:
                print(f"    Erreur Properties: {e}")
        
        # Essayer PropertySets
        if hasattr(doc_obj, 'PropertySets'):
            print(f"  PropertySets disponible")
            try:
                for prop_set in doc_obj.PropertySets:
                    print(f"  PropertySet: {prop_set.Name if hasattr(prop_set, 'Name') else 'Unknown'}")
                    for prop in prop_set:
                        nom = prop.Name if prop.Name else "Sans nom"
                        valeur = prop.Value if prop.Value is not None else ""
                        print(f"    - {nom}: {valeur}")
            except Exception as e:
                print(f"    Erreur PropertySets: {e}")
        
        # Essayer SummaryInformation
        if hasattr(doc_obj, 'SummaryInformation'):
            print(f"  SummaryInformation disponible")
            try:
                print(f"    Title: {doc_obj.SummaryInformation.Title}")
            except: pass
    except Exception as e:
        print(f"  Erreur listing propriétés: {e}")

def extraire_metadonnees_rapide(chemin_fichier, debug=False, use_cache=True):
    """Extrait les propriétés sans ouvrir le fichier dans Solid Edge (ultra rapide).
    
    Utilise SolidEdge.FileProperties qui lit les métadonnées directement depuis le fichier
    sans lancer l'interface graphique de Solid Edge.
    
    Args:
        chemin_fichier: Chemin complet du fichier (.asm, .par, .psm, .dft)
        debug: Activer le mode debug
        use_cache: Utiliser le cache pour éviter les doublons
    """
    meta = {"designation": "", "revision": "1", "version": "-"}
    
    # Vérifier le cache si activé
    if use_cache:
        cached = g_metadata_cache.get(chemin_fichier)
        if cached:
            if debug:
                print(f"  -> Cache hit pour {os.path.basename(chemin_fichier)}")
            return cached.copy()
    
    try:
        # Appelle le lecteur de propriétés (ultra rapide, ne lance pas l'interface 3D)
        prop_reader = win32com.client.Dispatch("SolidEdge.FileProperties")
        prop_reader.Open(chemin_fichier)
        
        # 1. Lire la désignation (SummaryInformation = 1)
        try:
            summary_props = prop_reader.Item("SummaryInformation")
            meta["designation"] = summary_props.Item("Title").Value
        except:
            pass
            
        # 2. Lire la révision/version (Custom = 4)
        noms_possibles = ["indice de modification", "revision index", "index", "revision", "rev"]
        try:
            custom_props = prop_reader.Item("Custom")
            for i in range(1, custom_props.Count + 1):
                prop = custom_props.Item(i)
                if prop.Name.lower() in noms_possibles:
                    meta["version"] = str(prop.Value).strip()
                    if debug:
                        print(f"  -> Version trouvée (FileProperties): {meta['version']}")
                    break
        except:
            pass
            
        prop_reader.Close()
        
    except Exception as e:
        if debug:
            print(f"  -> Erreur lecture rapide pour {os.path.basename(chemin_fichier)}: {e}")
    
    # Mettre en cache si activé
    if use_cache:
        g_metadata_cache.set(chemin_fichier, meta.copy())
    
    return meta

def extraire_metadonnees(doc_obj, debug=False, use_cache=True):
    """Récupère le titre, la version et force la révision à 1 depuis Solid Edge.
    
    Args:
        doc_obj: Objet document Solid Edge
        debug: Activer le mode debug
        use_cache: Utiliser le cache pour éviter les doublons
    """
    # Récupérer l'identifiant unique du document
    doc_id = None
    try:
        doc_id = doc_obj.FullName
    except:
        pass
    
    # Vérifier le cache si activé et document identifié
    if use_cache and doc_id:
        cached = g_metadata_cache.get(doc_id)
        if cached:
            if debug:
                print(f"  -> Cache hit pour {os.path.basename(doc_id)}")
            return cached.copy()
    
    meta = {"designation": "", "revision": "1", "version": "-"}
    
    if debug:
        lister_proprietes(doc_obj)
    
    # Optimisation: collecter toutes les propriétés en une seule passe
    custom_props = {}
    all_props = {}
    
    # Collecter Properties -> Custom en une passe
    if hasattr(doc_obj, 'Properties'):
        try:
            for prop_set in doc_obj.Properties:
                if hasattr(prop_set, 'Name') and prop_set.Name == "Custom":
                    for prop in prop_set:
                        if hasattr(prop, 'Name') and prop.Name:
                            nom_lower = prop.Name.lower()
                            if hasattr(prop, 'Value') and prop.Value:
                                custom_props[nom_lower] = str(prop.Value).strip()
                    break  # On a trouvé Custom, pas besoin de continuer
        except:
            pass
    
    # Collecter PropertySets en une passe (pour la version uniquement si pas trouvée)
    if hasattr(doc_obj, 'PropertySets'):
        try:
            for prop_set in doc_obj.PropertySets:
                for prop in prop_set:
                    if hasattr(prop, 'Name') and prop.Name and hasattr(prop, 'Value') and prop.Value:
                        nom_lower = prop.Name.lower()
                        all_props[nom_lower] = str(prop.Value).strip()
        except:
            pass
    
    # Essayer SummaryInformation pour la désignation (plus rapide)
    try:
        meta["designation"] = doc_obj.SummaryInformation.Title
    except:
        pass
    
    # Si pas trouvé dans SummaryInformation, chercher dans Custom
    if not meta["designation"]:
        for key in ["désignation", "designation", "title", "titre"]:
            if key in custom_props:
                meta["designation"] = custom_props[key]
                if debug:
                    print(f"  -> designation trouvée dans Custom: {meta['designation']}")
                break
    
    # Récupération de l'attribut "indice de modification"
    noms_possibles = [
        "indice de modification", "revision index", 
        "modification index", "index", "revision", "rev"
    ]
    
    # Chercher d'abord dans Custom (plus rapide)
    for key in custom_props:
        if any(nom in key for nom in noms_possibles):
            meta["version"] = custom_props[key]
            if debug:
                print(f"  -> Version trouvée dans Custom: {meta['version']}")
            break
    
    # Si pas trouvé, chercher dans PropertySets
    if meta["version"] == "-":
        for key in all_props:
            if any(nom in key for nom in noms_possibles):
                meta["version"] = all_props[key]
                if debug:
                    print(f"  -> Version trouvée dans PropertySets: {meta['version']}")
                break
    
    # Mettre en cache si activé
    if use_cache and doc_id:
        g_metadata_cache.set(doc_id, meta.copy())
    
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

        print("\nConnexion à Solid Edge...")
        try:
            # 1. Tente de se brancher sur un Solid Edge déjà ouvert (Instantané)
            app = win32com.client.GetActiveObject("SolidEdge.Application")
            print("Connecté à l'instance existante de Solid Edge.")
        except pythoncom.com_error:
            # 2. S'il n'est pas ouvert, on le lance (Prend quelques secondes)
            print("Démarrage de Solid Edge en arrière-plan...")
            app = win32com.client.dynamic.Dispatch("SolidEdge.Application")
            app.Visible = False
            print("Solid Edge démarré.")
        
        # Désactiver les alertes pour accélérer l'ouverture des fichiers
        app.DisplayAlerts = False
        
        doc_racine = app.Documents.Open(chemin_asm)
        # COM API bloque jusqu'à ce que le document soit chargé, pas besoin de sleep 

        lignes_excel = []
        compteur_ordre = 1
        # Liste pour stocker les couples (DFT, 3D) pour les ajouter à la fin
        liste_plans_a_rajouter = [] 
        stats = {"3d": 0, "2d": 0}

        def get_suffixe_fichier(nom_fichier):
            """Retourne le suffixe approprié selon l'extension du fichier."""
            ext = os.path.splitext(nom_fichier)[1].lower()
            if ext == '.asm':
                return "(ASM)"
            elif ext in ['.par', '.psm']:
                return "(PRT)"
            elif ext == '.dft':
                return "(DRW)"
            return ""

        def ajouter_ligne(niveau, relation, nom_fichier, chemin_complet, classe, qte=1, rev="1", desig="", ver="-"):
            nonlocal compteur_ordre
            ref_util = os.path.splitext(nom_fichier)[0]
            special_cad = os.path.splitext(nom_fichier)[0]  # Sans extension
            suffixe = get_suffixe_fichier(nom_fichier)
            # Normaliser le chemin (remplace / par \) et concaténer sans espace
            chemin_normalise = os.path.normpath(chemin_complet)
            attachement = f"{chemin_normalise}{suffixe}" if suffixe else chemin_normalise

            lignes_excel.append([
                niveau, relation, compteur_ordre, qte, "", special_cad,
                classe, ref_util, ver, rev, desig, "", attachement
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
                # Utiliser la méthode rapide avec le chemin du fichier
                meta = extraire_metadonnees_rapide(data["chemin"])

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
        
        # Racine du projet - utiliser la méthode rapide
        meta_root = extraire_metadonnees_rapide(doc_racine.FullName)
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
        print("Extraction métadonnées des plans...")
        plans_deja_traites = set()

        for item in liste_plans_a_rajouter:
            if item["dft_path"] not in plans_deja_traites:
                # Utiliser la méthode ultra-rapide (FileProperties) au lieu d'ouvrir le document
                meta_dft = extraire_metadonnees_rapide(item["dft_path"])
                
                # Le plan (DFT) est le Parent (Level 0) - utilise la designation de la pièce 3D associée
                ajouter_ligne(0, "", item["dft_nom"], item["dft_path"], "CAD_DRAWING_A", 1, meta_dft["revision"], item["src_desig"], meta_dft["version"])
                # Le fichier 3D associé devient l'enfant (Level 1)
                ajouter_ligne(1, "Drawing", item["src_nom"], item["src_path"], item["src_classe"], 1, item["src_rev"], item["src_desig"], item["src_ver"])
                plans_deja_traites.add(item["dft_path"])

        # Génération Excel
        print("Génération du fichier Excel...")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Structure"
        
        headers = ["Level", "Relationship", "ordre", "quantite", "repere", "SpecialCAD", "Class", "ref_utilisat", "version", "revision", "designation", "dia_se", "Attachments"]
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