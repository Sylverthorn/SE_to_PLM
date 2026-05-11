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

def normaliser_date(valeur):
    """Convertit divers formats de date en '25/04/2025 12:00:00 AM'.
    
    Formats supportés:
    - JJ/MM/AA (25/04/25)
    - JJ/MM/AAAA (25/04/2025)
    - JJ-MM-AA (25-04-25)
    - JJ-MM-AAAA (25-04-2025)
    - JJ/MM/AAAA HH:MM (25/04/2025 14:30)
    - JJ/MM/AAAA HH:MM:SS (25/04/2025 14:30:45)
    - AAAA-MM-JJ (2025-04-25)
    - AAAA/MM/JJ (2025/04/25)
    - JJ/MM/AA (avec espaces) ( 25/04/25 )
    - Formats avec points : JJ.MM.AA (25.04.25)
    """
    if not valeur:
        return valeur
    
    from datetime import datetime
    import re
    
    # Nettoyer la valeur : enlever les espaces superflus
    valeur_propre = valeur.strip()
    
    # Liste des formats à essayer (du plus spécifique au plus général)
    formats = [
        "%d/%m/%Y %H:%M:%S",      # 25/04/2025 14:30:45
        "%d/%m/%Y %H:%M",         # 25/04/2025 14:30
        "%d-%m-%Y %H:%M:%S",      # 25-04-2025 14:30:45
        "%d-%m-%Y %H:%M",         # 25-04-2025 14:30
        "%Y-%m-%d %H:%M:%S",      # 2025-04-25 14:30:45
        "%Y-%m-%d %H:%M",         # 2025-04-25 14:30
        "%Y/%m/%d %H:%M:%S",      # 2025/04/25 14:30:45
        "%Y/%m/%d %H:%M",         # 2025/04/25 14:30
        "%d/%m/%Y",               # 25/04/2025
        "%d-%m-%Y",               # 25-04-2025
        "%d.%m.%Y",               # 25.04.2025
        "%Y-%m-%d",               # 2025-04-25
        "%Y/%m/%d",               # 2025/04/25
        "%d/%m/%y",               # 25/04/25
        "%d-%m-%y",               # 25-04-25
        "%d.%m.%y",               # 25.04.25
    ]
    
    # Essayer chaque format
    for fmt in formats:
        try:
            date_obj = datetime.strptime(valeur_propre, fmt)
            return date_obj.strftime("%d/%m/%Y 12:00:00 AM")
        except ValueError:
            continue
    
    # Essayer de détecter et corriger les formats avec séparateurs mixtes
    # Par exemple: 25/04-2025 ou 25-04/2025
    separateurs = ['/', '-', '.']
    for sep in separateurs:
        if sep in valeur_propre:
            # Remplacer tous les séparateurs par le même
            valeur_normalisee = re.sub(r'[/\-\.]', sep, valeur_propre)
            for fmt in ["%d%sep%m%sep%Y", "%d%sep%m%sep%y"]:
                fmt = fmt.replace("%sep", sep)
                try:
                    date_obj = datetime.strptime(valeur_normalisee, fmt)
                    return date_obj.strftime("%d/%m/%Y 12:00:00 AM")
                except ValueError:
                    continue
    
    # Si aucun format ne correspond, retourner la valeur originale
    return valeur


def extraire_metadonnees_rapide(chemin_fichier, debug=False, use_cache=True):
    """Extrait les propriétés sans ouvrir le fichier dans Solid Edge (ultra rapide).
    
    Utilise SolidEdge.FileProperties qui lit les métadonnées directement depuis le fichier
    sans lancer l'interface graphique de Solid Edge.
    
    Args:
        chemin_fichier: Chemin complet du fichier (.asm, .par, .psm, .dft)
        debug: Activer le mode debug
        use_cache: Utiliser le cache pour éviter les doublons
    """
    meta = {
        "designation": "", 
        "revision": "1", 
        "version": "-",
        "auteur": "",
        "date_creation": "",
        "auteur_modif": "",
        "date_modif": ""
    }
    
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
        
        # 1. Auteur = ExtendedSummaryInformation -> "Username"
        try:
            ext_props = prop_reader.Item("ExtendedSummaryInformation")
            for i in range(1, ext_props.Count + 1):
                try:
                    p = ext_props.Item(i)
                    if p.Name == "Username":
                        val = str(p.Value).strip()
                        if val:
                            meta["auteur"] = val
                        break
                except:
                    pass
        except:
            pass

        # 2. Custom -> Désignation, Date de création, version, auteur_modif, date_modif
        noms_version = ["indice de modification", "revision index", "index", "revision", "rev"]
        try:
            custom_props = prop_reader.Item("Custom")

            # Accès direct par nom pour les champs sensibles à la casse
            for nom_champ, cle_meta in [("auteur modif", "auteur_modif"), ("date modif", "date_modif")]:
                try:
                    p = custom_props.Item(nom_champ)
                    val = str(p.Value).strip() if p.Value is not None else ""
                    meta[cle_meta] = normaliser_date(val) if "date" in cle_meta else val
                except:
                    pass

            # Itération par index pour les autres champs
            for i in range(1, custom_props.Count + 1):
                try:
                    prop = custom_props.Item(i)
                    nom_lower = prop.Name.lower()
                    val = str(prop.Value).strip() if prop.Value is not None else ""

                    if nom_lower in ("désignation", "designation", "desig"):
                        meta["designation"] = val
                    elif nom_lower in ("date de création", "date de creation"):
                        meta["date_creation"] = normaliser_date(val)
                    elif any(n in nom_lower for n in noms_version):
                        meta["version"] = val if val.strip() else "-"
                        if debug:
                            print(f"  -> Version: {val if val.strip() else '-'}")
                except:
                    pass
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
    
    meta = {"designation": "", "revision": "1", "version": "-", "auteur": "", "date_creation": "", "auteur_modif": "", "date_modif": ""}

    if debug:
        lister_proprietes(doc_obj)

    noms_version = ["indice de modification", "revision index", "modification index", "index", "revision", "rev"]

    if hasattr(doc_obj, 'Properties'):
        try:
            for prop_set in doc_obj.Properties:
                if not hasattr(prop_set, 'Name'):
                    continue
                # Auteur depuis ExtendedSummaryInformation -> Username
                if prop_set.Name == "ExtendedSummaryInformation":
                    for prop in prop_set:
                        try:
                            if prop.Name == "Username":
                                val = str(prop.Value).strip() if prop.Value else ""
                                if val:
                                    meta["auteur"] = val
                                break
                        except:
                            pass
                # Désignation, dates, version depuis Custom
                elif prop_set.Name == "Custom":
                    # Accès direct par nom pour les champs sensibles à la casse
                    for nom_champ, cle_meta in [("auteur modif", "auteur_modif"), ("date modif", "date_modif")]:
                        try:
                            p = prop_set.Item(nom_champ)
                            val = str(p.Value).strip() if p.Value is not None else ""
                            meta[cle_meta] = normaliser_date(val) if "date" in cle_meta else val
                        except:
                            pass
                    # Itération par index pour les autres champs
                    for prop in prop_set:
                        try:
                            nom_lower = prop.Name.lower()
                            val = str(prop.Value).strip() if prop.Value is not None else ""
                            if nom_lower in ("désignation", "designation", "desig"):
                                meta["designation"] = val
                            elif nom_lower in ("date de création", "date de creation"):
                                meta["date_creation"] = normaliser_date(val)
                            elif any(n in nom_lower for n in noms_version):
                                meta["version"] = val if val.strip() else "-"
                        except:
                            pass
        except:
            pass

    # Mettre en cache si activé
    if use_cache and doc_id:
        g_metadata_cache.set(doc_id, meta.copy())
    
    return meta

def calculer_indices_precedents(version):
    """Calcule indice n-1 et n-2 depuis la version courante (logique alphabétique).
    Ex: B -> (A, -), C -> (B, A), A -> (-, -), - -> (-, -)
    Gère aussi les doubles lettres: AA -> (Z, -), AB -> (AA, Z)
    """
    if not version or version.strip() in ("-", ""):
        return "-", "-"

    def precedent(v):
        v = v.strip().upper()
        if not v:
            return "-"
        # Lettre simple: A->-, B->A, Z->Y
        if len(v) == 1:
            if v == "A":
                return "-"
            return chr(ord(v) - 1)
        # Double lettre: AA->Z, AB->AA, AZ->AY, BA->AZ
        last = v[-1]
        prefix = v[:-1]
        if last == "A":
            # Réduire le préfixe
            new_prefix = precedent(prefix)
            if new_prefix == "-":
                return "Z"  # AA -> Z
            return new_prefix + "Z"
        return prefix + chr(ord(last) - 1)

    n1 = precedent(version)
    n2 = precedent(n1) if n1 != "-" else "-"
    return n1, n2


def determiner_classe(nom_fichier, chemin_complet="", est_projet=False):
    """Détermine la classe PLM en fonction de l'extension du fichier et du chemin.
    
    Les fichiers dans le répertoire "Bibliothèque" sont classés comme PART_PURCH_A.
    """
    if est_projet: return "SUB_ASSY_A"
    ext = os.path.splitext(nom_fichier)[1].lower()
    if ext == '.asm': return "SUB_ASSY_A"
    if ext in ['.par', '.psm']:
        # Vérifier si le fichier est dans un répertoire "Bibliothèque" à n'importe quel niveau
        if chemin_complet:
            # Normaliser le chemin et diviser en composants
            chemin_normalise = os.path.normpath(chemin_complet)
            composants = chemin_normalise.split(os.sep)
            # Vérifier chaque composant du chemin
            for composant in composants:
                if composant.lower() in ["bibliothèque", "bibliotheque", "library"]:
                    return "PART_PURCH_A"
        return "PART_A"
    if ext == '.dft': return "CAD_DRAWING_A"
    return "Folder"

def generer_export_excel(chemin_fichier, dossier_sortie, nom_sortie, dossier_dft=None, mode_recherche="les_deux", type_fichier="asm",
                         callback_log=None, callback_progress=None, check_cancelled=None):
    """
    Fonction moteur principale pour générer l'export PLM.
    
    Args:
        chemin_fichier: Chemin du fichier principal (.asm, .par, .psm)
        dossier_sortie: Dossier de sortie pour le fichier Excel
        nom_sortie: Nom du fichier de sortie
        dossier_dft: Dossier spécifique pour les plans (optionnel)
        mode_recherche: Mode de recherche des plans ("arborescence", "dossier_specifique", "les_deux")
        type_fichier: Type de fichier principal ("asm", "pieces", "les_deux")
        callback_log: Fonction callback(message, type) pour les logs (type: 'info', 'success', 'error', 'warning')
        callback_progress: Fonction callback(valeur, maximum, message) pour la progression
        check_cancelled: Fonction callback() qui retourne True si l'opération doit être annulée
    
    Returns:
        dict: {'chemin_fichier': chemin du fichier généré, 'stats': {'3d': int, '2d': int}}
    """
    # Fonctions utilitaires pour les logs
    def log(message, msg_type='info'):
        if callback_log:
            callback_log(message, msg_type)
        else:
            print(message)
    
    def progress(value, maximum, message):
        if callback_progress:
            callback_progress(value, maximum, message)
    
    def is_cancelled():
        if check_cancelled:
            return check_cancelled()
        return False
    
    try:
        log("=" * 60, 'info')
        log("Début de l'extraction PLM", 'info')
        log("=" * 60, 'info')
        progress(0, 100, "Initialisation...")

        # Callback pour la progression de l'indexation
        def on_index_progress(scanned, found, total):
            if total > 0:
                pct = min(10, int((scanned / total) * 10))
                progress(pct, 100, f"Indexation: {scanned} dossiers scannés, {found} plans trouvés")

        log(f"\n--- Indexation des plans (.dft) [Mode: {mode_recherche}] ---", 'info')
        progress(0, 100, "Indexation des plans...")
        index_plans = indexer_les_plans_projet_entier(chemin_fichier, dossier_dft, mode_recherche, callback_progress=on_index_progress)
        log(f"-> {len(index_plans)} plan(s) détecté(s).", 'info')
        progress(10, 100, f"{len(index_plans)} plans indexés")
        
        log("\nConnexion à Solid Edge...", 'info')
        try:
            app = win32com.client.GetActiveObject("SolidEdge.Application")
            log("Connecté à l'instance existante de Solid Edge.", 'success')
        except pythoncom.com_error:
            log("Démarrage de Solid Edge en arrière-plan...", 'info')
            app = win32com.client.dynamic.Dispatch("SolidEdge.Application")
            app.Visible = False
            log("Solid Edge démarré.", 'success')
        
        app.DisplayAlerts = False

        # Vérifier si le document est déjà ouvert pour éviter l'ouverture en lecture seule
        doc_racine = None
        chemin_fichier_norm = os.path.normcase(chemin_fichier)
        try:
            for doc in app.Documents:
                try:
                    if os.path.normcase(doc.FullName) == chemin_fichier_norm:
                        doc_racine = doc
                        log("Document déjà ouvert, réutilisation de l'instance existante.", 'info')
                        break
                except:
                    pass
        except:
            pass

        if doc_racine is None:
            doc_racine = app.Documents.Open(chemin_fichier)
        log("Document chargé.", 'success')
        
        lignes_excel = []
        compteur_ordre = 1
        liste_plans_a_rajouter = []
        stats = {"3d": 0, "2d": 0}
        
        def get_suffixe_fichier(nom_fichier):
            ext = os.path.splitext(nom_fichier)[1].lower()
            if ext == '.asm': return "(ASM)"
            elif ext in ['.par', '.psm']: return "(PRT)"
            elif ext == '.dft': return "(DRW)"
            return ""
        
        def normaliser_chemin_reseau(chemin):
            """Normalise un chemin en préservant les adresses IP réseau."""
            if not chemin:
                return chemin
            
            # Si c'est un chemin réseau UNC, vérifier si c'est une adresse IP
            if chemin.startswith('\\\\'):
                # Extraire la partie serveur/partage
                parts = chemin[2:].split('\\', 2)
                if len(parts) >= 2:
                    serveur = parts[0]
                    # Si le serveur est une adresse IP, la préserver
                    if serveur.replace('.', '').isdigit():
                        # C'est une adresse IP, ne pas normaliser pour éviter la résolution DNS
                        return chemin.replace('/', '\\')  # Juste corriger les slashs inversés
                    else:
                        # C'est un nom d'hôte, on peut normaliser
                        return os.path.normpath(chemin)
            
            # Pour les chemins locaux, utiliser la normalisation standard
            return os.path.normpath(chemin)

        def ajouter_ligne(niveau, relation, nom_fichier, chemin_complet, classe, qte=1, rev="1", desig="", ver="-", auteur="", date_crea="", auteur_modif="", date_modif=""):
            nonlocal compteur_ordre
            ref_util = os.path.splitext(nom_fichier)[0]
            special_cad = os.path.splitext(nom_fichier)[0]
            suffixe = get_suffixe_fichier(nom_fichier)
            chemin_normalise = normaliser_chemin_reseau(chemin_complet)
            attachement = f"{chemin_normalise}{suffixe}" if suffixe else chemin_normalise
            lignes_excel.append([niveau, relation, compteur_ordre, qte, "", special_cad, classe, ref_util, ver, *calculer_indices_precedents(ver), rev, desig, auteur, date_crea, auteur_modif, date_modif, "", attachement])
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
                    
                    # Essayer plusieurs méthodes pour récupérer le chemin
                    try:
                        path_reel = occ.OccurrenceDocument.FullName
                        nom_reel = os.path.basename(path_reel)
                    except:
                        # Fallback 1: OccurrenceFileName
                        try:
                            path_reel = occ.OccurrenceFileName
                            nom_reel = os.path.basename(path_reel)
                        except:
                            # Fallback 2: FileName
                            try:
                                path_reel = occ.FileName
                                nom_reel = os.path.basename(path_reel)
                            except:
                                # Fallback 3: Utiliser le nom de l'occurrence
                                nom_reel = occ.Name.split(':')[0]
                                log(f"  -> Attention: Chemin non trouvé pour {nom_reel}", 'warning')
                    
                    if nom_reel not in dict_occ:
                        dict_occ[nom_reel] = {"qte": 1, "obj": occ, "chemin": path_reel}
                    else:
                        dict_occ[nom_reel]["qte"] += 1
                except: continue
            
            for nom, data in dict_occ.items():
                stats["3d"] += 1
                classe_3d = determiner_classe(nom, data["chemin"])
                
                # Si pas de chemin, utiliser des valeurs par défaut
                if not data["chemin"]:
                    meta = {"revision": "1", "designation": nom, "version": "-", "auteur": "", "date_creation": "", "auteur_modif": "", "date_modif": ""}
                    log(f"  -> Métadonnées par défaut pour {nom} (chemin manquant)", 'warning')
                else:
                    meta = extraire_metadonnees_rapide(data["chemin"])
                
                ajouter_ligne(niveau, "ComposedOf", nom, data["chemin"], classe_3d, data["qte"], meta["revision"], meta["designation"], meta["version"], meta["auteur"], meta["date_creation"], meta["auteur_modif"], meta["date_modif"])
                
                nom_sans_ext = os.path.splitext(nom)[0].lower()
                if nom_sans_ext in index_plans:
                    chemin_dft = index_plans[nom_sans_ext]
                    nom_dft = os.path.basename(chemin_dft)
                    liste_plans_a_rajouter.append({
                        "dft_nom": nom_dft, "dft_path": chemin_dft,
                        "src_nom": nom, "src_path": data["chemin"],
                        "src_classe": classe_3d, "src_rev": meta["revision"], "src_desig": meta["designation"],
                        "src_ver": meta["version"]
                    })
                    stats["2d"] += 1
                
                if data["obj"].Subassembly:
                    try: explorer_occurrences(data["obj"].OccurrenceDocument.Occurrences, niveau + 1)
                    except: pass
        
        log("\nAnalyse de la structure...", 'info')
        meta_root = extraire_metadonnees(doc_racine)
        nom_root = os.path.basename(doc_racine.FullName)
        
        # Déterminer la classe du fichier racine selon le type
        ext_root = os.path.splitext(nom_root)[1].lower()
        if type_fichier == "asm":
            classe_root = "SUB_ASSY_A"
        elif type_fichier == "pieces":
            classe_root = determiner_classe(nom_root, doc_racine.FullName)
        else:  # "les_deux"
            if ext_root == '.asm':
                classe_root = "SUB_ASSY_A"
            else:
                classe_root = determiner_classe(nom_root, doc_racine.FullName)
        
        ajouter_ligne(0, "", nom_root, doc_racine.FullName, classe_root, 1, meta_root["revision"], meta_root["designation"], meta_root["version"], meta_root["auteur"], meta_root["date_creation"], meta_root["auteur_modif"], meta_root["date_modif"])
        
        nom_root_pur = os.path.splitext(nom_root)[0].lower()
        if nom_root_pur in index_plans:
            path_dft_root = index_plans[nom_root_pur]
            nom_dft_root = os.path.basename(path_dft_root)
            liste_plans_a_rajouter.append({
                "dft_nom": nom_dft_root, "dft_path": path_dft_root,
                "src_nom": nom_root, "src_path": doc_racine.FullName,
                "src_classe": classe_root, "src_rev": meta_root["revision"], "src_desig": meta_root["designation"],
                "src_ver": meta_root["version"]
            })
            stats["2d"] += 1
        
        # Explorer les occurrences seulement si c'est un assemblage
        if ext_root == '.asm' and hasattr(doc_racine, 'Occurrences'):
            explorer_occurrences(doc_racine.Occurrences, 1)
        else:
            # Pour les pièces simples, incrémenter juste les stats 3D
            stats["3d"] += 1
        
        log("\nExtraction métadonnées des plans...", 'info')
        progress(50, 100, "Extraction des métadonnées des plans...")
        
        plans_uniques = {}
        for item in liste_plans_a_rajouter:
            if item["dft_path"] not in plans_uniques:
                plans_uniques[item["dft_path"]] = item
        
        plans_list = list(plans_uniques.values())
        total_plans = len(plans_list)
        plans_deja_traites = set()
        
        for idx, item in enumerate(plans_list):
            if is_cancelled():
                break
            
            if idx % 5 == 0 and total_plans > 0:
                progress_pct = 50 + int((idx / total_plans) * 20)
                progress(progress_pct, 100, f"Plan {idx+1}/{total_plans}: {item['dft_nom'][:30]}...")
            
            if item["dft_path"] not in plans_deja_traites:
                meta_dft = extraire_metadonnees_rapide(item["dft_path"])
                ajouter_ligne(0, "", item["dft_nom"], item["dft_path"], "CAD_DRAWING_A", 1, meta_dft["revision"], item["src_desig"], meta_dft["version"], meta_dft["auteur"], meta_dft["date_creation"], meta_dft["auteur_modif"], meta_dft["date_modif"])
                ajouter_ligne(1, "Drawing", item["src_nom"], item["src_path"], item["src_classe"], 1, item["src_rev"], item["src_desig"], item["src_ver"], "", "", "", "")
                plans_deja_traites.add(item["dft_path"])
        
        log(f"Analyse terminée : {stats['3d']} fichiers 3D, {stats['2d']} plans", 'success')
        progress(70, 100, f"Analyse terminée: {stats['3d']} fichiers 3D, {stats['2d']} plans")
        
        log("\nGénération du fichier Excel...", 'info')
        progress(75, 100, "Génération du fichier Excel...")
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Structure"
        
        headers = ["Level", "Relationship", "ordre", "quantite", "repere", "SpecialCAD", "Class", "ref_utilisat", "version", "indice_1", "indice_2", "revision", "designation", "cus_createur", "cus_date_crea", "user_version_1", "date_version_1", "dia_se", "Attachments"]
        ws.append(headers)
        
        header_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
        orange_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
        
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = header_fill if cell.column < 13 else orange_fill
            cell.alignment = Alignment(horizontal="left")
        
        progress(80, 100, "Écriture des données...")
        for l in lignes_excel: ws.append(l)
        
        progress(90, 100, "Calcul des largeurs de colonnes...")
        columns = list(ws.columns)
        
        def calc_column_width(col_data):
            col_cells, idx = col_data
            max_length = 0
            for cell in col_cells:
                try: max_length = max(max_length, len(str(cell.value)))
                except: pass
            return (idx, max_length + 2)
        
        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = [executor.submit(calc_column_width, (col, i)) for i, col in enumerate(columns)]
            for future in futures:
                idx, width = future.result()
                ws.column_dimensions[columns[idx][0].column_letter].width = width
        
        if not nom_sortie.endswith('.xlsx'):
            nom_sortie += '.xlsx'
        
        chemin_complet = os.path.join(dossier_sortie, nom_sortie)
        progress(95, 100, "Sauvegarde du fichier...")
        wb.save(chemin_complet)
        
        progress(100, 100, "Terminé!")
        log(f"\nFichier généré : {chemin_complet}", 'success')
        log("=" * 60, 'info')
        log("Extraction terminée avec succès !", 'success')
        
        return {'chemin_fichier': chemin_complet, 'stats': stats}
        
    except Exception as e:
        log(f"\nErreur : {e}", 'error')
        log("=" * 60, 'error')
        raise

def lancer_extraction_plm():
    try:
        chemin_asm = demander_fichier_asm()
        if not chemin_asm: return

        dossier_sortie = os.path.dirname(chemin_asm)
        nom_sortie = f"Export_PLM_{int(time.time())}.xlsx"
        
        # Utiliser le moteur centralisé
        resultat = generer_export_excel(
            chemin_asm=chemin_asm,
            dossier_sortie=dossier_sortie,
            nom_sortie=nom_sortie,
            dossier_dft=None,
            mode_recherche="les_deux"
        )
        
        print(f"\nFichier généré : {resultat['chemin_fichier']}")
        print(f"3D: {resultat['stats']['3d']} | Plans: {resultat['stats']['2d']}")

    except Exception as e:
        print(f"\nErreur : {e}")

if __name__ == "__main__":
    lancer_extraction_plm()