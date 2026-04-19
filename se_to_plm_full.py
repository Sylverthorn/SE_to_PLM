"""
Solid Edge → PDM/PLM BOM Extractor (version fusionnée)
=======================================================
Parcourt récursivement un assemblage Solid Edge (.asm) et génère un fichier
Excel prêt à importer dans le PLM Andros (format structureData).

Logique PDM  : traversée hiérarchique + Level / Relationship / Class
               (issue de extract_relational_links.py)
Colonnes PLM : format complet Andros + mapping propriétés CAO → PLM
               + propriétés physiques + UI Tkinter
               (issue de se_to_plm.py)

Prérequis :
  - Windows + Solid Edge installé
  - pip install pywin32 openpyxl

Usage :
  python se_to_plm_full.py
"""

import os
import sys
import datetime
import traceback
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

try:
    import win32com.client
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("[WARN] pywin32 introuvable : mode démo activé.")


# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

SE_APP_ID = "SolidEdge.Application"
EXT_ASM   = {".asm"}
EXT_PART  = {".par", ".psm"}
EXT_DRW   = {".dft"}
EXT_ALL   = EXT_ASM | EXT_PART | EXT_DRW

# Colonnes cibles pour Andros (ordre conservé depuis se_to_plm.py)
COLUMNS = [
    "Level", "Relationship", "ordre", "quantite", "repere", "SpecialCAD",
    "erp_bomenddat", "erp_bomsho", "erp_bomstrdat", "erp_cpnope", "erp_cpntyp",
    "erp_scoflg", "unite", "Class", "ref_utilisat", "version", "revision",
    "designation", "statut", "APLMC_3D_reference", "APLMC_date_CAD_change",
    "approbateur", "createur_ori", "date_approbation", "date_creation_ori",
    "date_verification", "date_version_1", "date_version_2", "format",
    "indice_1", "indice_2", "libelle_modif_1", "libelle_modif_2",
    "type_objet", "user_version_1", "user_version_2", "verificateur",
    "APLMC_item_mgt", "APLMC_my3Dviewer_ID", "APLMC_my3Dviewer_timestamp",
    "APLMC_pivot", "conditionnement", "erp_expdat", "erp_itmref", "erp_manage",
    "erp_stu", "erp_tclcod", "famille_objet", "poids", "rohs", "surface",
    "volume3D", "densite", "dim1", "dim2", "dim3", "matiere", "Attachments",
]

# Mapping propriétés Solid Edge → colonnes PLM
SE_PROP_MAP = {
    "Title":        "designation",  "Désignation":   "designation",  "Designation":  "designation",
    "PartNumber":   "ref_utilisat", "Part Number":   "ref_utilisat", "Référence":    "ref_utilisat",
    "Reference":    "ref_utilisat",
    "Revision":     "version",      "Révision":      "version",      "Version":      "version",
    "Material":     "matiere",      "Matière":       "matiere",      "Matiere":      "matiere",
    "Author":       "createur_ori", "Auteur":        "createur_ori",
    "Status":       "statut",       "Statut":        "statut",
    "Weight":       "poids",        "Masse":         "poids",
    "Surface":      "surface",
    "Volume":       "volume3D",
    "Density":      "densite",      "Densité":       "densite",
    "Format":       "format",
    "ERP_Manage":   "erp_manage",   "ERP_Stu":       "erp_stu",
    "ERP_CpnTyp":   "erp_cpntyp",  "ERP_ScoFlg":    "erp_scoflg",
    "APLMC_ItemMgt":"APLMC_item_mgt","RoHS":         "rohs",
}

# Valeurs par défaut pour les champs critiques
SE_DEFAULT_VALUES = {
    "statut":        "Valide",
    "erp_manage":    "AUDROS",
    "erp_stu":       "Unité [UN]",
    "erp_cpntyp":    "Normal [1]",
    "erp_scoflg":    "Interne [1]",
    "APLMC_item_mgt":"[0] NON",
    "rohs":          "NON",
    "version":       "A",
}

# Sous-dossiers courants pour la recherche de fichiers .dft
DRAWING_SUBDIRS = ["DESSINS", "MISES EN PLAN", "PLANS", "DRAWINGS", "DRAFTS"]


# ---------------------------------------------------------------------------
# Lecture Solid Edge (COM)
# ---------------------------------------------------------------------------

class SolidEdgeReader:
    """
    Gère la connexion à Solid Edge via COM et lit les propriétés d'un document.
    Tente d'abord de se connecter à une instance existante, sinon en crée une.
    """

    def __init__(self, visible: bool = False):
        self.app = None
        self.visible = visible
        self._app_started_by_us = False

    def start(self):
        pythoncom.CoInitialize()
        try:
            # Réutiliser une instance déjà ouverte
            self.app = win32com.client.GetActiveObject(SE_APP_ID)
            self._app_started_by_us = False
            print("[SE] Connecté à l'instance existante de Solid Edge.")
        except Exception:
            self.app = win32com.client.dynamic.Dispatch(SE_APP_ID)
            self.app.Visible = self.visible
            self._app_started_by_us = True
            print("[SE] Nouvelle instance de Solid Edge démarrée.")
        try:
            self.app.DisplayAlerts = False
        except Exception:
            pass

    def stop(self):
        if self.app:
            if self._app_started_by_us:
                try:
                    self.app.Quit()
                    print("[SE] Solid Edge fermé.")
                except Exception:
                    pass
            self.app = None
        pythoncom.CoUninitialize()

    def close_all_documents(self):
        """Ferme tous les documents ouverts (utile avant de commencer le scan)."""
        if not self.app:
            return
        try:
            docs = self.app.Documents
            while docs.Count > 0:
                try:
                    docs.Item(1).Close(False)
                except Exception:
                    break
            print("[SE] Tous les documents fermés.")
        except Exception as e:
            print(f"[WARN] Fermeture documents : {e}")

    # ------------------------------------------------------------------
    # Lecture des propriétés d'un document déjà ouvert
    # ------------------------------------------------------------------

    def read_doc_properties(self, doc) -> dict:
        """
        Extrait les propriétés d'un document COM déjà ouvert et les retourne
        sous forme de dict {colonne_plm: valeur}.
        """
        result = dict(SE_DEFAULT_VALUES)  # Valeurs par défaut
        ext = os.path.splitext(doc.FullName)[1].lower()

        props_raw = {}
        try:
            prop_sets = doc.Properties

            # SummaryInformation
            try:
                summary = prop_sets.Item("SummaryInformation")
                for prop in ["Title", "Author", "Subject", "Keywords"]:
                    try:
                        props_raw[prop] = str(summary.Item(prop).Value).strip()
                    except Exception:
                        pass
            except Exception:
                pass

            # ProjectInformation
            try:
                project = prop_sets.Item("ProjectInformation")
                for prop in ["PartNumber", "Revision", "DocumentNumber"]:
                    try:
                        props_raw[prop] = str(project.Item(prop).Value).strip()
                    except Exception:
                        pass
            except Exception:
                pass

            # Propriétés personnalisées
            try:
                custom = prop_sets.Item("Custom")
                for i in range(custom.Count):
                    p = custom.Item(i + 1)
                    props_raw[p.Name] = str(p.Value).strip()
            except Exception:
                pass

            # Itération générique sur tous les PropertySets (fallback)
            try:
                for i in range(1, prop_sets.Count + 1):
                    try:
                        ps = prop_sets.Item(i)
                        for j in range(1, ps.Count + 1):
                            try:
                                p = ps.Item(j)
                                if p.Name not in props_raw:
                                    props_raw[p.Name] = str(p.Value).strip()
                            except Exception:
                                pass
                    except Exception:
                        pass
            except Exception:
                pass

        except Exception:
            pass

        # Propriétés physiques (pièces uniquement)
        if ext in EXT_PART:
            try:
                phys = doc.PhysicalProperties
                phys.Update()
                result.update({
                    "poids":    round(phys.Mass,    6),
                    "volume3D": round(phys.Volume,  5),
                    "surface":  round(phys.Area,    5),
                    "densite":  round(phys.Density, 6),
                })
            except Exception:
                pass

        # Application du mapping SE → PLM
        for se_name, plm_col in SE_PROP_MAP.items():
            if props_raw.get(se_name) and plm_col not in result:
                result[plm_col] = props_raw[se_name]

        return result


# ---------------------------------------------------------------------------
# Extraction hiérarchique PDM
# ---------------------------------------------------------------------------

class PDMExtractor:
    """
    Parcourt récursivement la structure d'un assemblage Solid Edge.
    Produit une liste d'entrées avec Level / Relationship / Class et toutes
    les colonnes PLM Andros.
    """

    def __init__(self, se_reader: SolidEdgeReader):
        self.reader = se_reader
        self.rows: list[dict] = []
        self._order = 0
        self._now_str = datetime.datetime.now().strftime("%d/%m/%Y 12:00:00 AM")

    # ------------------------------------------------------------------
    # Point d'entrée public
    # ------------------------------------------------------------------

    def extract(self, asm_path: str) -> list[dict]:
        self.rows = []
        self._order = 0
        print(f"\n[PDM] Ouverture de l'assemblage racine : {asm_path}")
        self.reader.close_all_documents()

        try:
            assembly = self.reader.app.Documents.Open(asm_path)
            self._traverse_assembly(assembly, level=0)
            assembly.Close(False)
        except Exception as e:
            print(f"[ERR] Impossible d'ouvrir {asm_path} : {e}")

        return self.rows

    # ------------------------------------------------------------------
    # Traversée récursive
    # ------------------------------------------------------------------

    def _traverse_assembly(self, assembly, level: int):
        """Traite l'assemblage lui-même puis ses occurrences (récursif)."""
        try:
            # --- L'assemblage / sous-assemblage courant ---
            self._add_entry_from_doc(assembly, level, "ComposedOf", is_root_doc=True)

            # --- Fichier de dessin associé ---
            drawing_path = self._find_associated_drawing(assembly)
            if drawing_path:
                self._add_drawing_entry(drawing_path, level)

            # --- Occurrences (composants enfants) ---
            occurrences = assembly.Occurrences
            for i in range(1, occurrences.Count + 1):
                try:
                    occ = occurrences.Item(i)
                    self._order += 1

                    # Filtrage BOM
                    try:
                        if not occ.IncludeInBom:
                            continue
                    except Exception:
                        pass

                    self._add_entry_from_occurrence(occ, level + 1, "ComposedOf")

                    # Récursion sur les sous-assemblages
                    try:
                        sub_doc = occ.OccurrenceDocument
                        if sub_doc.Type == 3:   # igAssemblyDocument
                            self._traverse_assembly(sub_doc, level + 1)
                    except Exception:
                        pass

                except Exception as e:
                    print(f"  [WARN] Occurrence {i} ignorée : {e}")

        except Exception as e:
            print(f"[ERR] Traversée assemblage niveau {level} : {e}")

    # ------------------------------------------------------------------
    # Création d'une entrée PDM/PLM
    # ------------------------------------------------------------------

    def _base_entry(self, level: int, relationship: str, class_type: str,
                    name: str, doc=None) -> dict:
        """Construit un dict avec toutes les colonnes COLUMNS initialisées."""
        entry = {col: "" for col in COLUMNS}

        # Champs structurels PDM
        entry["Level"]        = level
        entry["Relationship"] = relationship
        entry["Class"]        = class_type
        entry["ordre"]        = self._order
        entry["quantite"]     = 1
        entry["ref_utilisat"] = name.upper()
        entry["SpecialCAD"]   = name.lower()
        entry["date_creation_ori"] = self._now_str
        entry["APLMC_pivot"]  = name.upper()

        # Valeurs par défaut PLM
        entry.update(SE_DEFAULT_VALUES)

        # Lecture des propriétés COM si disponible
        if doc is not None:
            try:
                se_props = self.reader.read_doc_properties(doc)
                for col, val in se_props.items():
                    if col in entry:
                        entry[col] = val
                # S'assurer que ref_utilisat est renseigné
                if not entry.get("ref_utilisat"):
                    entry["ref_utilisat"] = name.upper()
            except Exception as e:
                print(f"  [WARN] Lecture propriétés '{name}' : {e}")

        return entry

    def _doc_type_to_class(self, doc_type: int) -> str:
        """Convertit le type COM Solid Edge en classe PLM."""
        mapping = {
            1: "PART_A",      # igPartDocument      (.par)
            3: "SUB_ASSY_A",  # igAssemblyDocument  (.asm)
            4: "PART_A",      # igSheetMetalDocument(.psm)
            5: "SUB_ASSY_A",  # igWeldmentDocument  (.pwd)
            2: "CAD_DRAWING_A",# igDraftDocument    (.dft)
        }
        return mapping.get(doc_type, f"UNKNOWN_{doc_type}")

    def _add_entry_from_doc(self, doc, level: int, relationship: str,
                             is_root_doc: bool = False):
        """Ajoute une entrée à partir d'un document COM directement ouvert."""
        try:
            name = os.path.splitext(os.path.basename(doc.FullName))[0]
            class_type = self._doc_type_to_class(doc.Type)
            entry = self._base_entry(level, relationship, class_type, name, doc)

            # Attachment
            fname = os.path.basename(doc.FullName)
            ext = os.path.splitext(fname)[1].lstrip(".")
            entry["Attachments"] = (
                f"[ROOT]\\attachments_vplm\\{name.lower()}\\{fname}({ext.upper()})"
            )
            self.rows.append(entry)
            print(f"  {'  ' * level}[+] {class_type} | L{level} | {name}")
        except Exception as e:
            print(f"  [ERR] add_entry_from_doc : {e}")

    def _add_entry_from_occurrence(self, occurrence, level: int, relationship: str):
        """Ajoute une entrée à partir d'une occurrence (composant enfant)."""
        try:
            name = occurrence.Name.split(':')[0]

            try:
                doc = occurrence.OccurrenceDocument
                doc_type = doc.Type
            except Exception:
                print(f"  [!] Document inaccessible pour l'occurrence : {name}")
                return

            class_type = self._doc_type_to_class(doc_type)
            entry = self._base_entry(level, relationship, class_type, name, doc)

            fname = os.path.basename(doc.FullName)
            ext = os.path.splitext(fname)[1].lstrip(".")
            entry["Attachments"] = (
                f"[ROOT]\\attachments_vplm\\{name.lower()}\\{fname}({ext.upper()})"
            )
            self.rows.append(entry)
            print(f"  {'  ' * level}[+] {class_type} | L{level} | {name}")
        except Exception as e:
            print(f"  [ERR] add_entry_from_occurrence '{getattr(occurrence, 'Name', '?')}' : {e}")

    def _add_drawing_entry(self, drawing_path: str, level: int):
        """Ajoute une entrée de type Drawing (relation documentaire)."""
        try:
            name = os.path.splitext(os.path.basename(drawing_path))[0]
            entry = self._base_entry(level, "Drawing", "CAD_DRAWING_A", name, doc=None)
            fname = os.path.basename(drawing_path)
            entry["Attachments"] = (
                f"[ROOT]\\attachments_vplm\\{name.lower()}\\{fname}(DFT)"
            )
            entry["type_objet"] = "Plan standard"
            self.rows.append(entry)
            print(f"  {'  ' * level}[D] CAD_DRAWING_A | L{level} | {name}")
        except Exception as e:
            print(f"  [ERR] add_drawing_entry : {e}")

    # ------------------------------------------------------------------
    # Recherche du fichier .dft associé
    # ------------------------------------------------------------------

    def _find_associated_drawing(self, doc) -> str | None:
        """Cherche le fichier .dft associé à un document dans les dossiers courants."""
        try:
            doc_path = doc.FullName
            doc_dir  = os.path.dirname(doc_path)
            doc_name = os.path.splitext(os.path.basename(doc_path))[0]

            search_dirs = [doc_dir]
            for subdir in DRAWING_SUBDIRS:
                candidate = os.path.join(doc_dir, subdir)
                if os.path.isdir(candidate):
                    search_dirs.append(candidate)

            for search_dir in search_dirs:
                for fname in os.listdir(search_dir):
                    fn, fext = os.path.splitext(fname)
                    if fext.lower() != ".dft":
                        continue
                    if fn == doc_name:
                        print(f"  [DFT] Dessin trouvé (exact) : {fname}")
                        return os.path.join(search_dir, fname)
                    if doc_name in fn or fn in doc_name:
                        print(f"  [DFT] Dessin trouvé (partiel) : {fname}")
                        return os.path.join(search_dir, fname)
        except Exception as e:
            print(f"  [WARN] Recherche dessin : {e}")
        return None

    # ------------------------------------------------------------------
    # Agrégation des quantités (optionnel)
    # ------------------------------------------------------------------

    def aggregate_quantities(self) -> list[dict]:
        """Cumule les quantités pour les mêmes ref_utilisat / Level / Class."""
        aggregated: dict = {}
        for item in self.rows:
            key = (item["Level"], item["ref_utilisat"], item["Class"], item["Relationship"])
            if key in aggregated:
                aggregated[key]["quantite"] += 1
            else:
                aggregated[key] = item.copy()
        return list(aggregated.values())


# ---------------------------------------------------------------------------
# Export Excel
# ---------------------------------------------------------------------------

def build_excel(rows: list[dict], output_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Structure PDM"

    header_font  = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    header_fill  = PatternFill("solid", start_color="2F4F8F")
    cell_border  = Border(
        left=Side(style="thin", color="C0C0C0"),
        right=Side(style="thin", color="C0C0C0"),
        top=Side(style="thin", color="C0C0C0"),
        bottom=Side(style="thin", color="C0C0C0"),
    )
    alt_fill = PatternFill("solid", start_color="EEF2F8")

    # En-têtes
    for col_idx, col_name in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = cell_border
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    # Données
    col_widths = {"designation": 40, "Attachments": 55, "date_creation_ori": 22}
    for row_idx, row_data in enumerate(rows, start=2):
        fill = alt_fill if row_idx % 2 == 0 else None
        for col_idx, col_name in enumerate(COLUMNS, start=1):
            val = row_data.get(col_name, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font      = Font(name="Arial", size=9)
            cell.border    = cell_border
            cell.alignment = Alignment(vertical="center")
            if fill:
                cell.fill = fill

    # Largeurs
    for col_idx, col_name in enumerate(COLUMNS, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = (
            col_widths.get(col_name, 14)
        )

    wb.save(output_path)
    print(f"[EXCEL] Fichier sauvegardé : {output_path}")


# ---------------------------------------------------------------------------
# Mode démo (sans Solid Edge)
# ---------------------------------------------------------------------------

def demo_rows() -> list[dict]:
    """Retourne quelques lignes fictives pour tester l'UI sans Solid Edge."""
    now = datetime.datetime.now().strftime("%d/%m/%Y 12:00:00 AM")
    samples = [
        ("GB2100J",  "SUB_ASSY_A",  "ComposedOf", 0, "Canal sup. TB Ø47-16.5",  0,     0),
        ("GB1331J",  "PART_A",      "ComposedOf", 1, "Anneau de canal taraudé", 0.12, 800),
        ("GB21062J", "PART_A",      "ComposedOf", 1, "Tige Ø6 rayon 67.5",     0.05, 320),
        ("GB2100J",  "CAD_DRAWING_A","Drawing",   0, "Plan canal sup. TB",      0,     0),
    ]
    rows = []
    for ref, cls, rel, lvl, desg, poids, vol in samples:
        entry = {col: "" for col in COLUMNS}
        entry.update(SE_DEFAULT_VALUES)
        entry["Level"]          = lvl
        entry["Relationship"]   = rel
        entry["Class"]          = cls
        entry["ref_utilisat"]   = ref
        entry["designation"]    = desg
        entry["quantite"]       = 1
        entry["poids"]          = poids or ""
        entry["volume3D"]       = vol or ""
        entry["date_creation_ori"] = now
        entry["APLMC_pivot"]    = ref
        rows.append(entry)
    return rows


# ---------------------------------------------------------------------------
# Interface Tkinter
# ---------------------------------------------------------------------------

class ConsoleRedirector:
    def __init__(self, widget):
        self.widget = widget

    def write(self, text):
        self.widget.configure(state="normal")
        self.widget.insert(tk.END, text)
        self.widget.see(tk.END)
        self.widget.configure(state="disabled")
        self.widget.update_idletasks()

    def flush(self):
        pass


class SEExportApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Export SE → Andros PLM (PDM complet)")
        self.root.geometry("680x590")

        self.export_dir = os.path.join(os.path.expanduser("~"), "Documents", "Exports_PLM")
        os.makedirs(self.export_dir, exist_ok=True)

        self.var_asm_path  = tk.StringVar()
        self.var_out_file  = tk.StringVar(value="export_plm.xlsx")
        self.var_visible   = tk.BooleanVar(value=False)
        self.var_demo      = tk.BooleanVar(value=False)
        self.var_aggregate = tk.BooleanVar(value=False)

        self._build_ui()

    def _build_ui(self):
        frm = ttk.LabelFrame(self.root, text="Configuration", padding=15)
        frm.pack(fill=tk.X, padx=10, pady=10)

        # Ligne 0 : fichier .asm
        ttk.Label(frm, text="Fichier assemblage (.asm) :").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(frm, textvariable=self.var_asm_path, width=45).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(frm, text="Parcourir", command=self._browse_asm).grid(row=0, column=2, padx=5)

        # Ligne 1 : nom fichier Excel
        ttk.Label(frm, text="Fichier Excel de sortie :").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(frm, textvariable=self.var_out_file, width=45).grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)

        ttk.Label(frm, text=f"Dossier cible : {self.export_dir}",
                  font=("", 8, "italic")).grid(row=2, column=1, sticky=tk.W, padx=5)

        # Options
        opt = ttk.Frame(frm)
        opt.grid(row=3, column=0, columnspan=3, pady=10, sticky=tk.W)
        ttk.Checkbutton(opt, text="Rendre Solid Edge visible",  variable=self.var_visible  ).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Checkbutton(opt, text="Cumuler les quantités",      variable=self.var_aggregate).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Checkbutton(opt, text="Mode test (sans API SE)",    variable=self.var_demo     ).pack(side=tk.LEFT)

        # Bouton lancer
        self.btn_run = ttk.Button(self.root, text="Lancer l'extraction", command=self._run_thread)
        self.btn_run.pack(pady=5)

        # Console
        console_frm = ttk.LabelFrame(self.root, text="Logs", padding=5)
        console_frm.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.console = scrolledtext.ScrolledText(
            console_frm, state="disabled",
            bg="#1e1e1e", fg="#d4d4d4", font=("Consolas", 9)
        )
        self.console.pack(fill=tk.BOTH, expand=True)
        sys.stdout = ConsoleRedirector(self.console)

    def _browse_asm(self):
        path = filedialog.askopenfilename(
            title="Sélectionner le fichier d'assemblage",
            filetypes=[("Solid Edge Assembly", "*.asm"), ("Tous les fichiers", "*.*")]
        )
        if path:
            self.var_asm_path.set(path)
            name = os.path.splitext(os.path.basename(path))[0]
            self.var_out_file.set(f"{name}_PLM.xlsx")

    def _run_thread(self):
        asm_path = self.var_asm_path.get().strip('"')
        if not self.var_demo.get() and not os.path.isfile(asm_path):
            messagebox.showwarning("Erreur", "Sélectionnez un fichier .asm valide.")
            return

        self.btn_run.configure(state="disabled", text="Travail en cours…")
        self.console.configure(state="normal")
        self.console.delete(1.0, tk.END)
        self.console.configure(state="disabled")
        threading.Thread(target=self._process, daemon=True).start()

    def _process(self):
        asm_path = self.var_asm_path.get().strip('"')
        filename = self.var_out_file.get()
        if not filename.endswith(".xlsx"):
            filename += ".xlsx"
        out_path = os.path.join(self.export_dir, filename)

        se_reader = None
        try:
            if self.var_demo.get() or not WIN32_AVAILABLE:
                print("[TEST] Mode démo — données fictives.")
                rows = demo_rows()
            else:
                se_reader = SolidEdgeReader(visible=self.var_visible.get())
                se_reader.start()

                extractor = PDMExtractor(se_reader)
                extractor.extract(asm_path)

                rows = (extractor.aggregate_quantities()
                        if self.var_aggregate.get()
                        else extractor.rows)

            if not rows:
                print("[WARN] Aucune donnée extraite.")
                return

            print(f"\n[INFO] {len(rows)} lignes à exporter.")
            build_excel(rows, out_path)
            print("\n[OK] Export terminé avec succès.")
            messagebox.showinfo("Terminé", f"Export réussi !\n{out_path}")

        except Exception as e:
            print(f"\n[CRITICAL] {e}")
            traceback.print_exc()
            messagebox.showerror("Erreur", str(e))
        finally:
            if se_reader:
                se_reader.stop()
            self.root.after(0, lambda: self.btn_run.configure(
                state="normal", text="Lancer l'extraction"))


# ---------------------------------------------------------------------------
# Point d'entrée
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    root = tk.Tk()
    SEExportApp(root)
    root.mainloop()
