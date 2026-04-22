import sys
import os
import time
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog, QMessageBox)
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QTextCursor, QColor

from se_to_plm import indexer_les_plans_projet_entier, extraire_metadonnees, determiner_classe
import win32com.client
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

class ExtractionThread(QThread):
    log_signal = pyqtSignal(str, str)
    finished_signal = pyqtSignal()
    
    def __init__(self, chemin_asm, dossier_sortie, nom_sortie, dossier_dft=None):
        super().__init__()
        self.chemin_asm = chemin_asm
        self.dossier_sortie = dossier_sortie
        self.nom_sortie = nom_sortie
        self.dossier_dft = dossier_dft
    
    def run(self):
        try:
            self.log_signal.emit("=" * 60, 'info')
            self.log_signal.emit("Début de l'extraction PLM", 'info')
            self.log_signal.emit("=" * 60, 'info')
            
            self.log_signal.emit("\n--- Indexation des plans (.dft) ---", 'info')
            index_plans = indexer_les_plans_projet_entier(self.chemin_asm, self.dossier_dft)
            self.log_signal.emit(f"-> {len(index_plans)} plan(s) détecté(s).", 'info')
            
            self.log_signal.emit("\nOuverture de Solid Edge...", 'info')
            app = win32com.client.dynamic.Dispatch("SolidEdge.Application")
            app.Visible = False
            doc_racine = app.Documents.Open(self.chemin_asm)
            time.sleep(3)
            self.log_signal.emit("Solid Edge connecté.", 'success')
            
            lignes_excel = []
            compteur_ordre = 1
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
                    
                    # On ajoute uniquement le 3D dans l'arbre principal
                    ajouter_ligne(niveau, "ComposedOf", nom, data["chemin"], classe_3d, data["qte"], meta["revision"], meta["designation"], meta["version"])
                    
                    nom_sans_ext = os.path.splitext(nom)[0].lower()
                    if nom_sans_ext in index_plans:
                        chemin_dft = index_plans[nom_sans_ext]
                        nom_dft = os.path.basename(chemin_dft)
                        
                        # On stocke pour l'ajouter à la fin du fichier
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
            
            self.log_signal.emit("\nAnalyse de la structure...", 'info')
            
            meta_root = extraire_metadonnees(doc_racine)
            nom_root = os.path.basename(doc_racine.FullName)
            ajouter_ligne(0, "", nom_root, doc_racine.FullName, "SUB_ASSY_A", 1, meta_root["revision"], meta_root["designation"], meta_root["version"])
            
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
            
            explorer_occurrences(doc_racine.Occurrences, 1)
            
            # --- A LA FIN : AJOUT DES PLANS EN LEVEL 0 ET LEURS 3D EN LEVEL 1 ---
            self.log_signal.emit("\nExtraction métadonnées des plans...", 'info')
            plans_deja_traites = set()
            for item in liste_plans_a_rajouter:
                if item["dft_path"] not in plans_deja_traites:
                    # Ouvrir le fichier DFT pour extraire ses métadonnées
                    meta_dft = {"designation": "", "revision": "1", "version": "-"}
                    try:
                        doc_dft = app.Documents.Open(item["dft_path"])
                        # Debug: activer pour le premier fichier
                        if len(plans_deja_traites) == 0:
                            meta_dft = extraire_metadonnees(doc_dft, debug=True)
                        else:
                            meta_dft = extraire_metadonnees(doc_dft)
                        doc_dft.Close()
                    except Exception as e:
                        self.log_signal.emit(f"  Erreur lecture {item['dft_nom']}: {e}", 'warning')
                    
                    ajouter_ligne(0, "", item["dft_nom"], item["dft_path"], "CAD_DRAWING_A", 1, meta_dft["revision"], meta_dft["designation"], meta_dft["version"])
                    ajouter_ligne(1, "Drawing", item["src_nom"], item["src_path"], item["src_classe"], 1, item["src_rev"], item["src_desig"], item["src_ver"])
                    plans_deja_traites.add(item["dft_path"])
            
            self.log_signal.emit(f"Analyse terminée : {stats['3d']} fichiers 3D, {stats['2d']} plans", 'success')
            
            self.log_signal.emit("\nGénération du fichier Excel...", 'info')
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
            
            nom_sortie = self.nom_sortie
            if not nom_sortie.endswith('.xlsx'):
                nom_sortie += '.xlsx'
            
            chemin_complet = os.path.join(self.dossier_sortie, nom_sortie)
            wb.save(chemin_complet)
            
            self.log_signal.emit(f"\nFichier généré : {chemin_complet}", 'success')
            self.log_signal.emit("=" * 60, 'info')
            self.log_signal.emit("Extraction terminée avec succès !", 'success')
            
        except Exception as e:
            self.log_signal.emit(f"\nErreur : {e}", 'error')
            self.log_signal.emit("=" * 60, 'error')
        
        finally:
            self.finished_signal.emit()

class PLMExtractorGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Extracteur PLM - Solid Edge")
        self.setGeometry(100, 100, 800, 600)
        self.extraction_en_cours = False
        
        self.appliquer_style()
        self.creer_interface()
    
    def appliquer_style(self):
        chemin_style = os.path.join(os.path.dirname(__file__), 'style.qss')
        if os.path.exists(chemin_style):
            with open(chemin_style, 'r', encoding='utf-8') as f:
                self.setStyleSheet(f.read())
        
    def creer_interface(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(10)
        
        asm_layout = QHBoxLayout()
        asm_layout.addWidget(QLabel("Fichier ASM :"))
        self.chemin_asm_edit = QLineEdit()
        self.chemin_asm_edit.setReadOnly(True)
        asm_layout.addWidget(self.chemin_asm_edit)
        btn_parcourir = QPushButton("Parcourir...")
        btn_parcourir.clicked.connect(self.choisir_fichier_asm)
        asm_layout.addWidget(btn_parcourir)
        main_layout.addLayout(asm_layout)
        
        dft_layout = QHBoxLayout()
        dft_layout.addWidget(QLabel("Dossier plans (.dft) :"))
        self.dossier_dft_edit = QLineEdit()
        self.dossier_dft_edit.setReadOnly(True)
        self.dossier_dft_edit.setPlaceholderText("Optionnel")
        dft_layout.addWidget(self.dossier_dft_edit)
        btn_parcourir_dft = QPushButton("Parcourir...")
        btn_parcourir_dft.clicked.connect(self.choisir_dossier_dft)
        dft_layout.addWidget(btn_parcourir_dft)
        main_layout.addLayout(dft_layout)
        
        sortie_layout = QHBoxLayout()
        sortie_layout.addWidget(QLabel("Nom de sortie :"))
        self.nom_sortie_edit = QLineEdit(f"Export_PLM_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        sortie_layout.addWidget(self.nom_sortie_edit)
        main_layout.addLayout(sortie_layout)
        
        self.dossier_sortie = os.path.join(os.path.expanduser("~"), "Documents", "Exports_PLM")
        lbl_dossier = QLabel(f"Dossier : {self.dossier_sortie}")
        lbl_dossier.setStyleSheet("color: gray; font-style: italic;")
        main_layout.addWidget(lbl_dossier)
        
        self.btn_extraire = QPushButton("Lancer l'extraction")
        self.btn_extraire.clicked.connect(self.lancer_extraction)
        self.btn_extraire.setMinimumHeight(40)
        main_layout.addWidget(self.btn_extraire)
        
        main_layout.addWidget(QLabel("Console de progression :"))
        self.console = QTextEdit()
        self.console.setReadOnly(True)
        main_layout.addWidget(self.console)
    
    def choisir_fichier_asm(self):
        chemin, _ = QFileDialog.getOpenFileName(
            self,
            "Sélectionnez l'assemblage principal (.asm)",
            "",
            "Assemblage Solid Edge (*.asm)"
        )
        if chemin:
            self.chemin_asm_edit.setText(chemin)
            self.log(f"Fichier sélectionné : {chemin}", 'info')
            nom_asm = os.path.splitext(os.path.basename(chemin))[0]
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            self.nom_sortie_edit.setText(f"Export_PLM_{nom_asm}_{timestamp}.xlsx")
    
    def choisir_dossier_dft(self):
        dossier = QFileDialog.getExistingDirectory(
            self,
            "Sélectionnez le dossier contenant les plans (.dft)",
            ""
        )
        if dossier:
            self.dossier_dft_edit.setText(dossier)
            self.log(f"Dossier plans sélectionné : {dossier}", 'info')
    
    def log(self, message, msg_type='info'):
        cursor = self.console.textCursor()
        cursor.movePosition(QTextCursor.End)
        color = QColor('black')
        if msg_type == 'success': color = QColor('green')
        elif msg_type == 'error': color = QColor('red')
        elif msg_type == 'warning': color = QColor('orange')
        self.console.setTextColor(color)
        cursor.insertText(message + '\n')
        self.console.setTextCursor(cursor)
        self.console.ensureCursorVisible()
    
    def lancer_extraction(self):
        if self.extraction_en_cours:
            QMessageBox.warning(self, "Attention", "Une extraction est déjà en cours.")
            return
        
        chemin_asm = self.chemin_asm_edit.text()
        if not chemin_asm:
            QMessageBox.warning(self, "Attention", "Veuillez sélectionner un fichier ASM.")
            return
        
        if not os.path.exists(chemin_asm):
            QMessageBox.critical(self, "Erreur", "Le fichier ASM sélectionné n'existe pas.")
            return
        
        os.makedirs(self.dossier_sortie, exist_ok=True)
        self.extraction_en_cours = True
        self.btn_extraire.setEnabled(False)
        self.btn_extraire.setText("Extraction en cours...")
        
        dossier_dft = self.dossier_dft_edit.text() if self.dossier_dft_edit.text() else None
        self.thread = ExtractionThread(chemin_asm, self.dossier_sortie, self.nom_sortie_edit.text(), dossier_dft)
        self.thread.log_signal.connect(self.log)
        self.thread.finished_signal.connect(self.extraction_terminee)
        self.thread.start()
    
    def extraction_terminee(self):
        self.extraction_en_cours = False
        self.btn_extraire.setEnabled(True)
        self.btn_extraire.setText("Lancer l'extraction")

def main():
    app = QApplication(sys.argv)
    window = PLMExtractorGUI()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()