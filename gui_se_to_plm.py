import sys
import os
import time
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog, QMessageBox)
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QTextCursor, QColor

# On importe les fonctions utiles depuis se_to_plm.py
from se_to_plm import indexer_les_plans_projet_entier, extraire_metadonnees, determiner_classe
import win32com.client
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment


class ExtractionThread(QThread):
    log_signal = pyqtSignal(str, str)  # message, type (info, success, error, warning)
    finished_signal = pyqtSignal()
    
    def __init__(self, chemin_asm, dossier_sortie, nom_sortie):
        super().__init__()
        self.chemin_asm = chemin_asm
        self.dossier_sortie = dossier_sortie
        self.nom_sortie = nom_sortie
    
    def run(self):
        try:
            self.log_signal.emit("=" * 60, 'info')
            self.log_signal.emit("Début de l'extraction PLM", 'info')
            self.log_signal.emit("=" * 60, 'info')
            
            # On commence par indexer les plans
            self.log_signal.emit("\n--- Indexation des plans (.dft) ---", 'info')
            index_plans = indexer_les_plans_projet_entier(self.chemin_asm)
            self.log_signal.emit(f"-> {len(index_plans)} plan(s) détecté(s).", 'info')
            
            # On se connecte à Solid Edge
            self.log_signal.emit("\nOuverture de Solid Edge...", 'info')
            app = win32com.client.dynamic.Dispatch("SolidEdge.Application")
            app.Visible = False
            doc_racine = app.Documents.Open(self.chemin_asm)
            time.sleep(3)
            self.log_signal.emit("Solid Edge connecté.", 'success')
            
            lignes_excel = []
            compteur_ordre = 1
            stats = {"3d": 0, "2d": 0}
            
            def ajouter_ligne(niveau, relation, nom_fichier, chemin_complet, classe, qte=1):
                nonlocal compteur_ordre
                designation = ""
                rev = ""
                ref_util = os.path.splitext(nom_fichier)[0]
                
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
                        
                        if nom_reel not in dict_occ:
                            dict_occ[nom_reel] = {"qte": 1, "obj": occ, "chemin": path_reel}
                        else:
                            dict_occ[nom_reel]["qte"] += 1
                    except: continue
                
                for nom, data in dict_occ.items():
                    stats["3d"] += 1
                    classe_3d = determiner_classe(nom)
                    
                    meta = {"designation": "", "revision": ""}
                    try:
                        meta = extraire_metadonnees(data["obj"].OccurrenceDocument)
                    except: pass
                    
                    lignes_excel.append([
                        niveau, "ComposedOf", compteur_ordre, data["qte"], "", "", nom,
                        classe_3d, os.path.splitext(nom)[0], "", meta["revision"], meta["designation"], "", data["chemin"]
                    ])
                    compteur_ordre += 1
                    
                    nom_sans_ext = os.path.splitext(nom)[0].lower()
                    if nom_sans_ext in index_plans:
                        chemin_dft = index_plans[nom_sans_ext]
                        ajouter_ligne(niveau + 1, "Drawing", os.path.basename(chemin_dft), chemin_dft, "Plan_A")
                        stats["2d"] += 1
                    
                    if data["obj"].Subassembly:
                        try: explorer_occurrences(data["obj"].OccurrenceDocument.Occurrences, niveau + 1)
                        except: pass
            
            # On lance l'analyse
            self.log_signal.emit("\nAnalyse de la structure...", 'info')
            
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
            
            self.log_signal.emit(f"Analyse terminée : {stats['3d']} fichiers 3D, {stats['2d']} plans", 'success')
            
            # On crée le fichier Excel
            self.log_signal.emit("\nGénération du fichier Excel...", 'info')
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
            
            # On sauvegarde avec le nom choisi
            nom_sortie = self.nom_sortie
            if not nom_sortie.endswith('.xlsx'):
                nom_sortie += '.xlsx'
            
            chemin_complet = os.path.join(self.dossier_sortie, nom_sortie)
            wb.save(chemin_complet)
            
            self.log_signal.emit(f"\nFichier généré : {chemin_complet}", 'success')
            self.log_signal.emit(f"3D: {stats['3d']} | Plans: {stats['2d']}", 'success')
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
        # On charge le fichier de style
        chemin_style = os.path.join(os.path.dirname(__file__), 'style.qss')
        if os.path.exists(chemin_style):
            with open(chemin_style, 'r', encoding='utf-8') as f:
                self.setStyleSheet(f.read())
        
    def creer_interface(self):
        # On crée le widget principal
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # On met en place le layout
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(10)
        
        # Section pour le fichier ASM
        asm_layout = QHBoxLayout()
        asm_layout.addWidget(QLabel("Fichier ASM :"))
        
        self.chemin_asm_edit = QLineEdit()
        self.chemin_asm_edit.setReadOnly(True)
        asm_layout.addWidget(self.chemin_asm_edit)
        
        btn_parcourir = QPushButton("Parcourir...")
        btn_parcourir.clicked.connect(self.choisir_fichier_asm)
        asm_layout.addWidget(btn_parcourir)
        
        main_layout.addLayout(asm_layout)
        
        # Section pour le nom de sortie
        sortie_layout = QHBoxLayout()
        sortie_layout.addWidget(QLabel("Nom de sortie :"))
        
        self.nom_sortie_edit = QLineEdit(f"Export_PLM_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        sortie_layout.addWidget(self.nom_sortie_edit)
        
        main_layout.addLayout(sortie_layout)
        
        # Dossier où on va sauvegarder
        self.dossier_sortie = os.path.join(os.path.expanduser("~"), "Documents", "Exports_PLM")
        lbl_dossier = QLabel(f"Dossier : {self.dossier_sortie}")
        lbl_dossier.setStyleSheet("color: gray; font-style: italic;")
        main_layout.addWidget(lbl_dossier)
        
        # Bouton pour lancer
        self.btn_extraire = QPushButton("Lancer l'extraction")
        self.btn_extraire.clicked.connect(self.lancer_extraction)
        self.btn_extraire.setMinimumHeight(40)
        main_layout.addWidget(self.btn_extraire)
        
        # Zone pour voir ce qui se passe
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
            
            # On met à jour le nom de sortie avec celui du fichier ASM
            nom_asm = os.path.splitext(os.path.basename(chemin))[0]
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            self.nom_sortie_edit.setText(f"Export_PLM_{nom_asm}_{timestamp}.xlsx")
    
    def log(self, message, msg_type='info'):
        """Ajoute un message dans la console."""
        cursor = self.console.textCursor()
        cursor.movePosition(QTextCursor.End)
        
        color = QColor('black')
        if msg_type == 'success':
            color = QColor('green')
        elif msg_type == 'error':
            color = QColor('red')
        elif msg_type == 'warning':
            color = QColor('orange')
        
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
        
        # On crée le dossier de sortie si besoin
        os.makedirs(self.dossier_sortie, exist_ok=True)
        
        # On désactive le bouton pendant le traitement
        self.extraction_en_cours = True
        self.btn_extraire.setEnabled(False)
        self.btn_extraire.setText("Extraction en cours...")
        
        # On lance l'extraction dans un thread à part
        self.thread = ExtractionThread(chemin_asm, self.dossier_sortie, self.nom_sortie_edit.text())
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
