import sys
import os
from datetime import datetime
import win32com.client
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog, QMessageBox, QComboBox, QDialog, QProgressBar)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QTextCursor, QColor

from se_to_plm import generer_export_excel

class ExtractionThread(QThread):
    log_signal = pyqtSignal(str, str)
    finished_signal = pyqtSignal()
    progress_signal = pyqtSignal(int, int, str)  # (valeur, maximum, message)

    def __init__(self, chemin_fichier, dossier_sortie, nom_sortie, dossier_dft=None, mode_recherche_dft="les_deux", type_fichier="asm"):
        super().__init__()
        self.chemin_fichier = chemin_fichier
        self.dossier_sortie = dossier_sortie
        self.nom_sortie = nom_sortie
        self.dossier_dft = dossier_dft
        self.mode_recherche_dft = mode_recherche_dft
        self.type_fichier = type_fichier
        self._cancelled = False
    
    def cancel(self):
        self._cancelled = True

    def run(self):
        try:
            # Définir les callbacks pour le moteur
            def callback_log(message, msg_type):
                self.log_signal.emit(message, msg_type)
            
            def callback_progress(value, maximum, message):
                self.progress_signal.emit(value, maximum, message)
            
            def check_cancelled():
                return self._cancelled
            
            # Utiliser le moteur centralisé
            resultat = generer_export_excel(
                chemin_fichier=self.chemin_fichier,
                dossier_sortie=self.dossier_sortie,
                nom_sortie=self.nom_sortie,
                dossier_dft=self.dossier_dft,
                mode_recherche=self.mode_recherche_dft,
                type_fichier=self.type_fichier,
                callback_log=callback_log,
                callback_progress=callback_progress,
                check_cancelled=check_cancelled
            )
            
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
    
    def closeEvent(self, event):
        """Demander confirmation pour fermer Solid Edge quand l'application est fermée."""
        # Vérifier si Solid Edge est en cours d'exécution
        try:
            app_se = win32com.client.GetActiveObject("SolidEdge.Application")
        except:
            # Solid Edge n'est pas lancé, fermer l'application normalement
            event.accept()
            return
        
        # Demander confirmation à l'utilisateur
        reply = QMessageBox.question(
            self, 
            "Fermeture de Solid Edge",
            "Voulez-vous fermer Solid Edge pour libérer la licence ?\n\n"
            "Cela libérera la licence pour d'autres utilisateurs.",
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
            QMessageBox.Yes
        )
        
        if reply == QMessageBox.Cancel:
            event.ignore()
            return
        elif reply == QMessageBox.No:
            event.accept()
            return
        
        # Fermer Solid Edge
        self._fermer_solid_edge(app_se)
        event.accept()
    
    def _fermer_solid_edge(self, app_se):
        """Ferme Solid Edge avec plusieurs méthodes si nécessaire."""
        import time
        import os
        
        # Afficher la fenêtre de chargement
        loading_dialog = self._creer_fenetre_chargement()
        
        try:
            # Méthode 1: Quit() normal
            if self._essayer_quitter(app_se, 3):
                self.log("Solid Edge a été fermé avec succès.", 'success')
                return
            
            # Méthode 2: Fermer les documents puis quitter
            if self._fermer_documents_puis_quitter(app_se):
                self.log("Solid Edge a été fermé avec succès.", 'success')
                return
            
            # Méthode 3: taskkill (dernier recours)
            os.system("taskkill /f /im SolidEdge.exe")
            time.sleep(1)
            self.log("Solid Edge a été fermé de force.", 'warning')
            
        except Exception as e:
            self.log(f"Erreur lors de la fermeture de Solid Edge: {e}", 'error')
            # Dernière tentative avec taskkill
            try:
                os.system("taskkill /f /im SolidEdge.exe")
                self.log("Solid Edge a été fermé de force.", 'warning')
            except:
                self.log("Impossible de fermer Solid Edge automatiquement.", 'error')
        
        finally:
            loading_dialog.close()
    
    def _creer_fenetre_chargement(self):
        """Crée et affiche une fenêtre de chargement."""
        loading_dialog = QDialog(self)
        loading_dialog.setWindowTitle("Fermeture")
        loading_dialog.setFixedSize(300, 100)
        layout = QVBoxLayout(loading_dialog)
        label = QLabel("Fermeture de Solid Edge...")
        label.setAlignment(Qt.AlignCenter)
        layout.addWidget(label)
        loading_dialog.show()
        QApplication.processEvents()
        return loading_dialog
    
    def _essayer_quitter(self, app_se, delai_attente=3):
        """Essaye de fermer Solid Edge avec Quit() et vérifie le résultat."""
        import time
        try:
            app_se.Quit()
            time.sleep(delai_attente)
            # Vérifier si Solid Edge est encore ouvert
            win32com.client.GetActiveObject("SolidEdge.Application")
            return False  # Encore ouvert
        except:
            return True  # Fermé avec succès
    
    def _fermer_documents_puis_quitter(self, app_se):
        """Ferme tous les documents puis essaie de quitter."""
        import time
        try:
            # Fermer tous les documents
            for doc in app_se.Documents:
                try:
                    doc.Close()
                except:
                    pass
            
            time.sleep(1)
            app_se.Quit()
            time.sleep(2)
            
            # Vérifier si c'est fermé
            win32com.client.GetActiveObject("SolidEdge.Application")
            return False
        except:
            return True
    
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
        
        # Type de fichier principal
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("Type de fichier :"))
        self.type_fichier_combo = QComboBox()
        self.type_fichier_combo.addItems([
            "Assemblage (.asm)",
            "Pièces (.par/.psm)",
            "Les deux"
        ])
        self.type_fichier_combo.setCurrentIndex(0)  # "Assemblage" par défaut
        self.type_fichier_combo.currentIndexChanged.connect(self.on_type_fichier_changed)
        type_layout.addWidget(self.type_fichier_combo)
        main_layout.addLayout(type_layout)
        
        # Fichier principal
        fichier_layout = QHBoxLayout()
        self.label_fichier = QLabel("Fichier ASM :")
        fichier_layout.addWidget(self.label_fichier)
        self.chemin_fichier_edit = QLineEdit()
        self.chemin_fichier_edit.setReadOnly(True)
        fichier_layout.addWidget(self.chemin_fichier_edit)
        btn_parcourir = QPushButton("Parcourir...")
        btn_parcourir.clicked.connect(self.choisir_fichier)
        fichier_layout.addWidget(btn_parcourir)
        main_layout.addLayout(fichier_layout)
        
        # Mode de recherche DFT
        mode_layout = QHBoxLayout()
        mode_layout.addWidget(QLabel("Mode recherche DFT :"))
        self.mode_dft_combo = QComboBox()
        self.mode_dft_combo.addItems([
            "Arborescence uniquement",
            "Dossier spécifique uniquement",
            "Les deux (arborescence + dossier)"
        ])
        self.mode_dft_combo.setCurrentIndex(2)  # "Les deux" par défaut
        self.mode_dft_combo.currentIndexChanged.connect(self.on_mode_dft_changed)
        mode_layout.addWidget(self.mode_dft_combo)
        main_layout.addLayout(mode_layout)

        # Dossier DFT spécifique
        dft_layout = QHBoxLayout()
        dft_layout.addWidget(QLabel("Dossier plans (.dft) :"))
        self.dossier_dft_edit = QLineEdit()
        self.dossier_dft_edit.setReadOnly(True)
        self.dossier_dft_edit.setPlaceholderText("Sélectionner un dossier...")
        self.dossier_dft_edit.setEnabled(False)  # Désactivé par défaut
        dft_layout.addWidget(self.dossier_dft_edit)
        self.btn_parcourir_dft = QPushButton("Parcourir...")
        self.btn_parcourir_dft.clicked.connect(self.choisir_dossier_dft)
        self.btn_parcourir_dft.setEnabled(False)  # Désactivé par défaut
        dft_layout.addWidget(self.btn_parcourir_dft)
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
        
        # Barre de progression
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p% - %v")
        main_layout.addWidget(self.progress_bar)
        
        self.lbl_progress = QLabel("Prêt")
        self.lbl_progress.setAlignment(Qt.AlignCenter)
        self.lbl_progress.setStyleSheet("color: gray; font-size: 11px;")
        main_layout.addWidget(self.lbl_progress)
        
        main_layout.addWidget(QLabel("Console de progression :"))
        self.console = QTextEdit()
        self.console.setReadOnly(True)
        main_layout.addWidget(self.console)
    
    def on_type_fichier_changed(self, index):
        """Gère le changement de type de fichier principal."""
        if index == 0:  # Assemblage
            self.label_fichier.setText("Fichier ASM :")
        elif index == 1:  # Pièces
            self.label_fichier.setText("Fichier pièce :")
        else:  # Les deux
            self.label_fichier.setText("Fichier principal :")
        
        # Vider le champ de fichier
        self.chemin_fichier_edit.clear()
    
    def choisir_fichier(self):
        """Ouvre le dialogue de sélection selon le type de fichier choisi."""
        type_index = self.type_fichier_combo.currentIndex()
        
        if type_index == 0:  # Assemblage
            titre = "Sélectionnez l'assemblage principal (.asm)"
            filtre = "Assemblage Solid Edge (*.asm)"
        elif type_index == 1:  # Pièces
            titre = "Sélectionnez une pièce (.par/.psm)"
            filtre = "Pièces Solid Edge (*.par *.psm)"
        else:  # Les deux
            titre = "Sélectionnez un fichier Solid Edge"
            filtre = "Fichiers Solid Edge (*.asm *.par *.psm)"
        
        chemin, _ = QFileDialog.getOpenFileName(
            self,
            titre,
            "",
            filtre
        )
        if chemin:
            self.chemin_fichier_edit.setText(chemin)
            self.log(f"Fichier sélectionné : {chemin}", 'info')
            nom_fichier = os.path.splitext(os.path.basename(chemin))[0]
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            self.nom_sortie_edit.setText(f"Export_PLM_{nom_fichier}_{timestamp}.xlsx")
    
    def on_mode_dft_changed(self, index):
        """Active/désactive le champ dossier DFT selon le mode sélectionné."""
        # Mode 0 = Arborescence uniquement (désactivé)
        # Mode 1 = Dossier spécifique uniquement (activé)
        # Mode 2 = Les deux (activé)
        if index == 1:  # Dossier spécifique uniquement
            self.dossier_dft_edit.setEnabled(True)
            self.btn_parcourir_dft.setEnabled(True)
        elif index == 2:  # Les deux
            self.dossier_dft_edit.setEnabled(True)
            self.btn_parcourir_dft.setEnabled(True)
        else:  # Arborescence uniquement
            self.dossier_dft_edit.setEnabled(False)
            self.btn_parcourir_dft.setEnabled(False)
            self.dossier_dft_edit.clear()

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
        
        chemin_fichier = self.chemin_fichier_edit.text()
        type_index = self.type_fichier_combo.currentIndex()
        
        if not chemin_fichier:
            type_nom = ["ASM", "pièce", "fichier"][type_index]
            QMessageBox.warning(self, "Attention", f"Veuillez sélectionner un {type_nom}.")
            return
        
        if not os.path.exists(chemin_fichier):
            QMessageBox.critical(self, "Erreur", "Le fichier sélectionné n'existe pas.")
            return
        
        os.makedirs(self.dossier_sortie, exist_ok=True)
        self.extraction_en_cours = True
        self.btn_extraire.setEnabled(False)
        self.btn_extraire.setText("Extraction en cours...")
        
        dossier_dft = self.dossier_dft_edit.text() if self.dossier_dft_edit.text() else None
        
        # Convertir l'index du combo en mode de recherche
        mode_index = self.mode_dft_combo.currentIndex()
        if mode_index == 0:
            mode_recherche = "arborescence"
        elif mode_index == 1:
            mode_recherche = "dossier_specifique"
        else:
            mode_recherche = "les_deux"
        
        # Convertir l'index du type de fichier
        if type_index == 0:
            type_fichier = "asm"
        elif type_index == 1:
            type_fichier = "pieces"
        else:
            type_fichier = "les_deux"
        
        self.thread = ExtractionThread(chemin_fichier, self.dossier_sortie, self.nom_sortie_edit.text(), dossier_dft, mode_recherche, type_fichier)
        self.thread.log_signal.connect(self.log)
        self.thread.progress_signal.connect(self.update_progress)
        self.thread.finished_signal.connect(self.extraction_terminee)
        self.thread.start()
    
    def update_progress(self, value, maximum, message):
        self.progress_bar.setValue(value)
        self.progress_bar.setMaximum(maximum)
        self.lbl_progress.setText(message)
    
    def extraction_terminee(self):
        self.extraction_en_cours = False
        self.btn_extraire.setEnabled(True)
        self.btn_extraire.setText("Lancer l'extraction")
        self.progress_bar.setValue(0)
        self.lbl_progress.setText("Prêt")

def main():
    app = QApplication(sys.argv)
    window = PLMExtractorGUI()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()