import sys
import os
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog, QMessageBox, QComboBox, QDialog, QProgressBar)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QTextCursor, QColor

from se_to_plm import generer_export_excel

class ExtractionThread(QThread):
    log_signal = pyqtSignal(str, str)
    finished_signal = pyqtSignal()
    progress_signal = pyqtSignal(int, int, str)  # (valeur, maximum, message)

    def __init__(self, chemin_asm, dossier_sortie, nom_sortie, dossier_dft=None, mode_recherche_dft="les_deux"):
        super().__init__()
        self.chemin_asm = chemin_asm
        self.dossier_sortie = dossier_sortie
        self.nom_sortie = nom_sortie
        self.dossier_dft = dossier_dft
        self.mode_recherche_dft = mode_recherche_dft
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
                chemin_asm=self.chemin_asm,
                dossier_sortie=self.dossier_sortie,
                nom_sortie=self.nom_sortie,
                dossier_dft=self.dossier_dft,
                mode_recherche=self.mode_recherche_dft,
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
        """Fermer Solid Edge quand l'application est fermée avec une fenêtre de chargement."""
        # Créer une fenêtre de chargement
        loading_dialog = QDialog(self)
        loading_dialog.setWindowTitle("Fermeture")
        loading_dialog.setFixedSize(300, 100)
        layout = QVBoxLayout(loading_dialog)
        label = QLabel("Fermeture de Solid Edge...")
        label.setAlignment(Qt.AlignCenter)
        layout.addWidget(label)
        loading_dialog.show()
        
        # Forcer la mise à jour de l'interface
        QApplication.processEvents()
        
        try:
            app = win32com.client.dynamic.Dispatch("SolidEdge.Application")
            app.Quit()
        except:
            pass
        
        loading_dialog.close()
        event.accept()
    
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
        
        # Convertir l'index du combo en mode de recherche
        mode_index = self.mode_dft_combo.currentIndex()
        if mode_index == 0:
            mode_recherche = "arborescence"
        elif mode_index == 1:
            mode_recherche = "dossier_specifique"
        else:
            mode_recherche = "les_deux"
        
        self.thread = ExtractionThread(chemin_asm, self.dossier_sortie, self.nom_sortie_edit.text(), dossier_dft, mode_recherche)
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