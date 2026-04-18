"""
Script pour extraire les liens relationnels d'un projet Solid Edge
et les exporter vers un fichier Excel.
"""

import os
import win32com.client
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment


class SolidEdgeRelationalExtractor:
    """Classe pour extraire les liens relationnels depuis Solid Edge"""
    
    def __init__(self):
        self.application = None
        self.relational_data = []
        self.application_started = False  # Pour savoir si on a démarré l'application
        
    def close_all_documents(self):
        """Fermer tous les documents ouverts dans Solid Edge"""
        if not self.application:
            return
        
        try:
            documents = self.application.Documents
            while documents.Count > 0:
                try:
                    doc = documents.Item(1)
                    doc.Close(False)  # Fermer sans sauvegarder
                except:
                    pass
            print("Tous les documents fermés")
        except Exception as e:
            print(f"Erreur lors de la fermeture des documents: {e}")
    
    def connect_to_solid_edge(self):
        """Se connecter à une instance existante de Solid Edge ou en démarrer une nouvelle"""
        try:
            # Essayer de se connecter à une instance existante
            self.application = win32com.client.GetActiveObject("SolidEdge.Application")
            print("Connecté à l'instance existante de Solid Edge")
            self.application_started = False
        except:
            try:
                # Démarrer une nouvelle instance
                self.application = win32com.client.Dispatch("SolidEdge.Application")
                self.application.Visible = True
                print("Nouvelle instance de Solid Edge démarrée")
                self.application_started = True
            except Exception as e:
                raise Exception(f"Impossible de se connecter à Solid Edge: {e}")
    
    def extract_from_assembly(self, assembly_path):
        """Extraire les liens relationnels d'un fichier d'assemblage"""
        if not self.application:
            self.connect_to_solid_edge()
        
        try:
            # Fermer tous les documents ouverts avant de commencer
            self.close_all_documents()
            
            # Ouvrir le fichier d'assemblage
            documents = self.application.Documents
            assembly = documents.Open(assembly_path)
            
            print(f"Ouverture de l'assemblage: {assembly_path}")
            
            # Parcourir la structure de l'assemblage
            self._traverse_assembly(assembly, level=0, parent_info=None)
            
            # Fermer le document sans sauvegarder
            assembly.Close(False)
            
        except Exception as e:
            print(f"Erreur lors de l'extraction de {assembly_path}: {e}")
    
    def _traverse_assembly(self, assembly, level=0, parent_info=None, order=0):
        """Parcourir récursivement la structure de l'assemblage pour extraire la hiérarchie PDM"""
        try:
            # L'assemblage racine est ajouté ici. 
            # Les sous-assemblages, eux, sont ajoutés par la boucle de leur parent.
            if level == 0:
                self._add_component_entry(assembly, level, "ComposedOf", is_assembly=True)
            
            # Chercher le fichier de dessin associé
            drawing_path = self._find_associated_drawing(assembly)
            if drawing_path:
                self._add_drawing_entry(drawing_path, level, assembly)
            
            # Accéder aux occurrences de l'assemblage
            occurrences = assembly.Occurrences
            
            for i in range(1, occurrences.Count + 1):
                occurrence = occurrences.Item(i)
                order += 1
                
                # Ajouter le composant avec la relation ComposedOf
                self._add_component_entry(occurrence, level + 1, "ComposedOf")
                
                # Si c'est un sous-assemblage, parcourir récursivement
                try:
                    sub_assembly = occurrence.OccurrenceDocument
                    if sub_assembly.Type == 3:  # Type 3 = Assembly document
                        self._traverse_assembly(sub_assembly, level + 1, None, order)
                except:
                    # Ce n'est pas un sous-assemblage ou erreur d'accès
                    pass
                    
        except Exception as e:
            print(f"Erreur lors du parcours de l'assemblage: {e}")
    
    def _add_component_entry(self, component, level, relationship, is_assembly=False):
        """Ajouter une entrée de composant à la liste des données (Version Robuste)"""
        try:
            # 1. Récupération sécurisée du document
            if is_assembly:
                doc = component
                # Nom sans l'extension
                name = os.path.splitext(os.path.basename(doc.FullName))[0]
                doc_type = doc.Type
            else:
                # Nom de base de l'occurrence (ex: "PIECE.par:1" -> on garde "PIECE")
                name = component.Name.split(':')[0] 
                
                # Ignorer les composants qui ne sont pas inclus dans la nomenclature (BOM)
                try:
                    if not component.IncludeInBom:
                        return
                except:
                    pass # Si l'option n'est pas lisible, on continue

                # Essayer d'accéder au fichier 3D de l'occurrence
                try:
                    doc = component.OccurrenceDocument
                    doc_type = doc.Type
                except Exception:
                    # Si la pièce est inactive ou introuvable, on l'ignore proprement
                    print(f"  [!] Impossible d'accéder au document de l'occurrence: {name}")
                    return

            # 2. Détermination de la classe (Corrigé avec les bonnes constantes Solid Edge)
            if doc_type == 1:       # igPartDocument (.par)
                class_type = "PART_A"
            elif doc_type in [3, 5]: # igAssemblyDocument (.asm) ou igWeldmentDocument (.pwd)
                class_type = "SUB_ASSY_A"
            elif doc_type == 4:      # igSheetMetalDocument (.psm)
                class_type = "PART_A"
            elif doc_type == 2:      # igDraftDocument (.dft)
                class_type = "CAD_DRAWING_A"
            else:
                class_type = f"UNKNOWN"

            # 3. Création de l'entrée
            component_info = {
                'Level': level,
                'Relationship': relationship,
                'Class': class_type,
                'quantite': 1, # Dans Solid Edge, 1 occurrence = 1 pièce physique. On ne groupe pas ici.
                'ref_utilisat': name,
                'version': self._get_document_property(doc, 'Version', '1'),
                'revision': self._get_document_property(doc, 'Revision', 'A')
            }
            
            self.relational_data.append(component_info)
            
        except Exception as e:
            # Message d'erreur plus clair pour savoir quelle pièce pose problème
            comp_name = getattr(component, 'Name', 'Inconnu')
            print(f"Erreur sur le composant '{comp_name}' : {e}")
    
    def _add_drawing_entry(self, drawing_path, level, parent_doc):
        """Ajouter une entrée de dessin avec la relation Drawing"""
        try:
            drawing_name = os.path.splitext(os.path.basename(drawing_path))[0]
            
            drawing_info = {
                'Level': level,
                'Relationship': "Drawing",
                'Class': "CAD_DRAWING_A",
                'quantite': 1,
                'ref_utilisat': drawing_name,
                'version': "1",
                'revision': "A"
            }
            
            self.relational_data.append(drawing_info)
            print(f"  Ajouté: Dessin '{drawing_name}' avec relation 'Drawing' au niveau {level}")
            
        except Exception as e:
            print(f"Erreur lors de l'ajout de l'entrée dessin: {e}")
    
    def _find_associated_drawing(self, doc):
        """Chercher le fichier de dessin (.dft) associé à un document (recherche étendue)"""
        try:
            doc_path = doc.FullName
            doc_dir = os.path.dirname(doc_path)
            doc_name = os.path.splitext(os.path.basename(doc_path))[0]
            
            print(f"  Recherche de dessin pour: {doc_name}")
            
            # Liste des dossiers à chercher: dossier actuel + sous-dossiers courants
            search_dirs = [doc_dir]
            
            # Ajouter les sous-dossiers courants (DESSINS, MISES EN PLAN, PLANS, DRAWINGS)
            common_subdirs = ['DESSINS', 'MISES EN PLAN', 'PLANS', 'DRAWINGS', 'DRAFTS']
            for subdir in common_subdirs:
                subdir_path = os.path.join(doc_dir, subdir)
                if os.path.exists(subdir_path):
                    search_dirs.append(subdir_path)
            
            # Chercher dans tous les dossiers
            for search_dir in search_dirs:
                if not os.path.exists(search_dir):
                    continue
                    
                print(f"    Recherche dans: {search_dir}")
                
                for file in os.listdir(search_dir):
                    file_name, file_ext = os.path.splitext(file)
                    if file_ext.lower() == '.dft':
                        # Correspondance exacte
                        if file_name == doc_name:
                            print(f"  Dessin trouvé (nom exact): {file}")
                            return os.path.join(search_dir, file)
                        # Correspondance partielle (contient le nom)
                        elif doc_name in file_name or file_name in doc_name:
                            print(f"  Dessin trouvé (nom partiel): {file}")
                            return os.path.join(search_dir, file)
            
            print(f"  Aucun dessin trouvé pour {doc_name}")
            return None
        except Exception as e:
            print(f"  Erreur lors de la recherche de dessin: {e}")
            return None
    
    def _get_document_property(self, doc, prop_name, default_value):
        """Extraire une propriété du document (Version Robuste)"""
        try:
            # Dans Solid Edge, les propriétés sont rangées dans des "PropertySets"
            # (Résumé, Projet, Personnalisé, etc.)
            prop_sets = doc.Properties
            for i in range(1, prop_sets.Count + 1):
                prop_set = prop_sets.Item(i)
                for j in range(1, prop_set.Count + 1):
                    prop = prop_set.Item(j)
                    if prop_name.lower() in prop.Name.lower():
                        return str(prop.Value)
            return default_value
        except:
            # Si on n'arrive pas à lire les propriétés (fichier verrouillé, etc.), on renvoie la valeur par défaut
            return default_value
    
    
    def extract_from_file(self, file_path):
        """Extraire les liens d'un fichier d'assemblage spécifique"""
        self.relational_data = []
        
        if file_path.lower().endswith('.asm'):
            print(f"Traitement de: {file_path}")
            self.extract_from_assembly(file_path)
        else:
            print(f"Erreur: Le fichier doit être un fichier .asm")
        
        return self.relational_data
    
    def extract_from_folder(self, folder_path):
        """Parcourt tous les fichiers .asm d'un dossier et les traite comme racines"""
        self.relational_data = []
        
        # Récupérer tous les fichiers .asm
        asm_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.asm')]
        
        if not asm_files:
            print("Aucun fichier .asm trouvé dans ce dossier.")
            return []

        print(f"Trouvé {len(asm_files)} fichiers d'assemblage à traiter.")
        
        for asm_file in asm_files:
            full_path = os.path.join(folder_path, asm_file)
            print(f"\n>>> TRAITEMENT DE LA RACINE : {asm_file}")
            # On appelle l'extraction pour ce fichier spécifique
            self.extract_from_assembly(full_path)
            
        return self.relational_data
    
    def _aggregate_quantities(self, data):
        """Cumule les quantités pour les mêmes ref_utilisat au même niveau"""
        aggregated = {}
        
        for item in data:
            # Clé unique: Level + ref_utilisat + Class + Relationship
            key = (item['Level'], item['ref_utilisat'], item['Class'], item['Relationship'])
            
            if key in aggregated:
                # Cumuler la quantité
                aggregated[key]['quantite'] += item['quantite']
            else:
                # Nouvelle entrée
                aggregated[key] = item.copy()
        
        return list(aggregated.values())
    
    def export_to_excel(self, output_path, aggregate_quantities=False):
        """Exporter les données extraites vers un fichier Excel"""
        # Cumuler les quantités si demandé
        data_to_export = self._aggregate_quantities(self.relational_data) if aggregate_quantities else self.relational_data
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Liens Relationnels"
        
        # Définir les styles
        header_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        header_font = Font(bold=True)
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        center_alignment = Alignment(horizontal='center', vertical='center')
        
        # En-têtes
        headers = ["Level", "Relationship", "Class", "quantite", "ref_utilisat", "version", "revision"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.border = thin_border
            cell.alignment = center_alignment
        
        # Données
        for row_idx, data in enumerate(data_to_export, 2):
            ws.cell(row=row_idx, column=1, value=data['Level']).border = thin_border
            ws.cell(row=row_idx, column=2, value=data['Relationship']).border = thin_border
            ws.cell(row=row_idx, column=3, value=data['Class']).border = thin_border
            ws.cell(row=row_idx, column=4, value=data['quantite']).border = thin_border
            ws.cell(row=row_idx, column=5, value=data['ref_utilisat']).border = thin_border
            ws.cell(row=row_idx, column=6, value=data['version']).border = thin_border
            ws.cell(row=row_idx, column=7, value=data['revision']).border = thin_border
        
        # Ajuster la largeur des colonnes
        ws.column_dimensions['A'].width = 10
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 10
        ws.column_dimensions['E'].width = 30
        ws.column_dimensions['F'].width = 10
        ws.column_dimensions['G'].width = 10
        
        # Sauvegarder
        wb.save(output_path)
        print(f"Fichier Excel créé: {output_path}")


def main():
    """Fonction principale"""
    print("=== Extraction de structure PDM/PLM Solid Edge ===")
    print("Choisissez le mode d'extraction:")
    print("1. Fichier .asm unique")
    print("2. Dossier complet (tous les fichiers .asm)")
    
    choice = input("Votre choix (1 ou 2): ").strip()
    
    if choice == "1":
        path = input("Entrez le chemin du fichier d'assemblage (.asm): ").strip('"')
        if not os.path.exists(path):
            print(f"Erreur: Le fichier '{path}' n'existe pas.")
            return
        if not path.lower().endswith('.asm'):
            print(f"Erreur: Le fichier doit être un fichier .asm")
            return
        mode = "file"
    elif choice == "2":
        path = input("Entrez le chemin du dossier contenant les fichiers .asm: ").strip('"')
        if not os.path.isdir(path):
            print(f"Erreur: '{path}' n'est pas un dossier valide.")
            return
        mode = "folder"
    else:
        print("Choix invalide.")
        return
    
    # Option d'agrégation des quantités
    aggregate = input("Cumuler les quantités par référence? (o/n): ").strip().lower()
    aggregate_quantities = aggregate == 'o' or aggregate == 'y' or aggregate == 'oui' or aggregate == 'yes'
    
    # Chemin de sortie pour le fichier Excel
    if mode == "file":
        output_dir = os.path.dirname(path)
    else:
        output_dir = path
    output_path = os.path.join(output_dir, "liens_relationnels.xlsx")
    
    # Créer l'extracteur et exécuter l'extraction
    extractor = SolidEdgeRelationalExtractor()
    
    try:
        print("\nDébut de l'extraction...")
        
        if mode == "file":
            data = extractor.extract_from_file(path)
        else:
            data = extractor.extract_from_folder(path)
        
        if not data:
            print("Aucune donnée trouvée.")
            return
        
        print(f"\n{len(data)} liens relationnels extraits.")
        
        # Exporter vers Excel
        print("Export vers Excel...")
        extractor.export_to_excel(output_path, aggregate_quantities=aggregate_quantities)
        
        print(f"\nExtraction terminée avec succès!")
        print(f"Fichier créé: {output_path}")
        
    except Exception as e:
        print(f"Erreur lors de l'extraction: {e}")
    finally:
        # Nettoyage - ne quitter l'application que si on l'a démarrée
        if extractor.application and extractor.application_started:
            try:
                extractor.application.Quit()
                print("Solid Edge fermé")
            except:
                pass


if __name__ == "__main__":
    main()
