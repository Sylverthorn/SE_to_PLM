import win32com.client
import os

# Chemin de votre fichier
chemin = r"C:\Users\ykorichi\Desktop\DEV\test\CODEUR SUR ARBRE\2d\Connecteur arbre-codeur M6.dft"

print(f"Le fichier existe-t-il ? {os.path.exists(chemin)}")

if os.path.exists(chemin):
    # Initialisation de l'objet COM
    prop_reader = win32com.client.Dispatch("SolidEdge.FileProperties")
    
    try:
        # Ouverture du fichier
        prop_reader.Open(chemin)
        
        # 1. On cible l'onglet "Custom"
        custom_props = prop_reader.Item("Custom")
        
        # 2. On cible directement le nom de la propriété voulue
        nom_recherche = "auteur modif"
        
        print(f"\nTentative de récupération de la propriété : '{nom_recherche}'...")
        
        try:
            # Récupération directe par le nom
            ma_propriete = custom_props.Item(nom_recherche)
            
            print("=" * 50)
            print(f"✅ SUCCÈS ! ")
            print(f"Nom   : {ma_propriete.Name}")
            print(f"Valeur: '{ma_propriete.Value}'")
            print(f"Type  : {type(ma_propriete.Value)}")
            print("=" * 50)
            
        except Exception as e:
            print(f"\n❌ ERREUR : Impossible de trouver ou lire '{nom_recherche}'.")
            print(f"Détails techniques : {e}")
            print("Vérifiez l'orthographe exacte dans Solid Edge (majuscules/espaces).")

    except Exception as e:
        print(f"Erreur lors de l'ouverture des propriétés du fichier : {e}")
        
    finally:
        # On s'assure de TOUJOURS fermer le fichier pour ne pas le bloquer dans Windows
        try:
            prop_reader.Close()
            print("\n[Lecteur de propriétés fermé avec succès]")
        except:
            pass