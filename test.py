import win32com.client
import os

chemin = r"C:\Users\ykorichi\Desktop\DEV\test\CODEUR SUR ARBRE\3d\ASM - Moto reducteur  codeur.asm"

print(os.path.exists(chemin))

prop_reader = win32com.client.Dispatch("SolidEdge.FileProperties")
prop_reader.Open(chemin)

for set_name in ["SummaryInformation", "ExtendedSummaryInformation", "Custom", "DocSummaryInformation"]:
    print(f"\n=== {set_name} ===")
    try:
        ps = prop_reader.Item(set_name)
        print(f"  Count: {ps.Count}")
        for i in range(1, ps.Count + 1):
            try:
                p = ps.Item(i)
                print(f"  [{i}] Name='{p.Name}' | Value='{p.Value}' | Type={type(p.Value)}")
            except Exception as e:
                print(f"  [{i}] ERREUR: {e}")
    except Exception as e:
        print(f"  Non disponible: {e}")

prop_reader.Close()