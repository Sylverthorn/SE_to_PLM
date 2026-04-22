Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Force le script à s'exécuter dans le dossier actuel
currentFolder = fso.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = currentFolder

' Vérifie si le venv existe
venvPath = currentFolder & "\venv"
If Not fso.FolderExists(venvPath) Then
    ' Crée le venv
    WScript.Echo "Création de l'environnement virtuel..."
    shell.Run "python -m venv venv", 1, True
    
    ' Installe les requirements
    WScript.Echo "Installation des dépendances..."
    shell.Run "venv\Scripts\pip.exe install -r requirements.txt", 1, True
    WScript.Echo "Installation terminée."
End If

' Lance pythonw en arrière-plan avec le script GUI
shell.Run "venv\Scripts\pythonw.exe gui_se_to_plm.py", 0, False