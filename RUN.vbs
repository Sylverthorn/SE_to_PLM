Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' On s'assure que le script tourne dans le bon dossier
currentFolder = fso.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = currentFolder

' On regarde si le venv est déjà là
venvPath = currentFolder & "\venv"
If Not fso.FolderExists(venvPath) Then
    ' Si pas de venv, on le crée
    WScript.Echo "Création de l'environnement virtuel..."
    shell.Run "python -m venv venv", 1, True
    
    ' On installe les dépendances
    WScript.Echo "Installation des dépendances..."
    shell.Run "venv\Scripts\pip.exe install -r requirements.txt", 1, True
    WScript.Echo "Installation terminée."
End If

' On lance l'interface graphique
shell.Run "venv\Scripts\pythonw.exe gui_se_to_plm.py", 0, False