Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' On s'assure que le script tourne dans le bon dossier
currentFolder = fso.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = currentFolder

' On regarde si le venv est déjà là
venvPath = currentFolder & "\venv"
If Not fso.FolderExists(venvPath) Then
    ' On affiche une notification de début
    shell.Run "powershell -Command Add-Type -AssemblyName System.Windows.Forms; $notify = New-Object System.Windows.Forms.NotifyIcon; $notify.Icon = [System.Drawing.SystemIcons]::Information; $notify.BalloonTipTitle = 'Installation'; $notify.BalloonTipText = 'Creation de l''environnement virtuel...'; $notify.Visible = $true; $notify.ShowBalloonTip(3000); Start-Sleep -Seconds 3; $notify.Dispose()", 0, False
    
    ' Si pas de venv, on le crée
    shell.Run "python -m venv venv", 0, True
    
    ' On installe les dépendances
    shell.Run "venv\Scripts\pip.exe install -r requirements.txt", 0, True
    
    ' On affiche une notification de fin
    shell.Run "powershell -Command Add-Type -AssemblyName System.Windows.Forms; $notify = New-Object System.Windows.Forms.NotifyIcon; $notify.Icon = [System.Drawing.SystemIcons]::Information; $notify.BalloonTipTitle = 'Installation'; $notify.BalloonTipText = 'Installation terminee !'; $notify.Visible = $true; $notify.ShowBalloonTip(3000); Start-Sleep -Seconds 3; $notify.Dispose()", 0, False
End If

' On lance l'interface graphique
shell.Run "venv\Scripts\pythonw.exe gui_se_to_plm.py", 0, False