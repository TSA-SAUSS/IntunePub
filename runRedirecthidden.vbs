Set objShell = CreateObject("WScript.Shell")
objShell.Run "powershell.exe -ExecutionPolicy Bypass -File ""C:\TSA\InTuneAppLog\Redirect-FoldersOneDrive\Redirect-FoldersOneDrive.ps1""", 0, False
