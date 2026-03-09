Set objShell = CreateObject("WScript.Shell")
objShell.Run "powershell.exe -ExecutionPolicy Bypass -File ""C:\TSA\Scripts\Redirect-FoldersOneDrive.ps1""", 0, False
