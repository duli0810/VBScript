Dim objShell

Set objShell = CreateObject("WScript.Shell")

strCMD = "powershell -sta -noProfile -NonInteractive -nologo -command " & Chr(34) & "c:\AD\nome-do-seu-script-Powershell.ps1" & Chr(34)

objShell.Run strCMD,0