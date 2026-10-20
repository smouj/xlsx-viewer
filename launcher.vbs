Option Explicit
Dim fso, sh, base, url, candidates, i
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
base = fso.GetParentFolderName(WScript.ScriptFullName)
url = "file:///" & Replace(base & "\index.html", "\", "/")
candidates = Array( _
  "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", _
  "C:\Program Files\Google\Chrome\Application\chrome.exe", _
  "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", _
  "C:\Program Files\Microsoft\Edge\Application\msedge.exe")
For i = 0 To UBound(candidates)
  If fso.FileExists(candidates(i)) Then
    sh.Run """" & candidates(i) & """ --app=""" & url & """ --window-size=1440,900", 1, False
    WScript.Quit
  End If
Next
sh.Run """" & base & "\index.html""", 1, False
