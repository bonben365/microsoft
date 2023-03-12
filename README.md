# Installation

- Right click on the Windows Start icon 🪟 then select Windows PowerShell (Admin).
- Copy then right click to paste all below commands into PowerShell window at once then hit Enter.
- Allow system to running a script: https://go.microsoft.com/fwlink/?LinkID=135170

```ps
Set-ExecutionPolicy Bypass -Scope Process -Force
$url="https://raw.githubusercontent.com/bonben365/microsoft/main/msx.ps1"
iex ((New-Object System.Net.WebClient).DownloadString($url))
```
