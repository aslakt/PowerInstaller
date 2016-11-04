# PowerInstallerLoader.ps1
Get-Module PowerInstaller* | Remove-Module -Verbose
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Update-TypeData -AppendPath "C:\Users\aslak\OneDrive\My Code\GitHub\PowerInstaller\bin\comObject.types.ps1xml"

Import-Module "C:\Users\aslak\OneDrive\My Code\GitHub\PowerInstaller\bin\\PowerInstaller.ps1"