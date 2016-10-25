# PowerInstallerLoader.ps1
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Update-TypeData -AppendPath $PSScriptRoot\comObject.types.ps1xml

Import-Module $PSScriptRoot\PowerInstaller.ps1