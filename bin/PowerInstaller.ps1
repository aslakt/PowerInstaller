Class InstallSequence {

    [Installer[]]$Sequence
    
    InstallSequence(){
        $this.Sequence = [System.Collections.ArrayList]::New()
    }

    InstallSequence([Installer]$Sequence){
        $this.Sequence = [System.Collections.ArrayList]$Sequence
        $this.AddInstaller($Sequence)
    }
    
    InstallSequence([Installer[]]$Sequence){
        $this.Sequence = [System.Collections.ArrayList]$Sequence
    }

    [int]AddInstaller([Installer]$Sequence) {
        return ($this.Sequence.Add($Sequence))
    }

    [int]InsertInstaller([int]$index, [Installer]$Sequence) {
        return ($this.Sequence.Insert($index, $Sequence))
    }

    [int]InsertInstallerSequence([int]$index, [Installer[]]$Sequence){
        return ($this.Sequence.InsertRange($index, [System.Collections.ArrayList]$Sequence))
    }

    [int]RemoveInstaller([Installer]$Installer) {
        return ($this.Sequence.Remove($Installer))
    }

    [int]RemoveInstaller([int]$index) {
        return ($this.Sequence.RemoveAt($index))
    }

    [bool]ProcessSequence(){
        ForEach ($installer in $this.Sequence) {
            If (!($installer.Install())){
                return $false
            }
        }
        return $true
    }
}

Class Installer {

    [string]$Name
    [Version]$Version
    [DetectionMethod]$DetectionMethod
    hidden [string]$Path
    [int[]]$ValidExitCodes
    [string]$FilePath
    [string[]]$InstallArgumentList
    [string[]]$UninstallArgumentList

    [bool]Install() {
        If (!($this.DetectionMethod.Detected())) {
            $RetVal = Start-Process -FilePath $this.Path -ArgumentList $this.InstallArgumentList -NoNewWindow -Wait -PassThru
            return ($this.ValidExitCodes -contains $RetVal)
        } else {
            return $false
        }
    }

    [bool]Uninstall() {
        If ($this.DetectionMethod.Detected()) {
            $RetVal = Start-Process -FilePath $this.Path -ArgumentList $this.UninstallArgumentList -NoNewWindow -Wait -PassThru
            return ($this.ValidExitCodes -contains $RetVal)
        } else {
            return $false
        }
    }

}

Class SetupInstaller : Installer {

    SetupInstaller([string]$Path, [DetectionMethod]$DetectionMethod){
        $this.Path = $Path
        $this.FilePath = $Path
        $this.DetectionMethod = $DetectionMethod
        $this.ValidExitCodes = @(0,3010)
    }

    SetupInstaller([string]$Path, [DetectionMethod]$DetectionMethod, [string]$Name){
        $this.Path = $Path
        $this.FilePath = $Path
        $this.DetectionMethod = $DetectionMethod
        $this.Name = $Name
        $this.ValidExitCodes = @(0,3010)
    }

    SetupInstaller([string]$Path, [DetectionMethod]$DetectionMethod, [string]$Name, [Version]$Version){
        $this.Path = $Path
        $this.FilePath = $Path
        $this.DetectionMethod = $DetectionMethod
        $this.Name = $Name
        $this.Version = $Version
        $this.ValidExitCodes = @(0,3010)
    }
    
}

Class MSIInstaller : Installer {

    [guid]$ProductCode
    [guid]$UpgradeCode
    hidden [hashtable]$Properties

    MSIInstaller([string]$Path) {

        # Load some TypeData
        Update-TypeData -AppendPath $PSScriptRoot\comObject.types.ps1xml

        $this.Path = "msiexec.exe"
        $this.FilePath = $Path
        $this.ValidExitCodes = @(0,3010)

        $Installer = New-Object -ComObject WindowsInstaller.Installer
        $InstallerDataBase = $Installer.InvokeMethod("OpenDatabase", (Resolve-Path $Path).Path, 0)
        $View = $InstallerDataBase.InvokeMethod("OpenView", "SELECT `*` FROM `Property` ")
        $View.InvokeMethod("Execute")
        $Record = $View.InvokeMethod("Fetch")
        $ProductProperties = [hashtable]::new()
        While ($Record -ne $null) {
            $ProductProperties.Add($Record.InvokeParamProperty("StringData", 1), $Record.InvokeParamProperty("StringData", 2))
            $Record = $View.InvokeMethod("Fetch")
        }
        $this.Properties = $ProductProperties
        $this.ProductCode = [guid]$ProductProperties.Item("ProductCode")
        $this.UpgradeCode = [guid]$ProductProperties.Item("UpgradeCode")
        $this.Name = $ProductProperties.Item("ProductName")
        $this.Version = $ProductProperties.Item("ProductVersion")

        $this.DetectionMethod = [DetectionMethod]::new("MSI", $this.ProductCode)
        $this.InstallArgumentList = @("-i", (Get-Item $Path).Name, "-qb-", "ALLUSERS=1", "REBOOT=r")
        $this.UninstallArgumentList = @("-x", $ProductProperties.Item("UpgradeCode"), "-qb-", "ALLUSERS=1", "REBOOT=r")
    }

    [string]GetProductCode() {
        return ("{" + $this.ProductCode + "}")
    }

    [string]GetUpgradeCode() {
        return ("{" + $this.UpgradeCode + "}")
    }

}

Class DetectionMethod {
    [ValidateSet("File", "MSI", "Registry")][string]$DetectionType
    [string]$Path
    hidden [guid]$ProductCode

    DetectionMethod([string]$DetectionType, [string]$Path) {
        $this.DetectionType = $DetectionType
        $this.Path = $Path
        If ($DetectionType -eq "MSI") {
            $Installer = New-Object -ComObject WindowsInstaller.Installer
            $InstallerDataBase = $Installer.InvokeMethod("OpenDatabase", (Resolve-Path $Path).Path, 0)
            $View = $InstallerDataBase.InvokeMethod("OpenView", "SELECT `*` FROM `Property` ")
            $View.InvokeMethod("Execute")
            $Record = $View.InvokeMethod("Fetch")
            $ProductProperties = [hashtable]::new()
            While ($Record -ne $null) {
                $ProductProperties.Add($Record.InvokeParamProperty("StringData", 1), $Record.InvokeParamProperty("StringData", 2))
                $Record = $View.InvokeMethod("Fetch")
            }
            $this.ProductCode = [guid]$ProductProperties.Item("ProductCode")
        }
    }

    [bool]Detected() {
        if ($this.DetectionType -eq "File" -or $this.DetectionType -eq "Registry") {
            return (Test-Path $this.Path)
        } elseif ($this.DetectionType -eq "MSI") {
            $Installer = New-Object -ComObject WindowsInstaller.Installer
            return ($Installer.Productstate(("{" + $this.ProductCode + "}")) -eq 5)
        } else {
            return $false
        }        
    }
}