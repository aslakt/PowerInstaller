Class DetectionMethod {

    [ValidateSet("File", "MSI", "Registry")][string]$DetectionType
    [string]$Path

    DetectionMethod([string]$DetectionType, [string]$Path) {
        $this.DetectionType = $DetectionType
        $this.Path = $Path
    }

    [bool]Detected() {
        if ($this.DetectionType -eq "File" -or $this.DetectionType -eq "Registry") {
            return (Test-Path $this.Path)
        } elseif ($this.DetectionType -eq "MSI") {
            $Installer = New-Object -ComObject WindowsInstaller.Installer
            return ($Installer.Productstate($this.Path) -eq 5)
        } else {
            return $false
        }        
    }

}

Class Installer {

    [string]$Name
    [Version]$Version
    [DetectionMethod]$DetectionMethod
    [int[]]$ValidExitCodes
    [string]$FilePath
    [string[]]$InstallArgumentList
    [string[]]$UninstallArgumentList
    hidden [string]$Path

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

    SetupInstaller([string]$Path, [DetectionMethod]$DetectionMethod) {
        $this.Path = $Path
        $this.FilePath = $Path
        $this.DetectionMethod = $DetectionMethod
        $this.ValidExitCodes = @(0,3010)
    }

    SetupInstaller([string]$Path, [DetectionMethod]$DetectionMethod, [string]$Name) {
        $this.Path = $Path
        $this.FilePath = $Path
        $this.DetectionMethod = $DetectionMethod
        $this.Name = $Name
        $this.ValidExitCodes = @(0,3010)
    }

    SetupInstaller([string]$Path, [DetectionMethod]$DetectionMethod, [string]$Name, [Version]$Version) {
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

Class SequenceItem {

    [Installer[]]$Installers
    [string]$Name

    SequenceItem([Installer]$Installer) {
        $this.Name = $Installer.Name
        $NewInstallers = New-Object System.Collections.ArrayList
        $NewInstallers.Add($Installer)
        $this.Installers = $NewInstallers
    }

    SequenceItem([string]$Name, [Installer]$Installer) {
        $this.Name = $Name
        $NewInstallers = New-Object System.Collections.ArrayList
        $NewInstallers.Add($Installer)
        $this.Installers = $NewInstallers
    }

    SequenceItem([string]$Name, [Installer[]]$Installers) {
        $this.Name = $Name
        $this.Installers = $Installers
    }

    [int]AddInstaller([Installer]$Sequence) {
        return ($this.Sequence.Add($Sequence))
    }

    [int]InsertInstaller([int]$index, [Installer]$Sequence) {
        return ($this.Sequence.Insert($index, $Sequence))
    }
    
    [int]RemoveInstaller([Installer]$Installer) {
        return ($this.Sequence.Remove($Installer))
    }

    [int]RemoveInstaller([int]$index) {
        return ($this.Sequence.RemoveAt($index))
    }

    [System.Windows.Controls.TreeViewItem]GetTreeViewItem([System.Windows.Controls.TreeViewItem]$Parent){
        $TreeViewItem = [System.Windows.Controls.TreeViewItem]::new()
        $TreeViewItem.Header = $this.Name
        $TreeViewItem.Name = ("seq_" + ($this.Name -replace '[\W_]', ""))
        $TreeViewItem.Tag = ($Parent.Tag + "\" + $this.Name)
        ForEach ($Installer in $this.Installers) {
            $ChildItem = [System.Windows.Controls.TreeViewItem]::new()
            $ChildItem.Header = $Installer.Name
            $ChildItem.Name = ("inst_" + ($Installer.Name -replace '[\W_]', ""))
            $ChildItem.Tag = ($TreeViewItem.Tag + "\" + $Installer.Name)
            $TreeViewItem.Items.Add($ChildItem)
        }
        return $TreeViewItem
    }
}

Class InstallSequence {

    [SequenceItem[]]$Sequence
    [string]$Name
    
    InstallSequence([string]$Name) {
        $this.Name = $Name
        $this.Sequence = [System.Collections.ArrayList]::New()
    }

    InstallSequence([string]$Name, [SequenceItem]$SequenceItem) {
        $this.Name = $Name
        $NewSequence = [System.Collections.ArrayList]::New()
        $NewSequence.Add($SequenceItem)
        $this.Sequence = $NewSequence
    }
    
    InstallSequence([string]$Name, [SequenceItem[]]$Sequence) {
        $this.Name = $Name
        $this.Sequence = $Sequence
    }

    [int]InsertInstallerSequence([int]$index, [SequenceItem[]]$Sequence){
        return ($this.Sequence.InsertRange($index, [System.Collections.ArrayList]$Sequence))
    }

    [bool]ProcessSequence(){
        ForEach ($installer in $this.Sequence) {
            If (!($installer.Install())){
                return $false
            }
        }
        return $true
    }

    [System.Windows.Controls.TreeViewItem]GetTreeViewItem(){
        $TreeViewItem = [System.Windows.Controls.TreeViewItem]::new()
        $TreeViewItem.Header = $this.Name
        $TreeViewItem.Name = ("instSeq_" + ($this.Name -replace '[\W_]', ""))
        $TreeViewItem.Tag = $this.Name
        ForEach ($Sequence in $this.Sequence) {
            [System.Windows.Controls.TreeViewItem]$ChildItem = $Sequence.GetTreeViewItem($TreeViewItem)
            $ChildItem.Tag = ($TreeViewItem.Tag + "\" + $ChildItem.Tag)
            $TreeViewItem.Items.Add($ChildItem)
        }
        return $TreeViewItem
    }

}