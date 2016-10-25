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
        $NewInstallers = [System.Collections.ArrayList]::new()
        $NewInstallers.Add($Installer)
        $this.Installers = $NewInstallers
    }

    SequenceItem([string]$Name, [Installer]$Installer) {
        $this.Name = $Name
        $NewInstallers = [System.Collections.ArrayList]::new()
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
        $TreeViewItem.Header = ("seq_" + ($this.Name -replace '[\W_]', ""))
        $TreeViewItem.Name = ("seq_" + ($this.Name -replace '[\W_]', ""))
        $TreeViewItem.Tag = ($Parent.Tag + "\" + $this.Name)
        ForEach ($Installer in $this.Installers) {
            $ChildItem = [System.Windows.Controls.TreeViewItem]::new()
            $ChildItem.Header = ("inst_" + ($Installer.Name -replace '[\W_]', ""))
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
        $TreeViewItem.Header = ("instSeq_" + ($this.Name -replace '[\W_]', ""))
        $TreeViewItem.Name = ("instSeq_" + ($this.Name -replace '[\W_]', ""))
        $TreeViewItem.Tag = $this.Name
        ForEach ($Sequence in $this.Sequence) {
            $ChildItem = $Sequence.GetTreeViewItem($TreeViewItem)
            $ChildItem.Tag = ($TreeViewItem.Tag + "\" + $ChildItem.Tag)
            $TreeViewItem.Items.Add($ChildItem)
        }
        return $TreeViewItem
    }

}

Class PowerInstallerApp {

    PowerInstallerApp([InstallSequence]$InstallSequence) {

        $XAML = @'
<Window x:Class="PowerInstaller.MainWindow"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:local="clr-namespace:PowerInstaller"
    mc:Ignorable="d"
    Title="Power Installer" Height="350" Width="525">
    <Grid>
        <TreeView x:Name="twInstallSequence" HorizontalAlignment="Left" Height="300" Margin="10,10,0,0" VerticalAlignment="Top" Width="450"/>
    </Grid>
</Window>
'@

        $XAML = [xml]($XAML -replace 'mc:Ignorable="d"','' -replace "x:Na",'Na' -replace '^<Win.*', '<Window')

        $reader = New-Object System.Xml.XmlNodeReader $xaml
        try {
            $Form=[Windows.Markup.XamlReader]::Load( $reader )
        } catch {
            Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
            Exit
        }
        
        <#
         # NOT ABLE TO USE THIS IN CLASSES
         # Load XAML Objects
         $xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "$($_.Name)" -Value $Form.FindName($_.Name)}
        #>

        # // Initial Bindings
        $twInstallSequence = $Form.FindName('twInstallSequence')
        $twInstallSequence.Items.Add($InstallSequence.GetTreeViewItem())

        $Form.Add_SourceInitialized( {            [System.Windows.RoutedEventHandler]$Event = {                            <#if($_.OriginalSource -is [System.Windows.Controls.TreeViewItem]){                    $TreeItem = $_.OriginalSource                    $TreeItem.items.clear()                    $TreeItem.Items.Add($InstallSequence.GetTreeViewItem())                }#>            }            $twInstallSequence.AddHandler([System.Windows.Controls.TreeViewItem]::ExpandedEvent,$Event)        })
        
        $Form.ShowDialog() | Out-Null

    }
}