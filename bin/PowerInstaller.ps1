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
    [String]$InstallString
    [String]$UninstallString
    hidden [string[]]$InstallArgumentList
    hidden [string[]]$UninstallArgumentList
    hidden [string]$Path
    hidden [string]$Tag
    hidden [String]$Header

    [bool]Install() {
        If (!($this.DetectionMethod.Detected())) {
            $RetVal = Start-Process -FilePath $this.Path -ArgumentList $this.InstallArgumentList -NoNewWindow -Wait -PassThru
            Write-Host ("Return value: " + $RetVal.ExitCode)
            return ($this.ValidExitCodes -contains $RetVal.ExitCode)
        } else {
            return $false
        }
    }

    [bool]Uninstall() {
        If ($this.DetectionMethod.Detected()) {
            $RetVal = Start-Process -FilePath $this.Path -ArgumentList $this.UninstallArgumentList -NoNewWindow -Wait -PassThru
            Write-Host ("Return value: " + $RetVal.ExitCode)
            return ($this.ValidExitCodes -contains $RetVal.ExitCode)
        } else {
            return $false
        }
    }

}

Class SetupInstaller : Installer {

    SetupInstaller([string]$Path, [String[]]$InstallArgumentList, [String[]]$UninstallArgumentList, [DetectionMethod]$DetectionMethod) {
        $this.Path = $Path
        $this.FilePath = $Path
        $this.DetectionMethod = $DetectionMethod
        $this.ValidExitCodes = @(0,3010)
        $this.InstallString = ($Path + $($out = "";$InstallArgumentList.ForEach({$out += (" " + $_)})))
        $this.UninstallString = ($Path + $($out = "";$UninstallArgumentList.ForEach({$out += (" " + $_)})))
        $this.Tag = [guid]::NewGuid()
        $this.Header = (Get-Item $Path).Name
    }

    SetupInstaller([string]$Path, [String[]]$InstallArgumentList, [String[]]$UninstallArgumentList, [DetectionMethod]$DetectionMethod, [string]$Name) {
        $this.Path = $Path
        $this.FilePath = $Path
        $this.DetectionMethod = $DetectionMethod
        $this.Name = $Name
        $this.ValidExitCodes = @(0,3010)
        $this.InstallString = ($Path + $($out = "";$InstallArgumentList.ForEach({$out += (" " + $_)})))
        $this.UninstallString = ($Path + $($out = "";$UninstallArgumentList.ForEach({$out += (" " + $_)})))
        $this.Tag = [guid]::NewGuid()
        $this.Header = $Name
    }

    SetupInstaller([string]$Path, [String[]]$InstallArgumentList, [String[]]$UninstallArgumentList, [DetectionMethod]$DetectionMethod, [string]$Name, [Version]$Version) {
        $this.Path = $Path
        $this.FilePath = $Path
        $this.DetectionMethod = $DetectionMethod
        $this.Name = $Name
        $this.Version = $Version
        $this.ValidExitCodes = @(0,3010)
        $this.InstallString = ($Path + $($out = "";$InstallArgumentList.ForEach({$out += (" " + $_)})))
        $this.UninstallString = ($Path + $($out = "";$UninstallArgumentList.ForEach({$out += (" " + $_)})))
        $this.Tag = [guid]::NewGuid()
        $this.Header = $Name
    }
    
}

Class MSIInstaller : Installer {

    [string]$ProductCode
    [string]$UpgradeCode
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
        $this.ProductCode = $ProductProperties.Item("ProductCode")
        $this.UpgradeCode = $ProductProperties.Item("UpgradeCode")
        $this.Name = $ProductProperties.Item("ProductName")
        $this.Version = $ProductProperties.Item("ProductVersion")

        $this.DetectionMethod = [DetectionMethod]::new("MSI", $this.ProductCode)
        $this.InstallArgumentList = @("-i", (Get-Item $Path).FullName, "-qb-", "ALLUSERS=1", "REBOOT=r")
        $this.UninstallArgumentList = @("-x", $ProductProperties.Item("ProductCode"), "-qb-", "ALLUSERS=1", "REBOOT=r")
        $this.InstallString = ($this.Path + $($out = "";($this.InstallArgumentList).ForEach({$out += (" " + $_)});$out))
        $this.UninstallString = ($this.Path + $($out = "";($this.UninstallArgumentList).ForEach({$out += (" " + $_)});$out))
        $this.Tag = [guid]::NewGuid()
        $this.Header = $this.Name
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
    hidden [String]$Tag
    hidden [String]$Header

    SequenceItem([Installer]$Installer) {
        $this.Name = $Installer.Name
        $NewInstallers = [System.Collections.ArrayList]::new()
        $NewInstallers.Add($Installer)
        $this.Installers = $NewInstallers
        $this.Tag = "Installers"
        $this.Header = $this.Name
        $this.Installers.ForEach({ $_.Tag = $this.Tag + "\" + $_.Tag })
    }

    SequenceItem([string]$Name, [Installer]$Installer) {
        $this.Name = $Name
        $NewInstallers = [System.Collections.ArrayList]::new()
        $NewInstallers.Add($Installer)
        $this.Installers = $NewInstallers
        $this.Tag = "Installers"
        $this.Header = $this.Name
        $this.Installers.ForEach({ $_.Tag = $this.Tag + "\" + $_.Tag })
    }

    SequenceItem([string]$Name, [Installer[]]$Installers) {
        $this.Name = $Name
        $this.Installers = $Installers
        $this.Tag = "Installers"
        $this.Header = $this.Name
        $this.Installers.ForEach({ $_.Tag = $this.Tag + "\" + $_.Tag })
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

}

Class PowerInstallerApp {

    PowerInstallerApp ($InstallSequence) {

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
        <TreeView x:Name="TreeView" HorizontalAlignment="Left" Height="300" Margin="10,10,0,0" VerticalAlignment="Top" Width="450"/>
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

        $TreeView = $Form.FindName('TreeView')
        try {
            $NewTreeViewItem = New-Object System.Windows.Controls.TreeViewItem
            $NewTreeViewItem.Header = $InstallSequence.Name
            $NewTreeViewItem.Tag = $InstallSequence
            [void]$TreeView.items.Add($NewTreeViewItem)
        } catch {
            Write-Host "Error adding $_ to the TreeView"
        }

        $Form.Add_SourceInitialized( {            [System.Windows.RoutedEventHandler]$Event = {                if($_.OriginalSource -is [System.Windows.Controls.TreeViewItem]){                    $TreeItem = $_.OriginalSource                    $TreeItem.items.clear()                    [PowerInstallerApp]::AddTreeViewItem($TreeItem)                }            }            $TreeView.AddHandler([System.Windows.Controls.TreeViewItem]::ExpandedEvent,$Event)        })
        
        $Form.ShowDialog() | Out-Null

    }

    static [void]AddTreeViewItem($TreeViewItem) {        $ListInstaller = $false        If ($TreeViewItem.Tag -ne $null) {            $ParentObject = $TreeViewItem.Tag            Switch ((($ParentObject).GetType()).Name) {                "InstallSequence" {                    ($ParentObject.Sequence).ForEach({
                        <#
                        # create stack panel
                        $StackPanel = [System.Windows.Controls.StackPanel]::new()
                        $StackPanel.Orientation = [System.Windows.Controls.Orientation]::Horizontal

                        # create Image
                        $Image = [System.Windows.Controls.Image]::new()
                        $Image.Source = [System.Windows.Media.Imaging.BitmapImage]::new([uri]::new("$PSScriptRoot\img\treeview\installsequence.bmp"))

                        # Label
                        $Label = [System.Windows.Controls.Label]::new()
                        $Label.Content = $_.Name

                        # Add into stack panel
                        $StackPanel.Children.Add($Image)
                        $StackPanel.Children.Add($Label)
                        #>
                        try {
                            $NewTreeViewItem = New-Object System.Windows.Controls.TreeViewItem
                            $NewTreeViewItem.Header = $_.Name
                            $NewTreeViewItem.Tag = $_
                            [void]$TreeViewItem.items.Add($NewTreeViewItem)
                        } catch {
                            Write-Host "Error adding $_ to the TreeView"
                        }                    })                }                "SequenceItem" {                    ($ParentObject.Installers).ForEach({
                        try {
                            $NewTreeViewItem = New-Object System.Windows.Controls.TreeViewItem
                            $NewTreeViewItem.Header = $_.Name
                            $NewTreeViewItem.Tag = $_
                            [void]$TreeViewItem.items.Add($NewTreeViewItem)
                        } catch {
                            Write-Host "Error adding $_ to the TreeView"
                        }                    })                }                "SetupInstaller" {                    $ListInstaller = $true                }                "MSIInstaller" {                    $ListInstaller = $true                }                "DetectionMethod" {                    $ListInstaller = $true                    }                default {
                    try {
                        $NewTreeViewItem = New-Object System.Windows.Controls.TreeViewItem
                        $NewTreeViewItem.Header = $ParentObject
                        $NewTreeViewItem.Tag = $null
                        [void]$TreeViewItem.items.Add($NewTreeViewItem)
                    } catch {
                        Write-Host "Error adding $_ to the TreeView"
                    }                                }            }
            If($ListInstaller) {                ForEach ($Property in (Get-Member -MemberType Property -InputObject ($TreeViewItem.Tag)).Name) {
                    try {
                        $NewTreeViewItem = New-Object System.Windows.Controls.TreeViewItem
                        $NewTreeViewItem.Header = $Property
                        $NewTreeViewItem.Tag = $ParentObject.($Property)
                        [void]$TreeViewItem.items.Add($NewTreeViewItem)
                    } catch {
                        Write-Host "Error adding $_ to the TreeView"
                    }                }            }
        }
    }

}