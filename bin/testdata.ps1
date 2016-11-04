$MyInstaller = [MSIInstaller]::new("D:\DeploymentManager\DeploymentManager.msi")

$MySequenceItem = [SequenceItem]::new($MyInstaller)

$MyInstallSequence = [InstallSequence]::new("My Install Sequence", $MySequenceItem)

[PowerInstallerApp]::new($MyInstallSequence)

# Testing GitHub connector for Teams