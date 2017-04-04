; Include all Excel addin files in the setup project.
Source: {#sourcedir}\*.xlam; DestDir: {code:GetDestDir}\; Check: ShouldInstallFile(12,16); AfterInstall: ActivateAddin(12,16)
Source: {#sourcedir}\*.xla; DestDir: {code:GetDestDir}\; Check: ShouldInstallFile(9,11); AfterInstall: ActivateAddin(9,11); Excludes: *.xlam
