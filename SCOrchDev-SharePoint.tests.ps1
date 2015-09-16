$here = Split-Path -Parent $MyInvocation.MyCommand.Path

$manifestPath = "$here\SCOrchDev-SharePoint.psd1"
Import-Module SCOrchDev-SharePoint -Force

Describe -Tags 'VersionChecks' 'SCOrchDev-PasswordVault manifest' {
    $script:manifest = $null
    It 'has a valid manifest' {
        {
            $script:manifest = Test-ModuleManifest -Path $manifestPath -ErrorAction Stop -WarningAction SilentlyContinue
        } | Should Not Throw
    }

    It 'has a valid name in the manifest' {
        $script:manifest.Name | Should Be SCOrchDev-SharePoint
    }

    It 'has a valid guid in the manifest' {
        $script:manifest.Guid | Should Be 'adfd9a51-2775-4886-9803-ff96ddfb1cf2'
    }

    It 'has a valid version in the manifest' {
        $script:manifest.Version -as [Version] | Should Not BeNullOrEmpty
    }

    if (Get-Command git.exe -ErrorAction SilentlyContinue) {
        $script:tagVersion = $null
        It 'is tagged with a valid version' {
            $cwd = get-location
            Set-Location ($Path -as [System.IO.FileInfo]).Directory
            $thisCommit = git.exe log --decorate --oneline HEAD~1..HEAD
            Set-Location $cwd
            if ($thisCommit -match 'tag:\s*(\d+(?:\.\d+)*)')
            {
                $script:tagVersion = $matches[1]
            }

            

            $script:tagVersion                  | Should Not BeNullOrEmpty
            $script:tagVersion -as [Version]    | Should Not BeNullOrEmpty
            
        }

        It 'all versions are the same' {
            $script:manifest.Version -as [Version] | Should be ( $script:tagVersion -as [Version] )
        }

    }

    It 'should have all files listed in the FileList' {
        $ModuleFiles = (Get-ChildItem -Path $here -Recurse -Exclude .git).FullName
        $FileDifferences = Compare-Object -ReferenceObject $ModuleFiles -DifferenceObject $script:manifest.FileList
        
        if (($FileDifferences -as [array]).Count -gt 0)
        {
            Throw-Exception -Type 'MissingFiles' `
                            -Message 'Files missing or not tracked in FileList' `
                            -Property @{
                'Missing Files' = ($FileDifferences | Where-Object {$_.SideIndicator -eq '=>'}).InputObject ;
                'Non Tracked Files' = ($FileDifferences | Where-Object {$_.SideIndicator -eq '<='}).InputObject ;
            }
        }
    }
}

if ($PSVersionTable.PSVersion.Major -ge 3)
{
    $error.Clear()
    Describe 'Clean treatment of the $error variable' {
        It 'Performs a successful test' {
            $true | Should Be $true
        }

        It 'Did not add anything to the $error variable' {
            $error.Count | Should Be 0
        }
    }
}

Describe 'Style rules' {
    $SCOrchDevUtilityRoot = (Get-Module SCOrchDev-SharePoint).ModuleBase

    $files = @(
        Get-ChildItem $SCOrchDevUtilityRoot -Include *.ps1,*.psm1
    )

    It 'Module source files contain no trailing whitespace' {
        $badLines = @(
            foreach ($file in $files)
            {
                $lines = [System.IO.File]::ReadAllLines($file.FullName)
                $lineCount = $lines.Count

                for ($i = 0; $i -lt $lineCount; $i++)
                {
                    if ($lines[$i] -match '\s+$')
                    {
                        'File: {0}, Line: {1}' -f $file.FullName, ($i + 1)
                    }
                }
            }
        )

        if ($badLines.Count -gt 0)
        {
            throw "The following $($badLines.Count) lines contain trailing whitespace: `r`n`r`n$($badLines -join "`r`n")"
        }
    }

    It 'Module Source Files all end with a newline' {
        $badFiles = @(
            foreach ($file in $files)
            {
                $string = [System.IO.File]::ReadAllText($file.FullName)
                if ($string.Length -gt 0 -and $string[-1] -ne "`n")
                {
                    $file.FullName
                }
            }
        )

        if ($badFiles.Count -gt 0)
        {
            throw "The following files do not end with a newline: `r`n`r`n$($badFiles -join "`r`n")"
        }
    }
}