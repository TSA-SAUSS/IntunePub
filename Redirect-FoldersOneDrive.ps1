# Requires -Version 3
<#
    .SYNOPSIS
        Creates a scheduled task to implement folder redirection for.

    .NOTES
        Name: Redirect-Folders.ps1
        Author: Aaron Parker
        Site: https://stealthpuppy.com
        Twitter: @stealthpuppy
#>
[CmdletBinding(ConfirmImpact = 'Low', HelpURI = 'https://stealthpuppy.com/', SupportsPaging = $False,
    SupportsShouldProcess = $False, PositionalBinding = $False)]
param ()

# Log file
$VerbosePreference = "Continue"
$stampDate = Get-Date
$scriptName = ([System.IO.Path]::GetFileNameWithoutExtension($(Split-Path $script:MyInvocation.MyCommand.Path -Leaf)))
$LogFile = "$env:LocalAppData\IntuneScriptLogs\$scriptName-" + $stampDate.ToFileTimeUtc() + ".log"
Start-Transcript -Path $LogFile

Function Set-KnownFolderPath {
    <#
        .SYNOPSIS
            Sets a known folder's path using SHSetKnownFolderPath.
        .PARAMETER KnownFolder
            The known folder whose path to set.
        .PARAMETER Path
            The target path to redirect the folder to.
        .NOTES
            Forked from: https://gist.github.com/semenko/49a28675e4aae5c8be49b83960877ac5
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet('AddNewPrograms', 'AdminTools', 'AppUpdates', 'CDBurning', 'ChangeRemovePrograms', 'CommonAdminTools', 'CommonOEMLinks', 'CommonPrograms', `
                'CommonStartMenu', 'CommonStartup', 'CommonTemplates', 'ComputerFolder', 'ConflictFolder', 'ConnectionsFolder', 'Contacts', 'ControlPanelFolder', 'Cookies', `
                'Desktop', 'Documents', 'Downloads', 'Favorites', 'Fonts', 'Games', 'GameTasks', 'History', 'InternetCache', 'InternetFolder', 'Links', 'LocalAppData', `
                'LocalAppDataLow', 'LocalizedResourcesDir', 'My Music', 'NetHood', 'NetworkFolder', 'OriginalImages', 'PhotoAlbums', 'Pictures', 'Playlists', 'PrintersFolder', `
                'PrintHood', 'Profile', 'ProgramData', 'ProgramFiles', 'ProgramFilesX64', 'ProgramFilesX86', 'ProgramFilesCommon', 'ProgramFilesCommonX64', 'ProgramFilesCommonX86', `
                'Programs', 'Public', 'PublicDesktop', 'PublicDocuments', 'PublicDownloads', 'PublicGameTasks', 'PublicMusic', 'PublicPictures', 'PublicVideos', 'QuickLaunch', `
                'Recent', 'RecycleBinFolder', 'ResourceDir', 'RoamingAppData', 'SampleMusic', 'SamplePictures', 'SamplePlaylists', 'SampleVideos', 'SavedGames', 'SavedSearches', `
                'SEARCH_CSC', 'SEARCH_MAPI', 'SearchHome', 'SendTo', 'SidebarDefaultParts', 'SidebarParts', 'StartMenu', 'Startup', 'SyncManagerFolder', 'SyncResultsFolder', `
                'SyncSetupFolder', 'System', 'SystemX86', 'Templates', 'TreeProperties', 'UserProfiles', 'UsersFiles', 'Videos', 'Windows')]
        [System.String] $KnownFolder,

        [Parameter(Mandatory = $true)]
        [System.String] $Path
    )

    # Define known folder GUIDs
    $KnownFolders = @{
        'Contacts'       = '56784854-C6CB-462b-8169-88E350ACB882';
        'Cookies'        = '2B0F765D-C0E9-4171-908E-08A611B84FF6';
        'Desktop'        = @('B4BFCC3A-DB2C-424C-B029-7FE99A87C641');
        'Documents'      = @('FDD39AD0-238F-46AF-ADB4-6C85480369C7', 'f42ee2d3-909f-4907-8871-4c22fc0bf756');
        'Downloads'      = @('374DE290-123F-4565-9164-39C4925E467B', '7d83ee9b-2244-4e70-b1f5-5393042af1e4');
        'Favorites'      = '1777F761-68AD-4D8A-87BD-30B759FA33DD';
        'Games'          = 'CAC52C1A-B53D-4edc-92D7-6B2E8AC19434';
        'GameTasks'      = '054FAE61-4DD8-4787-80B6-090220C4B700';
        'History'        = 'D9DC8A3B-B784-432E-A781-5A1130A75963';
        'InternetCache'  = '352481E8-33BE-4251-BA85-6007CAEDCF9D';
        'InternetFolder' = '4D9F7874-4E0C-4904-967B-40B0D20C3E4B';
        'Links'          = 'bfb9d5e0-c6a9-404c-b2b2-ae6db6af4968';
        'Music'          = @('4BD8D571-6D19-48D3-BE97-422220080E43', 'a0c69a99-21c8-4671-8703-7934162fcf1d');
        'NetHood'        = 'C5ABBF53-E17F-4121-8900-86626FC2C973';
        'OriginalImages' = '2C36C0AA-5812-4b87-BFD0-4CD0DFB19B39';
        'PhotoAlbums'    = '69D2CF90-FC33-4FB7-9A0C-EBB0F0FCB43C';
        'Pictures'       = @('33E28130-4E1E-4676-835A-98395C3BC3BB', '0ddd015d-b06c-45d5-8c4c-f59713854639');
        'QuickLaunch'    = '52a4f021-7b75-48a9-9f6b-4b87a210bc8f';
        'Recent'         = 'AE50C081-EBD2-438A-8655-8A092E34987A';
        'RoamingAppData' = '3EB685DB-65F9-4CF6-A03A-E3EF65729F3D';
        'SavedGames'     = '4C5C32FF-BB9D-43b0-B5B4-2D72E54EAAA4';
        'SavedSearches'  = '7d1d3a04-debb-4115-95cf-2f29da2920da';
        'StartMenu'      = '625B53C3-AB48-4EC1-BA1F-A1EF4146FC19';
        'Templates'      = 'A63293E8-664E-48DB-A079-DF759E0509F7';
        'Videos'         = @('18989B1D-99B5-455B-841C-AB7C74E4DDFC', '35286a68-3c57-41a1-bbb1-0eae73d76c95');
    }

    # Define SHSetKnownFolderPath if it hasn't been defined already
    $Type = ([System.Management.Automation.PSTypeName]'KnownFolders').Type
    If (-not $Type) {
        $Signature = @'
[DllImport("shell32.dll")]
public extern static int SHSetKnownFolderPath(ref Guid folderId, uint flags, IntPtr token, [MarshalAs(UnmanagedType.LPWStr)] string path);
'@
        $Type = Add-Type -MemberDefinition $Signature -Name 'KnownFolders' -Namespace 'SHSetKnownFolderPath' -PassThru
    }

    # Make path, if doesn't exist
    If (!(Test-Path $Path -PathType Container)) {
        if ($PSCmdlet.ShouldProcess($Path, ("New-Item '{0}'" -f $Path))) {
            New-Item -Path $Path -Type "Directory" -Force -Verbose
        }
    }

    # Validate the path
    If (Test-Path $Path -PathType Container) {
        # Call SHSetKnownFolderPath
        #  return $Type::SHSetKnownFolderPath([ref]$KnownFolders[$KnownFolder], 0, 0, $Path)
        ForEach ($guid in $KnownFolders[$KnownFolder]) {
            Write-Verbose "Redirecting $KnownFolders[$KnownFolder]"
            $result = $Type::SHSetKnownFolderPath([ref]$guid, 0, 0, $Path)
            If ($result -ne 0) {
                $errormsg = "Error redirecting $($KnownFolder). Return code $($result) = $((New-Object System.ComponentModel.Win32Exception($result)).message)"
                Throw $errormsg
            }
        }
    }
    Else {
        Throw New-Object System.IO.DirectoryNotFoundException "Could not find part of the path $Path."
    }

    # Fix up permissions, if we're still here
    Attrib +r $Path
    Write-Output $Path
}

Function Get-KnownFolderPath {
    <#
        .SYNOPSIS
            Gets a known folder's path using GetFolderPath.
        .PARAMETER KnownFolder
            The known folder whose path to get. Validates set to ensure only knwwn folders are passed.
        .NOTES
            https://stackoverflow.com/questions/16658015/how-can-i-use-powershell-to-call-shgetknownfolderpath
    #>
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateSet('AdminTools', 'ApplicationData', 'CDBurning', 'CommonAdminTools', 'CommonApplicationData', 'CommonDesktopDirectory', 'CommonDocuments', 'CommonMusic', `
                'CommonOemLinks', 'CommonPictures', 'CommonProgramFiles', 'CommonProgramFilesX86', 'CommonPrograms', 'CommonStartMenu', 'CommonStartup', 'CommonTemplates', `
                'CommonVideos', 'Cookies', 'Desktop', 'DesktopDirectory', 'Favorites', 'Fonts', 'History', 'InternetCache', 'LocalApplicationData', 'LocalizedResources', 'MyComputer', `
                'MyDocuments', 'MyMusic', 'MyPictures', 'MyVideos', 'NetworkShortcuts', 'Personal', 'PrinterShortcuts', 'ProgramFiles', 'ProgramFilesX86', 'Programs', 'Recent', `
                'Resources', 'SendTo', 'StartMenu', 'Startup', 'System', 'SystemX86', 'Templates', 'UserProfile', 'Windows')]
        [System.String] $KnownFolder
    )
    [Environment]::GetFolderPath($KnownFolder)
}

Function Redirect-Folder {
    <#
        .SYNOPSIS
            Function exists to reduce code required to redirect each folder.
    #>
    Param (
        [Parameter(Mandatory = $true)]
        [System.String] $SyncFolder,

        [Parameter(Mandatory = $true)]
        [System.String] $GetFolder,

        [Parameter(Mandatory = $true)]
        [System.String] $SetFolder,

        [Parameter(Mandatory = $true)]
        [System.String] $Target
    )

    # Get current Known folder path
    $Folder = Get-KnownFolderPath -KnownFolder $GetFolder

    # If paths don't match, redirect the folder
    If ($Folder -ne "$SyncFolder\$Target") {
        # Redirect the folder
        Write-Verbose "Redirecting $SetFolder to $SyncFolder\$Target"
        Set-KnownFolderPath -KnownFolder $SetFolder -Path "$SyncFolder\$Target"

        # Move files/folders into the redirected folder
        Write-Verbose "Moving data from $SetFolder to $SyncFolder\$Target"
        Move-File -Source $Folder -Destination "$SyncFolder\$Target" -Log "$env:LocalAppData\RedirectLogs\Robocopy$Target.log"

        # Hide the source folder (rather than delete it)
        Attrib +h $Folder
    }
    Else {
        Write-Verbose "Folder $GetFolder matches target. Skipping redirection."
    }
}

Function Invoke-Process {
    <#PSScriptInfo
        .VERSION 1.4
        .GUID b787dc5d-8d11-45e9-aeef-5cf3a1f690de
        .AUTHOR Adam Bertram
        .COMPANYNAME Adam the Automator, LLC
        .TAGS Processes
    #>
    <#
    .DESCRIPTION
        Invoke-Process is a simple wrapper function that aims to "PowerShellyify" launching typical external processes. There
        are lots of ways to invoke processes in PowerShell with Start-Process, Invoke-Expression, & and others but none account
        well for the various streams and exit codes that an external process returns. Also, it's hard to write good tests
        when launching external proceses.

        This function ensures any errors are sent to the error stream, standard output is sent via the Output stream and any
        time the process returns an exit code other than 0, treat it as an error.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [System.String] $FilePath,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [System.String] $ArgumentList
    )
    $ErrorActionPreference = 'Stop'

    try {
        $stdOutTempFile = "$env:TEMP\$((New-Guid).Guid)"
        $stdErrTempFile = "$env:TEMP\$((New-Guid).Guid)"

        $startProcessParams = @{
            FilePath               = $FilePath
            ArgumentList           = $ArgumentList
            RedirectStandardError  = $stdErrTempFile
            RedirectStandardOutput = $stdOutTempFile
            Wait                   = $true;
            PassThru               = $true;
            NoNewWindow            = $true;
        }
        if ($PSCmdlet.ShouldProcess("Process [$($FilePath)]", "Run with args: [$($ArgumentList)]")) {
            $cmd = Start-Process @startProcessParams
            $cmdOutput = Get-Content -Path $stdOutTempFile -Raw
            $cmdError = Get-Content -Path $stdErrTempFile -Raw
            if ($cmd.ExitCode -ne 0) {
                if ($cmdError) {
                    throw $cmdError.Trim()
                }
                if ($cmdOutput) {
                    throw $cmdOutput.Trim()
                }
            }
            else {
                if ([System.String]::IsNullOrEmpty($cmdOutput) -eq $false) {
                    Write-Output -InputObject $cmdOutput
                }
            }
        }
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
    finally {
        Remove-Item -Path $stdOutTempFile, $stdErrTempFile -Force -ErrorAction Ignore
    }
}

Function Move-File {
    <#
        .SYNOPSIS
            Moves contents of a folder with output to a log.
            Uses Robocopy to ensure data integrity and all moves are logged for auditing.
            Means we don't need to re-write functionality in PowerShell.
        .PARAMETER Source
            The source folder.
        .PARAMETER Destination
            The destination log.
        .PARAMETER Log
            The log file to store progress/output
    #>
    Param (
        $Source,
        $Destination,
        $Log
    )
    If (!(Test-Path (Split-Path $Log))) { New-Item -Path (Split-Path $Log) -ItemType Container }
    Write-Verbose "Moving data in folder $Source to $Destination."
    Robocopy.exe "$Source" "$Destination" /E /MOV /XJ /XF *.ini /R:1 /W:1 /NP /LOG+:$Log
}


# Get OneDrive sync folder
$SyncFolder = Get-ItemPropertyValue -Path 'HKCU:\Software\Microsoft\OneDrive\Accounts\Business1' -Name 'UserFolder' -ErrorAction SilentlyContinue
Write-Verbose "Target sync folder is $SyncFolder."

# Redirect select folders
If (Test-Path -Path $SyncFolder -ErrorAction SilentlyContinue) {
    # Redirect-Folder -SyncFolder $SyncFolder -GetFolder 'Desktop' -SetFolder 'Desktop' -Target 'Desktop'
    # Redirect-Folder -SyncFolder $SyncFolder -GetFolder 'MyDocuments' -SetFolder 'Documents' -Target 'Documents'
    Redirect-Folder -SyncFolder $SyncFolder -GetFolder 'MyMusic' -SetFolder 'Music' -Target 'My Music'
    Redirect-Folder -SyncFolder $SyncFolder -GetFolder 'MyVideos' -SetFolder 'Videos' -Target 'Videos'
}
Else {
    Write-Verbose "$SyncFolder does not (yet) exist. Skipping folder redirection until next logon."
}

Stop-Transcript -Verbose
# SIG # Begin signature block
# MIITHQYJKoZIhvcNAQcCoIITDjCCEwoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDucT3GdKXv/Bgs
# RXTPW366Xfm/bk34KZKkwbt/WV41oaCCEE4wgge0MIIFnKADAgECAhMoAAsmXzhC
# zDU6HDO0AAEACyZfMA0GCSqGSIb3DQEBCwUAMFwxCzAJBgNVBAYTAlVTMR8wHQYD
# VQQKExZUaGUgU2FsdmF0aW9uIEFybXkgVVNTMSwwKgYDVQQDEyNUaGUgU2FsdmF0
# aW9uIEFybXkgU0FVU1MgSXNzdWluZyBDQTAeFw0yMzAxMTkyMjEzNDJaFw0yNTAx
# MTgyMjEzNDJaMCIxIDAeBgNVBAMTF1VTUyBUSFEgSVQgTmV0d29yayBUZWFtMIIB
# IjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAp0Brsf7uWF542ikCU4NivNf/
# E0WYPoQqf2tGUlg8ulGSbSAlj69Y6BmoQdRYsfkCE7HrB9OnfdzTQQ0OKUpOcWmn
# rh4k99oOB2KIsBr4zIIdHiWPBS3PeUNvE7/d48R38t8HUrqxq+PqE98Ggc4N3pMO
# VmB1+F5GQIRNjm13Sl7p8C2vS/XjSbUbO0JaHn5d4wKl65AHoHO8dapgDm8DgVf4
# /Bi12+ZSjzbX6Df7z7Gp1mDt/xlJMLoVW636YxP1/Cj3bOou2oxeA6jZShXbMoRy
# W4c0rrIgPdBeOsdO4QLTuICv1bSWFxYO3mjBdLOzEnGO+PXs2t6bdyiG/heMVQID
# AQABo4IDpzCCA6MwPAYJKwYBBAGCNxUHBC8wLQYlKwYBBAGCNxUIgdafFITK4kiE
# mYc/h7+LN++efwSDhK4ah/vYZAIBZAIBBjATBgNVHSUEDDAKBggrBgEFBQcDAzAO
# BgNVHQ8BAf8EBAMCB4AwGwYJKwYBBAGCNxUKBA4wDDAKBggrBgEFBQcDAzAdBgNV
# HQ4EFgQUphxAwmVDDc93lDMMMvyC0F59XaQwHwYDVR0jBBgwFoAU6M3OIp5rAeXV
# 7Lfax2OHI+ZF52cwggFRBgNVHR8EggFIMIIBRDCCAUCgggE8oIIBOIaB5mxkYXA6
# Ly8vQ049VGhlJTIwU2FsdmF0aW9uJTIwQXJteSUyMFNBVVNTJTIwSXNzdWluZyUy
# MENBKDEpLENOPVNBVVNTQ0FJc3N1aW5nMSxDTj1DRFAsQ049UHVibGljJTIwS2V5
# JTIwU2VydmljZXMsQ049U2VydmljZXMsQ049Q29uZmlndXJhdGlvbixEQz1zYXVz
# cyxEQz1zYXVzcyxEQz1uZXQ/Y2VydGlmaWNhdGVSZXZvY2F0aW9uTGlzdD9iYXNl
# P29iamVjdENsYXNzPWNSTERpc3RyaWJ1dGlvblBvaW50hk1odHRwOi8vcGtpLnNh
# dXNzLm5ldC9wa2kvVGhlJTIwU2FsdmF0aW9uJTIwQXJteSUyMFNBVVNTJTIwSXNz
# dWluZyUyMENBKDEpLmNybDCCAYoGCCsGAQUFBwEBBIIBfDCCAXgwgdMGCCsGAQUF
# BzAChoHGbGRhcDovLy9DTj1UaGUlMjBTYWx2YXRpb24lMjBBcm15JTIwU0FVU1Ml
# MjBJc3N1aW5nJTIwQ0EsQ049QUlBLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2Vz
# LENOPVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9c2F1c3MsREM9c2F1c3Ms
# REM9bmV0P2NBQ2VydGlmaWNhdGU/YmFzZT9vYmplY3RDbGFzcz1jZXJ0aWZpY2F0
# aW9uQXV0aG9yaXR5MHkGCCsGAQUFBzAChm1odHRwOi8vcGtpLnNhdXNzLm5ldC9w
# a2kvU0FVU1NDQUlzc3VpbmcxLnNhdXNzLnNhdXNzLm5ldF9UaGUlMjBTYWx2YXRp
# b24lMjBBcm15JTIwU0FVU1MlMjBJc3N1aW5nJTIwQ0EoMSkuY3J0MCUGCCsGAQUF
# BzABhhlodHRwOi8vcGtpLnNhdXNzLm5ldC9vY3NwMA0GCSqGSIb3DQEBCwUAA4IC
# AQAQEmNxSuUTt3lHRsp/C7IfBGSK3zkYL/uWPqy1k0mmYrhoT0dsMJLZbaQuLAkc
# I8YpHl9DitswS96JfaxDUkHBt97d3LQbKR6HoYewWjcUuTY+ulFKmha0x7bd7hnK
# 5hSsUJNPvy0uoHGSVqPY6g6iCslF4wo5fN1UWFfvnr/Hz0KXrPSvZXxLmW3QZqFJ
# 3bUcRo43AceN4bjeqQ9gWyazEDn6oNwjZi89LRVFzUZTF+maRE+cPa+592UZe/vC
# o02ivC8+e/M+nlfxqCRNjhldlhPp4rACsLfG2ZJ2wKGFnrjD1WAOi0sJsMMy+8fe
# cduH5suoHSLrMykuT874DXerjARtzAaZs2CCUarh6/SDvZyR4Z1CmPMZ2fiR5JxF
# vjrKV7IOIezQ0ohuXW1ZfaEBsxFllA66urxvqk8k4XvtqwkITpluGN62AXt6xKzC
# VO8U4FZlF8IdWDjS1S+j/IM3L4qS9kqsTjQfsbGGlcLCIlzyz17C70eINkooOEMD
# oHuLHU8UmRgwsRw8bcNKGk1CMgKriMoEbyszeYKQLlJ5eLHsfpBFx2g7l6r9NdhL
# Quz6o8fAjoQ4FTC867YiUieOM0IsPJA7nh5ODyWJi6pVkF4NjzMLTRaKoAAB/E1s
# TpD9dn5YiAmaEESerLwIFwmLqYQM7vTKJyGq3sFIy4XWSzCCCJIwggZ6oAMCAQIC
# ExkAAAAGAD8dMuV32ckAAQAAAAYwDQYJKoZIhvcNAQELBQAwWTELMAkGA1UEBhMC
# VVMxHzAdBgNVBAoTFlRoZSBTYWx2YXRpb24gQXJteSBVU1MxKTAnBgNVBAMTIFRo
# ZSBTYWx2YXRpb24gQXJteSBTQVVTUyBSb290IENBMB4XDTE4MTAwMTE0MjQ0NVoX
# DTI4MTAwMTE0MzQ0NVowXDELMAkGA1UEBhMCVVMxHzAdBgNVBAoTFlRoZSBTYWx2
# YXRpb24gQXJteSBVU1MxLDAqBgNVBAMTI1RoZSBTYWx2YXRpb24gQXJteSBTQVVT
# UyBJc3N1aW5nIENBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA2hwq
# AMAKLx3adv3Txg+smbMQu4kqZCOYFOLm/yaIwsM+lq4p7u5ookyNvLZozmwbmst4
# cYgSuvAVOyYiw4yRDisDLdtM2QJrGql6rXK+4B9bbNd+LvCN42bJrKpsgfrG5QgL
# nt8ijBUJ0EBzFgUgp/wev6WcR7EirtbXYYNed4MjF01zCbZ0OiAsekjC94FvITnt
# Qs3b89hXHazWwZOPPty7Gq5To+DSo5WxC9kVn4UxHGgExCoUPw6THaXRdcQGAAwO
# kIESZY1WDqK8+WhOru2U5mD0ZK5sg8Ng1DJlJc1fkqy/B7O4yLN9rSAPFuRnStUf
# zC/By2XhNGEJzUanoH5lTRrdNVmFXIr0bKr2zlPpWNlU7tpnUw/i+uA62BWoIxuN
# Pj3e6D/PleZ88RjS9auk9TnHb/Vc47w9X86dyp1S9/VftUpmYi9wvsUkavCRW6M4
# GBBJBdNIy+zGr8NUcTZOfVlFYsSqymCdMZyDj0A+uoqhnnXEdD1llu+w0eQrwSCt
# s5qWmSIkNx1d7U5uJH3GuRf4aEB0BttxkbuzqAMxbqHzCl1dKyA1l3jXsaTkF6+Z
# SdZL1XpZNZfWoCPluuAOYSXjTJzBaxDB751zU8YZND0vaS+g6CseWabXTQvvgvDL
# WYtkf7csAi2w4e6ks6LG+xTS+GviX5xfC3KfXXcCAwEAAaOCA04wggNKMBIGCSsG
# AQQBgjcVAQQFAgMBAAEwIwYJKwYBBAGCNxUCBBYEFLU7o3IbkN789uiygkcFSo70
# G+pVMB0GA1UdDgQWBBTozc4inmsB5dXst9rHY4cj5kXnZzAZBgkrBgEEAYI3FAIE
# DB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNV
# HSMEGDAWgBQyD7AkPqGPZ1rQ/mmDe/PU3is/jjCCAUcGA1UdHwSCAT4wggE6MIIB
# NqCCATKgggEuhoHfbGRhcDovLy9DTj1UaGUlMjBTYWx2YXRpb24lMjBBcm15JTIw
# U0FVU1MlMjBSb290JTIwQ0EoMSksQ049U0FVU1NSb290Q0EsQ049Q0RQLENOPVB1
# YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRp
# b24sREM9c2F1c3MsREM9c2F1c3MsREM9bmV0P2NlcnRpZmljYXRlUmV2b2NhdGlv
# bkxpc3Q/YmFzZT9vYmplY3RDbGFzcz1jUkxEaXN0cmlidXRpb25Qb2ludIZKaHR0
# cDovL3BraS5zYXVzcy5uZXQvcGtpL1RoZSUyMFNhbHZhdGlvbiUyMEFybXklMjBT
# QVVTUyUyMFJvb3QlMjBDQSgxKS5jcmwwggFJBggrBgEFBQcBAQSCATswggE3MIHQ
# BggrBgEFBQcwAoaBw2xkYXA6Ly8vQ049VGhlJTIwU2FsdmF0aW9uJTIwQXJteSUy
# MFNBVVNTJTIwUm9vdCUyMENBLENOPUFJQSxDTj1QdWJsaWMlMjBLZXklMjBTZXJ2
# aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0aW9uLERDPXNhdXNzLERDPXNh
# dXNzLERDPW5ldD9jQUNlcnRpZmljYXRlP2Jhc2U/b2JqZWN0Q2xhc3M9Y2VydGlm
# aWNhdGlvbkF1dGhvcml0eTBiBggrBgEFBQcwAoZWaHR0cDovL3BraS5zYXVzcy5u
# ZXQvcGtpL1NBVVNTUm9vdENBX1RoZSUyMFNhbHZhdGlvbiUyMEFybXklMjBTQVVT
# UyUyMFJvb3QlMjBDQSgxKS5jcnQwDQYJKoZIhvcNAQELBQADggIBAEL72W26vJI9
# NrjDrT2URoRjFDL4Bj4xEYZJ5tQEx3U4hiD/3NPaSYOkO+Hjjg7f5MaiM1XpV99b
# WIl2detJFL3eqzzXEUlJ6WwWsOzbO9toWTAW16HuUSqdgjhwiPuAq18/DBO9sgTd
# IK46VuAVX7Vn649NwPraNpckSgekPP7tLr/ToNX9ikzBPVgB2qhTPjaeGsD9s5eS
# Ioj3hMayMlUJ9e34THxVfSzssQeE9Bf9QSCaH5yBxIpmoFvPN7QG6iawI+OBNWFM
# cAaPyysm0bBRjqEnmt37T0Oaun235wicyF6U2fzzxxB9YtnzaQMTfsDFeAtN7wqF
# 60e8JC1PqXiGiLkQDYbof37pkXC3An2ltFFvFVmogFogzEsulVIIsCNEO/PqXBXM
# G7UkELmXua5qPM9TeIcx0hCY96iPKCnMAvuEDD/4k3xJkoS6PxWcfD2c+9Auk/da
# sfLrFWoqKrLn2wC2BM+gfTO2UVcoBbi1op4wF3dzh4Wops2McLHs+b5bJOL1KmPa
# 5V1hPmUaQDbIC48rOLYFSCgfmSu3WCo2CK6yZ/6zj54oBldcR51B4Y96FimM11Sq
# fkeRRmZrY4eH4b67sVhDxPZm1IaEUfOf46RKd4hgdJHz2Fd2bPyTBQqzwBjqTJpl
# 0+0bhb+UrqXM6fr9LjFSCmoat6a7GWqEMYICJTCCAiECAQEwczBcMQswCQYDVQQG
# EwJVUzEfMB0GA1UEChMWVGhlIFNhbHZhdGlvbiBBcm15IFVTUzEsMCoGA1UEAxMj
# VGhlIFNhbHZhdGlvbiBBcm15IFNBVVNTIElzc3VpbmcgQ0ECEygACyZfOELMNToc
# M7QAAQALJl8wDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAA
# oQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4w
# DAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgkhfUw3WQEPGZeCADRXgh7BMT
# +VB737WSLFaP34BFcMkwDQYJKoZIhvcNAQEBBQAEggEAVznjEyt1zChORWK94oPc
# ow/lQZudDaARss5QEgt7TcL+ymrjLT0QtQd4S9BVrhyocvK/TKHYjRm/I06TOaZi
# PZIi0d0W6sI0IuweKpNP9EJCs4UuIFKM6DIKgUqybVc0dDVAbLdRA+5yayfAjS1+
# 5ZuaflGSrHbqq2nUotsdaIqxHu5VrW8E/w229HTIIB4/m+6NJyr/Bgh6gxORNElb
# mlURwBn4P5Px37E5DLFrKNSWLh1NOzn+KMK2YLjREPRAW4r5IRhBZcYqHPd2tDld
# Kti4JpIVNfH0Y2kyu1JlACD7IKnkRqA6Se3QWytYP7TAJ7S2lBm5u/OwIP7NIYqO
# qw==
# SIG # End signature block
