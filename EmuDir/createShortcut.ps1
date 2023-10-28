# this script allows the user to select a desired location for a game shortcut
#
# it works with this directory structure so far:
# /GameRoot
# Install Game Shortcut.bat -> bat script to launch ps script for the end user
#   /EmuDir
#      /games
#        Game_disk1.iso
#      createShortcut.ps1
#      [rest of emulator files/folders here]
#
# script is invoked with the "Install Game Shortcut.bat" script, or whatever its to be called

# add file browser functionality
Add-Type -AssemblyName System.Windows.Forms

# stolen from the answer for https://stackoverflow.com/questions/57547071/powershell-create-shortcut-to-network-printer, works well for this applicaton
# i should rewrite this further, it doesnt have to go this crazy
function New-Shortcut {
    [CmdletBinding()]  
    Param (   
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TargetPath,                # the path to the executable
        # the rest is all optional
        [string]$ShortcutPath = (Join-Path -Path ([Environment]::GetFolderPath("Desktop")) -ChildPath 'New Shortcut.lnk'),
        [string[]]$Arguments = $null,       # a string or string array holding the optional arguments.
        [string[]]$HotKey = $null,          # a string like "CTRL+SHIFT+F" or an array like 'CTRL','SHIFT','F'
        [string]$WorkingDirectory = $null,  
        [string]$Description = $null,
        [string]$IconLocation = $null,      # a string like "notepad.exe, 0", (can be a path to an image. 256x256 max res)
        [ValidateSet('Default','Maximized','Minimized')]
        [string]$WindowStyle = 'Default',
        [switch]$RunAsAdmin
    ) 
    switch ($WindowStyle) {
        'Default'   { $style = 1; break }
        'Maximized' { $style = 3; break }
        'Minimized' { $style = 7 }
    }
    $WshShell = New-Object -ComObject WScript.Shell

    # create a new shortcut
    $shortcut             = $WshShell.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath  = $TargetPath
    $shortcut.WindowStyle = $style
    if ($Arguments)        { $shortcut.Arguments = $Arguments -join ' ' }
    if ($HotKey)           { $shortcut.Hotkey = ($HotKey -join '+').ToUpperInvariant() }
    if ($IconLocation)     { $shortcut.IconLocation = $IconLocation }
    if ($Description)      { $shortcut.Description = $Description }
    if ($WorkingDirectory) { $shortcut.WorkingDirectory = $WorkingDirectory }

    # save the link file
    $shortcut.Save()

    if ($RunAsAdmin) {
        # read the shortcut file we have just created as [byte[]]
        [byte[]]$bytes = [System.IO.File]::ReadAllBytes($ShortcutPath)
        # $bytes[21] = 0x22      # set byte no. 21 to ASCII value 34
        $bytes[21] = $bytes[21] -bor 0x20 # set byte 21 bit 6 (0x20) ON
        [System.IO.File]::WriteAllBytes($ShortcutPath, $bytes)
    }

    # clean up the COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shortcut) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WshShell) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# store the current path of the script, should be /GameRoot or whatever GameRoot should correspond to
Write-Host "Creating shortcut for MyGame (Disk 1 and Disk 2). Select where you want the shortcut to be created"
$currPath = $(Get-Location).Path

# create an instanceof FolderBrowserDialog
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog

# to extend the script, rather than trying to bundle the game + emulator,
# we could share the script, allowing the user to select its emulator exe and rom to create a shortcut instead,
# although, this might be hard to include things like nice icons, but it might be fun to try and do that anyways
# file picker: System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = $currPath }; $filePicker.FileName

# start folder browser window
$null = $folderBrowser.ShowDialog();

# vars
# emulatorExePath - since we know the current directory of the script (the bat script root), EmuDir is implied.
# sould be replaced if the script lets the user select their own emulator exe
$emulatorExePath = "$currPath\EmuDir\My-Emulator.exe"

# path to the shortcut icon, stored in the EmurDir directory
$iconPath = "$currPath\EmuDir\gameicon.ico"

# set up parameter object for the New-Shortcut function above
# an array for multidisk games... dont need 2 disks? only use one object
$props = (
    @{
        # 'ShortcutPath' = Join-Path -Path ([Environment]::GetFolderPath("Desktop")) -ChildPath 'ConnectPrinter.lnk' - from stackoverflow
        # ensure these attributes are changed to your liking
        'ShortcutPath' = Join-Path -Path $folderBrowser.SelectedPath -ChildPath 'NAME-OF-SHORTCUT (Disk 1).lnk'
        'TargetPath'   = $emulatorExePath
        'Arguments'    = '-fullscreen', '-portable', '-slowboot', "$currPath\EmuDir\games\Game_disk1.cue" # duckstation arguments
        'IconLocation' = $iconPath
        'Description'  = 'Game (Disk 1)'
    },

    @{
        'ShortcutPath' = Join-Path -Path $folderBrowser.SelectedPath -ChildPath 'NAME-OF-SHORTCUT (Disk 2).lnk'
        'TargetPath'   = $emulatorExePath
        'Arguments'    = '-fullscreen', '-portable', '-slowboot', "$currPath\EmuDir\games\Game_disk2.cue"
        'IconLocation' = $iconPath
        'Description'  = 'Game (Disk 2)'
    }
)

foreach ($prop in $props) {
    New-Shortcut @prop
}

Write-Host "Shortcuts were placed in $($folderBrowser.SelectedPath)"
