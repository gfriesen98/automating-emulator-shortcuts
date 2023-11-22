# add file browser functionality
Add-Type -AssemblyName System.Windows.Forms

# stolen from the answer for https://stackoverflow.com/questions/57547071/powershell-create-shortcut-to-network-printer, works well for this applicaton
# i should rewrite this, it doesnt have to go this crazy
function New-Shortcut {
    [CmdletBinding()]
    Param (   
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TargetPath, # the path to the executable
        # the rest is all optional
        [string]$ShortcutPath = (Join-Path -Path ([Environment]::GetFolderPath("Desktop")) -ChildPath 'New Shortcut.lnk'),
        [string[]]$Arguments = $null, # a string or string array holding the optional arguments.
        [string[]]$HotKey = $null, # a string like "CTRL+SHIFT+F" or an array like 'CTRL','SHIFT','F'
        [string]$WorkingDirectory = $null,  
        [string]$Description = $null,
        [string]$IconLocation = $null, # a string like "notepad.exe, 0", (can be a path to an image .ico file.)
        [ValidateSet('Default', 'Maximized', 'Minimized')]
        [string]$WindowStyle = 'Default',
        [switch]$RunAsAdmin
    ) 
    switch ($WindowStyle) {
        'Default' { $style = 1; break }
        'Maximized' { $style = 3; break }
        'Minimized' { $style = 7 }
    }
    $WshShell = New-Object -ComObject WScript.Shell

    # create a new shortcut
    $shortcut = $WshShell.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $TargetPath
    $shortcut.WindowStyle = $style
    if ($Arguments) { $shortcut.Arguments = $Arguments -join ' ' }
    if ($HotKey) { $shortcut.Hotkey = ($HotKey -join '+').ToUpperInvariant() }
    if ($IconLocation) { $shortcut.IconLocation = $IconLocation }
    if ($Description) { $shortcut.Description = $Description }
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

# find image for psx games from https://psxdatacenter.com
function Get-PsxDatacenterImage {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [string]$RomName,
        [Parameter(Mandatory = $true)]
        [string]$Serial,
        [Parameter(Mandatory = $true)]
        [string]$Region,
        [Parameter(Mandatory = $true)]
        [string]$SaveLocation
    )

    # example url https://psxdatacenter.com/images/covers/U/0-9/SLUS-01300.jpg
    $baseUrl = "https://psxdatacenter.com/images/covers"
    $regionCode = switch ($Region) {
        USA { 'U' }
        Japan { 'J' }
        Europe { 'P' }
        Germany { 'P' }
        France { 'P' }
        Default { 'U' }
    }

    # parse the rom filename for the first character
    $categoryCode = "0-9"
    if ($RomName.ToUpper() -match '^(\d+)') {
        $categoryCode = "0-9"
    }
    elseif ($RomName.ToUpper() -match '^[A-Z]') {
        $categoryCode = $RomName.ToUpper()[0]
    }

    # multidisk games should re-use disk 1's thumbnail if it exists
    if ($RomName -match '\s\(Dis. [2-9]\)') {
        if ([System.IO.File]::Exists("$SaveLocation.ico")) {
            Write-Host "Using existing thumbnail"
            return $null
        }
    }

    Write-Host "Getting thumbnail for $Serial"
    Write-Host "URL: $baseUrl/$regionCode/$categoryCode/$Serial.jpg"
    Invoke-WebRequest "$baseUrl/$regionCode/$categoryCode/$Serial.jpg" -OutFile "$SaveLocation.jpg"
    # .\convert.exe "$SaveLocation.jpg" -resize 256x256 "$SaveLocation.ico"
    .\convert.exe "$SaveLocation.jpg" "$SaveLocation.ico"
    # Remove-Item -Path "$SaveLocation.jpg"
}

# search http://redump.org for psx game serial codes (like its SLES code or whatever, this is needed for psxdatacenter.com as thats how they name their game icon images)
function Search-Redump {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$RomName,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SelectedRegion
    )

    Write-Host "RomName $RomName"
    $replace = $($RomName -replace "\s\(Dis. [0-9]\)", '').toLower().trim() # redump's search wont show results if you have (Disk X) in the name... should just be the name of the game
    $replace = $($replace -replace "\s", '-') # url param inevitably becomes lowercase and spaces replaced with -
    Write-Host "Searching for serial code..."
    Write-Host "Gathering data for $replace ..."

    $res = Invoke-WebRequest "http://redump.org/discs/quicksearch/$replace"
    $parsed = $res.ParsedHtml

    # first 8 are the table headers
    # its insane to me that powershell can do this like why
    $tr = $parsed.getElementsByTagName('tr')
    $results = New-Object System.Collections.Generic.List[System.Object] # store correct results in a list
    for ($j = 0; $j -le $tr.length - 1; $j++) {
        # /discs/quicksearch/$replace html returns a <table> that contains the data we need, theres probably a better way to write this but whatever
        $region = [String]$tr[$j].children[0].children[0].title
        $name = [String]$tr[$j].children[1].children[0].innerhtml
        $serial = [String]$tr[$j].children[6].innerhtml
        if ($region -eq $SelectedRegion) {
            $results.Add(@{
                "Region" = $region
                "Name" = $name
                "Serial" = $serial
            })
        }
    }

    $r = $results.ToArray()
    for ($i = 0; $i -le $r.Length - 1; $i++) {
        Write-Host "[$i]: " $r[$i].Name $r[$i].Serial.trim().split('&')[0]
    }

    # multidisk psx games typically use disk 1's icon, because im lazy i just ask the user to select disk 1 for the icon
    if ($RomName -match '\s\(Dis. [2-9]\)') {
        Write-Host "It appears this is not the primary disk. An icon may not be available for this disk."
        Write-Host "!! If available, Disk 1's icon will be reused. Select Disk 1 regardless !!"
        $selection = Read-Host "Select best result"
        Write-Host $r[$selection]
        return $r[$selection].Serial.trim().split('&')[0]
    } else {
        $selection = Read-Host "Select best result"
        Write-Host $r[$selection]
        return $r[$selection].Serial.trim().split('&')[0]
    }
}

# start initial script

$currPath = $(Get-Location).Path

Write-Host "Select emulator executable: "
$emulatorPicker = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = $currPath }
$null = $emulatorPicker.ShowDialog()
$emulatorPath = $emulatorPicker.FileName
Write-Host "Selected $emulatorPath"

Write-Host "Select rom file"
$romPicker = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = $currPath }
$null = $romPicker.ShowDialog()
$romPath = $romPicker.FileName
Write-Host "Selected $romPath"

$createIcon = Read-Host "Create icon? (y/n)"
if ($createIcon -eq 'y') {
    Write-Host "Select region:"
    Write-Host "`t[0] USA"
    Write-Host "`t[1] Japan"
    Write-Host "`t[2] Europe"
    Write-Host "`t[3] France"
    Write-Host "`t[4] Germany"
    $regionIndex = Read-Host "Enter selection"
    $selectedRegion = switch ($regionIndex) {
        "0" { 'USA' }
        "1" { 'Japan' }
        "2" { 'Europe' }
        "3" { 'France' }
        "4" { 'Germany' }
        Default { 'USA' }
    }
    Write-Host $selectedRegion
    $romName = [System.IO.Path]::GetFileNameWithoutExtension($romPath)

    $serialCode = Search-Redump -RomName $romName -SelectedRegion $selectedRegion
    Write-Host "Found $serialCode"

    $iconLocation = Split-Path -Parent $emulatorPath
    $iconLocation = "$iconLocation\$serialCode"

    # create folder for the thumbnail
    # New-Item -ItemType Directory -Path "$iconLocation\$serialCode"

    Get-PsxDatacenterImage -RomName $romName -Region $selectedRegion -Serial $serialCode -SaveLocation $iconLocation

    Write-Host "Select shortcut save location"
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $null = $folderBrowser.ShowDialog();

    $prop = @{
        'ShortcutPath' = Join-Path -Path $folderBrowser.SelectedPath -ChildPath "$romName.lnk"
        'TargetPath'   = $emulatorPath
        'Arguments'    = '-fullscreen', '-portable', '-slowboot', $romPath
        'IconLocation' = "$iconLocation.ico"
        'Description'  = $romName
    }
    New-Shortcut @prop
} else {
    $prop = @{
        'ShortcutPath' = Join-Path -Path $folderBrowser.SelectedPath -ChildPath "$romName.lnk"
        'TargetPath'   = $emulatorPath
        'Arguments'    = '-fullscreen', '-portable', '-slowboot', $romPath
        'Description'  = $romName
    }
    New-Shortcut @prop
}