# This won't work if the user hasn't enabled powershell scripts on their system.
#.\Register-LoopCAD-Paths.ps1 : File E:\LoopCAD\Register-LoopCAD-Paths.ps1 cannot be loaded
# because running scripts is disabled on this system.
$cd = Get-Location
$loopCADPath = $PSScriptRoot
"LoopCAD Path: $loopCADPath"
$currentVersionMajor = (Get-ItemProperty -Path "hkcu:\SOFTWARE\Autodesk\AutoCAD\").CurVer
$currentVersionMinor = (Get-ItemProperty -Path "hkcu:\SOFTWARE\Autodesk\AutoCAD\$currentVersionMajor").CurVer
$profileKey = "hkcu:\SOFTWARE\Autodesk\AutoCAD\$currentVersionMajor\$currentVersionMinor\Profiles\<<Unnamed Profile>>\General"
#$path2 = (Get-ChildItem $path1).Name
#$configPath = "${path2}\progeCAD-PRO:ENG\Profiles\Default\Config".Replace("HKEY_CURRENT_USER\", "hkcu:")
$acadPaths = (Get-ItemProperty -Path $profileKey).ACAD

If(($acadPaths -split ";") -notcontains $loopCADPath) {
    Write-Host "Adding LoopCAD path ($loopCADPath) to AutoCAD Support Paths..."
    Set-ItemProperty -Path $profileKey -Name "ACAD" -Value "$acadPaths$loopCADPath;"
} else {
    Write-Host "LoopCAD path ($loopCADPath) already exists in AutoCAD Support Paths."
}

# Add Blocks to Support Path
$blockPath = "$loopCADPath\Blocks"
If(($acadPaths -split ";") -notcontains $blockPath ) {
    Write-Host "Adding LoopCAD Block path ($blockPath ) to AutoCAD Support Paths..."
    Set-ItemProperty -Path $profileKey -Name "ACAD" -Value "$acadPaths$blockPath;"
} else {
    Write-Host "LoopCAD Block path ($blockPath) already exists in AutoCAD Support Paths."
}

$supportFolderKey = "hkcu:\SOFTWARE\Autodesk\AutoCAD\$currentVersionMajor\$currentVersionMinor"
$supportFolderPath = (Get-ItemProperty -Path $supportFolderKey).SupportFolder

Write-Host "Support Folder: $supportFolderPath"

$acaddocPath = "$supportFolderPath/acaddoc.lsp"
If(!(Test-Path $acaddocPath)) {
    Write-Host "acaddoc.lsp file does not exist, creating..."
    ";New Auto-Start LISP file created by LoopCAD" | 
        Out-File $acaddocPath -Encoding utf8
}

$acaddoc = Get-Content $acaddocPath
If($acaddoc -notcontains "(load `"LoopCAD.lsp`")") {
    Write-Host "Adding a line to load LoopCAD.lsp to $acaddocPath..."
    "(load `"LoopCAD.lsp`")" | 
        Out-File -FilePath $acaddocPath -Append -Encoding utf8
} else {
    Write-Host "acaddoc.lsp file already contains a line to load LoopCAD.lsp"
    Write-Host "No changes made"
}

