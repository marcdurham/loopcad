# This won't work if the user hasn't enabled powershell scripts on their system.
#.\Register-LoopCAD-Paths.ps1 : File E:\LoopCAD\Register-LoopCAD-Paths.ps1 cannot be loaded
# because running scripts is disabled on this system.
$cd = Get-Location
$path1 = "hkcu:\Software\ProgeCAD\progeCAD"
$path2 = (Get-ChildItem $path1).Name
$configPath = "${path2}\progeCAD-PRO:ENG\Profiles\Default\Config".Replace("HKEY_CURRENT_USER\", "hkcu:")
$config = (Get-ItemProperty -Path "$configPath")

$existingBlocksPath = $config.BLOCKSPATH 
$newBlocksPath = "$cd\Blocks"
if ($existingBlocksPath.Contains("$newBlocksPath")) {
    "Blocks Path already configured."
} else {
    "Adding LoopCAD folder to Blocks Path..."
    Set-ItemProperty -Path "$configPath" -Name "BLOCKSPATH" -Value "$existingBlocksPath;$newBlocksPath"
}

$menuPath = $config.MENUPATH
if ($menuPath.Contains("$cd")) {
    "Menu Path already configured."
} else {
    "Adding LoopCAD folder to Menu Path..."
    Set-ItemProperty -Path "$configPath" -Name "MENUPATH" -Value "$menuPath;$cd"
}

# Computer\HKEY_CURRENT_USER\SOFTWARE\Autodesk\AutoCAD\R24.0\ACAD-4101:409\Profiles\<<Unnamed Profile>>\General