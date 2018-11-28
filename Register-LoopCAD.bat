%SYSTEMROOT%\Microsoft.NET\Framework\v4.0.30319\regasm.exe LoopCAD.dll /tlb:LoopCAD.tlb /codebase
make-link.exe -target:"C:\Program Files (x86)\ProgeCAD\progeCAD 2016 Professional ENG\icad.exe" -output:LoopCAD.lnk -arguments:"/b C:\LoopCAD\LoopCAD.scr"f
pause


REM Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\progeCAD.Drawing\protocol\StdFileEditing\server

REM Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\shell\open\command
REM Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\progeCAD.Drawing\shell\open\command
REM Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\progeCAD.DXF.Drawing\protocol\StdFileEditing\server