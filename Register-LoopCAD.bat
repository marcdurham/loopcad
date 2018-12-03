%SYSTEMROOT%\Microsoft.NET\Framework\v4.0.30319\regasm.exe LoopCAD.dll /tlb:LoopCAD.tlb /codebase
%SYSTEMROOT%\Microsoft.NET\Framework\v4.0.30319\regasm.exe FileConverter.dll /tlb:FileConverter.tlb /codebase
%SYSTEMROOT%\Microsoft.NET\Framework\v4.0.30319\regasm.exe LoopCalc.Core.dll /tlb:LoopCalc.Core.tlb /codebase
%SYSTEMROOT%\Microsoft.NET\Framework\v4.0.30319\regasm.exe LoopFlow.Common.dll /tlb:LoopFlow.Common.tlb /codebase
make-link.exe -target:"C:\Program Files (x86)\ProgeCAD\progeCAD 2014 Professional ENG\icad.exe" -output:LoopCAD.lnk -arguments:"/b C:\LoopCAD\LoopCAD.scr"
powershell .\Register-LoopCAD-Paths.ps1
pause


REM Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\progeCAD.Drawing\protocol\StdFileEditing\server

REM Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\shell\open\command
REM Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\progeCAD.Drawing\shell\open\command
REM Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\progeCAD.DXF.Drawing\protocol\StdFileEditing\server