rm .\combined.lsp
ls .\Lisp\*.lsp |
 % { 
        echo "; *****************************************" >> .\combined.lsp
        echo "; File: $($_.FullName)" >> .\combined.lsp
        echo "; *****************************************" >> .\combined.lsp
        cat $_.FullName >> .\combined.lsp 
    }

echo "; *****************************************" >> .\combined.lsp
echo "; File: LoopCAD.lsp" >> .\combined.lsp
echo "; *****************************************" >> .\combined.lsp
cat LoopCAD.lsp >> .\combined.lsp