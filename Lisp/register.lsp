; This won't work because the script doesn't know where it is, 
; or what folder it itself is in.
(setq pg-app-key "HKEY_CURRENT_USER\\Software\\ProgeCAD\\progeCAD\\")
(setq pg-ver-list (vl-registry-descendents pg-app-key))
(vl-registry-read 
    (strcat 
        pg-app-key 
        (last pg-ver-list) 
        "\\progeCAD-PRO:ENG\\Profiles\\Default\\Config") 
    "BLOCKSPATH"
)
