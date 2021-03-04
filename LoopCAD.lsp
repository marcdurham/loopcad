(prompt "\nLoading LoopCAD LISP modules...\n")
; Global Variables
(setq *failed-to-load* 0)

; Load LoopCAD LISP module files (*.lsp)
(setq loop-cad-folder (vl-filename-directory (findfile "LoopCAD.lsp")))
(setq loop-cad-lisp-folder (strcat loop-cad-folder "/Lisp"))
(load (strcat loop-cad-lisp-folder "/load-safely.lsp"))

(foreach f (cdr (cdr (vl-directory-files loop-cad-lisp-folder)))
    (if (= (strcase (vl-filename-extension f)) (strcase ".lsp"))
        (load-safely (strcat loop-cad-lisp-folder "/" f))
    )
)
(princ "\n")

; Check if files all loaded
(load-safely-check)

(define-labels)

(princ) ; exit quietly
