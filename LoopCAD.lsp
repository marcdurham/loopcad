(prompt "\nLoading LoopCAD LISP modules...\n")
; Global Variables
(setq *failed-to-load* 0)


; Load LoopCAD LISP module files (*.lsp)
(load "Lisp\\load-safely.lsp")

(foreach f (cdr (cdr (vl-directory-files ".\\Lisp")))
	(if (str= (vl-filename-extension f) ".lsp")
		(load-safely (strcat ".\\Lisp\\" f))
	)
)
(princ "\n")


; Check if files all loaded
(load-safely-check)

(princ) ; exit quietly