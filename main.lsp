(prompt "\nLoading LoopCAD LISP modules...\n")
; Global Variables
(setq *failed-to-load* 0)

; Load LoopCAD LISP module files
(load "load-safely.lsp")
(load-safely "data-request.lsp")
(load-safely "data-submit.lsp")
(load-safely "head-insert.lsp")
(load-safely "head-insert-user.lsp")
(load-safely "head-insert-coverage.lsp")
(load-safely "head-data-set.lsp")
(load-safely "pipe-size-color.lsp")
(load-safely "pipe-draw.lsp")
(load-safely "tee-insert.lsp")
(load-safely "commands.lsp")

; Tests
(load-safely "data-request-test.lsp")

(data-submit "DefaultHeadModel" "NEW")

; Show error if any modules failed to load
(if (> *failed-to-load* 0)
	(progn
		(setq message (strcat "ERROR: " (itoa *failed-to-load*) " LoopCAD LISP module files failed to load! Check command box for the names of the specific files."))
		(alert message)
		(prompt (strcat "\n" message "\n"))
	)
	(progn
		(prompt "\nAll LoopCAD LISP module files were loaded successfully.\n")
	)
)

(princ) ; Exit quietly