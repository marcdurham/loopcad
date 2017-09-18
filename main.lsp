(prompt "\nLoading LoopCAD LISP modules...\n")

; Load LoopCAD LISP module files
(load "load-safely.lsp")
(load-safely "request-data.lsp")
(load-safely "submit-data.lsp")
(load-safely "insert-head.lsp")
(load-safely "insert-head-user.lsp")
(load-safely "insert-head-coverage.lsp")
(load-safely "set-head-data.lsp")
(load-safely "pipe-size-color.lsp")
(load-safely "pipe-draw.lsp")
(load-safely "commands.lsp")

; Tests
(load-safely "request-data-test.lsp")

(submit-data "DefaultHeadModel" "NEW")

; Show error if any modules failed to load
(if (> failed-to-load 0)
	(progn
		(setq message (strcat "ERROR: " (itoa failed-to-load) " LoopCAD LISP module files failed to load! Check command box for the names of the specific files."))
		(alert message)
		(prompt (strcat "\n" message "\n"))
	)
	(progn
		(prompt "\nAll LoopCAD LISP module files were loaded successfully.\n")
	)
)

(princ) ; Exit quietly