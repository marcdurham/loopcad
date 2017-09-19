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
(load-safely "load-safely-check.lsp")

; Tests
(load-safely "data-request-test.lsp")
(load-safely-check)

; Default
(data-submit "DefaultHeadModel" "NEW")

(princ) ; exit quietly