(prompt "\nLoading LoopCAD LISP modules...\n")
; Global Variables
(setq *failed-to-load* 0)


; Load LoopCAD LISP module files (*.lsp)
(load "Lisp\\load-safely.lsp")
(load-safely "Lisp\\contains.lsp")
(load-safely "Lisp\\math.lsp")
(load-safely "Lisp\\data-request.lsp")
(load-safely "Lisp\\data-submit.lsp")
(load-safely "Lisp\\data-change-default.lsp")
(load-safely "Lisp\\head-insert.lsp")
(load-safely "Lisp\\head-insert-user.lsp")
(load-safely "Lisp\\head-insert-coverage.lsp")
(load-safely "Lisp\\head-data-set.lsp")
(load-safely "Lisp\\entity-creators.lsp")
(load-safely "Lisp\\break-pipes.lsp")
(load-safely "Lisp\\pipe-size-color.lsp")
(load-safely "Lisp\\pipe-draw.lsp")
(load-safely "Lisp\\tee-insert.lsp")
(load-safely "Lisp\\head-join.lsp")
(load-safely "Lisp\\elevation-box.lsp")
(load-safely "Lisp\\job-data-insert.lsp")
(load-safely "Lisp\\commands.lsp")
(load-safely "Lisp\\load-safely-check.lsp")
(load-safely "job_data.lsp")

; Check if files all loaded
(load-safely-check)

(princ) ; exit quietly