(defun load-safely (file-name)
	(prompt (strcat "\nLoading module: \"" file-name "\"\n"))
    (setq file-was-found (findfile file-name))
	(if file-was-found
	    (load file-name)
		(progn 
			(setq *failed-to-load* (1+ *failed-to-load*))
			(prompt (strcat "\nERROR: LoopCAD LISP module file failed to load: \"" file-name "\"\n"))
		)
	)
)

; Show error if any modules failed to load
(defun load-safely-check ( / message)
    (if (> *failed-to-load* 0)
        (progn
            (setq message (strcat "ERROR: " (itoa *failed-to-load*) " LoopCAD LISP module files failed to load! Check command box for the names of the specific files."))
            (alert message)
            (prompt "\n**** ERROR ****\n")
            (prompt (strcat "\n" message "\n"))
        )
        (progn
            (prompt "\n*****************************************************************\n")
            (prompt "\n**** All LoopCAD LISP module files were loaded successfully. ****\n")
            (prompt "\n*****************************************************************\n")
        )
    )
)