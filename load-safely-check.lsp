; Show error if any modules failed to load
(defun load-safely-check ( / message)
    (if (> *failed-to-load* 0)
        (progn
            (setq message (strcat "ERROR: " (itoa *failed-to-load*) " LoopCAD LISP module files failed to load! Check command box for the names of the specific files."))
            (alert message)
            (prompt "\n**** ERROR ****\n"S)
            (prompt (strcat "\n" message "\n"))
        )
        (progn
            (prompt "\n*****************************************************************\n")
            (prompt "\n**** All LoopCAD LISP module files were loaded successfully. ****\n")
            (prompt "\n*****************************************************************\n")
        )
    )
)