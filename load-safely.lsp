(defun load-safely (file-name)
    (setq file-was-found (findfile file-name))
	(if file-was-found
	    (load file-name)
		(progn 
			(setq *failed-to-load* (1+ *failed-to-load*))
			(prompt (strcat "LoopCAD LISP module file failed to load: " file-name "\n"))
		)
	)
)