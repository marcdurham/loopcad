(defun pipe-draw (size)
	(setq old-osmode (getvar "OSMODE"))
	(setq old-orthomode (getvar "ORTHOMODE"))
	(defun *error* (message)
	  (princ)
	  (princ message)
	  (princ)
	  (setvar "OSMODE" old-osmode)
	  (setvar "ORTHOMODE" old-orthomode)
	  (command "-COLOR" "BYLAYER")
	  (command "-LAYER" "SET" "0" "")
	)
	(setvar "OSMODE" 64) ; 64 = Snap to insertion points
	(setvar "ORTHOMODE" 1)
	(command "-LAYER" "NEW" "Pipe" "")
	(command "-LAYER" "COLOR" "Blue" "Pipe" "")
	(command "-LAYER" "SET" "Pipe" "")
	(command "-COLOR" (pipe-size-color size))
	(while T
		(prompt (strcat "\nPipe Size: " size "\n"))
		(prompt "\nDraw pipe to each head.\n")
		(command "-PLINE" pause "Width" "2\"" "2\"" pause)
	)
)