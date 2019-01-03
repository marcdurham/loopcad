; TODO: Maybe have two parameters: model-code and filename, use a separate function to determine them.
(defun head-insert (model coverage slope temperature / model-code)
  (setq old-osmode (getvar "OSMODE"))
  (setq temperror *error*)
  (defun *error* (message)
	  (princ)
	  (princ message)
	  (princ)
    (setvar "OSMODE" old-osmode)
	(command "-LAYER" "OFF" "HeadCoverage" "")
    (setvar "LWDISPLAY" 1)
	(setq *error* temperror)
  )
  (setvar "INSUNITS" 0) ;This line prevents inserted block refs from having a
                        ;different scale, being 12 times bigger than they should be.
  (setvar "OSMODE" 0)
  (command "-LAYER" "NEW" "Heads" "")
  (command "-LAYER" "NEW" "HeadCoverage" "")
  (command "-LAYER" "COLOR" "Red" "Heads" "")
  (command "-LAYER" "COLOR" "Yellow" "HeadCoverage" "")
  (command "-LAYER" "ON" "HeadCoverage" "")
  (setvar "LWDISPLAY" 0)
  (command "-LAYER" "SET" "Heads" "")
  (while T
    (setq model-code (model-code-from model coverage slope temperature))
	(prompt (strcat "\nInserting Head Model Code: " model-code "\n"))
    (prompt "\nPress Esc to quit inserting heads.\n")
    (command "-INSERT" (strcat "Head" coverage ".dwg") pause 1.0 1.0 0 model-code)
  )
)

(defun model-code-from (model coverage slope temperature)
	(cond
		(
			(and 
				;(and model coverage slope temperature) 
				(> (strlen model) 0)
				(> (strlen coverage) 0)
				(> (strlen slope) 0)
				(> (strlen temperature) 0)
			)
			(strcat model "-" coverage "-" slope "-" temperature)
		)
		(
			(and 
				(> (strlen model) 0)
				(> (strlen coverage) 0)
				(> (strlen slope) 0)
			)
			(strcat model "-" coverage "-" slope)
		)
		(			
			(and 
				(> (strlen model) 0)
				(> (strlen coverage) 0)
			)
			(strcat model "-" coverage)
		)
		(
			(> (strlen model) 0)
			(strcat model)
		)
	)
)