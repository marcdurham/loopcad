; TODO: Maybe have two parameters: model-code and filename, use a separate function to determine them.
(defun head-insert (model coverage slope temperature / model-code pt)
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
  (setvar "INSUNITS" 0) ; This line prevents inserted block refs from having a
                        ; different scale, being 12 times bigger than they should be.
  (setvar "OSMODE" 0)
  (command "-LAYER" "NEW" "Heads" "")
  (command "-LAYER" "NEW" "HeadCoverage" "")
  (command "-LAYER" "COLOR" "Red" "Heads" "")
  (command "-LAYER" "COLOR" "Yellow" "HeadCoverage" "")
  (command "-LAYER" "ON" "HeadCoverage" "")
  (setvar "LWDISPLAY" 0)
  (command "-LAYER" "SET" "Heads" "")
  (while T
	(setq model-code "HEAD-X")
;
;   This section is for heads that you already know the coverage
;   Now the default is that you don't know. 
;   See the head-insert.lsp file and the head-insert function. 
;
;   (setq model-code (model-code-from model coverage slope temperature))
;	(prompt (strcat "\nInserting Head Model Code: " model-code "\n"))
;   (prompt "\nPress Esc to quit inserting heads.\n")
;	(command "-INSERT" (strcat "Head" coverage ".dwg") pause 1.0 1.0 0 model-code)
;
	(command "-INSERT" (strcat "Head12-20.dwg") pause 1.0 1.0 0 model-code)
	(setq pt (cdr (assoc 10 (entget (entlast)))))
	(if (null global:head-coverage)
		(setq global:head-coverage "16")
	)
	(initget "12 14 16 18 20")
	(if (setq tmp (getkword (strcat "\nHead Coverage [12/14/16/18/20] <" global:head-coverage ">: ")))
		(setq global:head-coverage tmp)
	)
	(entdel (entlast))
	(setq model-code (model-code-from model global:head-coverage slope temperature))
	(prompt (strcat "\nInserting Head Model Code: " model-code "\n"))
    (prompt "\nPress Esc to quit inserting heads.\n")
	(command "-INSERT" (strcat "Head" global:head-coverage ".dwg") pt 1.0 1.0 0 model-code)
  )
)

(defun swhead-insert (direction model temperature / model-code pt tmp)
	(princ "\nSWHEAD-INSERT dir: ")
	(princ direction)
	(princ " model: ")
	(princ mode)
	(princ " temp: ")
	(princ temperature)
	(princ "\n")
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
  (setvar "INSUNITS" 0) ; This line prevents inserted block refs from having a
                        ; different scale, being 12 times bigger than they should be.
  (setvar "OSMODE" 0)
  (command "-LAYER" "NEW" "Heads" "")
  (command "-LAYER" "NEW" "HeadCoverage" "")
  (command "-LAYER" "COLOR" "Red" "Heads" "")
  (command "-LAYER" "COLOR" "Yellow" "HeadCoverage" "")
  (command "-LAYER" "ON" "HeadCoverage" "")
  (setvar "LWDISPLAY" 0)
  (command "-LAYER" "SET" "Heads" "")
  (while T
	(setq model-code "HEAD-X")
	(princ "\nHEAD-X happening now...")
	(command 
		"-INSERT" ; Command
		(strcat "SwHead12-20" global:head-spray-direction ".dwg") ; Block name
		pause ; Get insertion point
		1.0 ; X scale
		1.0 ; Y scale
		0 ; Rotation 
		model-code ; Model Code
	)
	(setq pt (cdr (assoc 10 (entget (entlast)))))
	(if (null global:head-coverage)
		(setq global:head-coverage "16")
	)
	(initget "12 14 16 18 20")
	(if (setq tmp (getkword (strcat "\nHead Coverage [12/14/16/18/20] <" global:head-coverage ">: ")))
		(setq global:head-coverage tmp)
	)
	(entdel (entlast))
	(setq model-code (model-code-from model global:head-coverage "" temperature))
	(prompt (strcat "\nInserting Head Model Code: " model-code "\n"))
    (prompt "\nPress Esc to quit inserting heads.\n")
	(command "-INSERT" (strcat "SwHead" global:head-coverage global:head-spray-direction) pt 1.0 1.0 0 model-code)
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

(defun head-insert-user ( / tmp) 
	(if (null global:head-model)
		(setq global:head-model "RFC43")
	)
	(if (setq tmp (getstring (strcat "\nHead Model <" global:head-model ">: ")))
		(setq global:head-model tmp)
	)
;
;   This section is for heads that you already know the coverage
;   Now the default is that you don't know. 
;   See the head-insert.lsp file and the head-insert function. 
;
;   (if (null global:head-coverage)
;		(setq global:head-coverage "16")
;	)
;	(initget "12 14 16 18 20")
;	(if (setq tmp (getkword (strcat "\nHead Coverage [12/14/16/18/20] <" global:head-coverage ">: ")))
;		(setq global:head-coverage tmp)
;	)
;
	(if (null global:head-slope)
		(setq global:head-slope "")
	)
	(if (setq tmp (getstring (strcat "\nHead Slope <" global:head-slope ">: ")))
		(setq global:head-slope tmp)
	)
	(if (null global:head-temperature)
		(setq global:head-temperature "")
	)
	(if (setq tmp (getstring (strcat "\nHead Temperature <" global:head-temperature ">: ")))
		(setq global:head-temperature tmp)
	)
    (head-insert
		global:head-model
		"20" ; global:head-coverage
		global:head-slope
		global:head-temperature
	)
)

; Insert a side wall head, prompt user for specs
(defun swhead-insert-user ( / tmp) 
	(if (null global:head-spray-direction)
		(setq global:head-spray-direction "U")
	)
	;(if (setq tmp (getstring (strcat "\nHead Spray Direction <" global:head-spray-direction ">: ")))
	;	(setq global:head-spray-direction tmp)
	;)
	(initget "U D L R")
	(setq 
		global:head-spray-direction
		(getkword 
			(strcat 
				"\nHead Spray Direction [U/D/L/R] <" 
				global:head-spray-direction
				">: "
			)
		)
	)
	;(setq global:head-spray-direction tmp)
	
	(if (null global:head-model)
		(setq global:head-model "RFC43")
	)
	(if (setq tmp (getstring (strcat "\nSidewall Head Model <" global:head-model ">: ")))
		(setq global:head-model tmp)
	)
	(if (null global:head-temperature)
		(setq global:head-temperature "")
	)
	(if (setq tmp (getstring (strcat "\nHead Temperature <" global:head-temperature ">: ")))
		(setq global:head-temperature tmp)
	)
    (swhead-insert
		global:head-spray-direction
		global:head-model
		;"20" ; global:head-coverage
		; No slope for side wall heads
		global:head-temperature
	)
)


(defun head-insert-coverage (coverage) 
    (head-insert 
	    (data-request "DefaultHeadModel")
		coverage
		(data-request "DefaultHeadSlope")
		(data-request "DefaultHeadTemperature")
	)
)