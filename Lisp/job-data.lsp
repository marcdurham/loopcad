(defun job-data-insert ( / old-osmode)
	(if (= (length (get-blocks (list "JobData" "Job Data"))) 0)
		(progn
			(setq old-osmode (getvar "OSMODE"))
			(defun *error* (message)
				(princ)
				(princ message)
				(princ)
				(setvar "OSMODE" old-osmode)
				(command "-LAYER" "SET" "0" "")
				(setvar "LWDISPLAY" 1)
			)
			(setvar "INSUNITS" 0) ;This line prevents inserted block refs from having a different scale, being 12 time bigger than they should be
			(setvar "OSMODE" 0)
			(command "-LAYER" "NEW" "JobData" "")
			(command "-LAYER" "COLOR" "140" "JobData" "")
			(command "-LAYER" "OFF" "JobData" "")
			(command "-LAYER" "SET" "JobData" "")
			(insert-job-data-block '(0 0 0))
		)
	)
	(vl-vbarun "ScanJobData")
	(vl-vbarun "EditJobData")
)

(defun insert-job-data-block ( point )
	(entmake
		(list 
			(cons 0 "INSERT")
			(cons 10 point)     ; Insertion point
			(cons 2 "JobData")  ; Block name
			(cons 8 "JobData")  ; Layer
			(cons 66 1)         ; Attributes follow
		)
	)
	(setq e (entlast))
	; Outer Circle
	(entmake
		(list
			(cons 0 "CIRCLE")      
			(cons 10 (list 0 0 0)) ; Center Point
			; Radius: 5.0 copied from old block so it looks the same
			(cons 40 4.5)        ; Radius
			(cons 62 color-cyan)  ; Color
			(cons 8 "JobData") 		  ; Layer
		)
	)
	(job-data-attribute "JOB_NUMBER" "L-TEST" 1)
	(job-data-attribute "JOB_NAME" "" 2)
	(job-data-attribute "JOB_SITE_ADDRESS" "" 3)
	(job-data-attribute "CALCULATED_BY_COMPANY" "" 4)
	(job-data-attribute "SUPPLY_STATIC_PRESSURE" "0" 5)
	(job-data-attribute "SUPPLY_RESIDUAL_PRESSURE" "0" 6)
	(job-data-attribute "SUPPLY_AVAILABLE_FLOW" "0" 7)
	(job-data-attribute "SUPPLY_ELEVATION" "0" 8)
	(job-data-attribute "SUPPLY_PIPE_LENGTH" "0" 9)
	(job-data-attribute "SUPPLY_PIPE_INTERNAL_DIAMETER" "0" 10)
	(job-data-attribute "SPRINKLER_PIPE_TYPE" "" 11)
	(job-data-attribute "SPRINKLER_FITTING_TYPE" "" 12)
	(job-data-attribute "SUPPLY_PIPE_TYPE" "" 13)
	(job-data-attribute "SUPPLY_PIPE_SIZE" "0" 14)
	(job-data-attribute "SUPPLY_NAME" "MTR" 15)
	(job-data-attribute "DOMESTIC_FLOW_ADDED" "0" 16)
	(job-data-attribute "WATER_FLOW_SWITCH_MAKE_MODEL" "" 17)
	(job-data-attribute "SUPPLY_PIPE_FITTINGS_SUMMARY" "" 18)
	(job-data-attribute "SUPPLY_PIPE_FITTINGS_EQUIV_LENGTH" "0" 19)
	(job-data-attribute "SUPPLY_PIPE_ADD_PRESSURE_LOSS" "0" 20)
	(job-data-attribute "WATER_FLOW_SWITCH_PRESSURE_LOSS" "0" 21)
	(entmake
		(list 
			(cons 0 "SEQEND") 
			;(cons -2 e)
		)
	)
	;(entupd e)
	(princ "\nBlock inserted\n")
	(princ)
)


(defun job-data-attribute ( tag-string text-value position / y-offset )
	(setq y-offset (* 10.0 position))
	(entmake
		(list 
			(cons 0 "ATTRIB") ; Entity type
			(cons 10 (list 0.0 (- 0.0 y-offset))) ; Label insertion point
			(cons 1 text-value)    ; Text value
			(cons 2 tag-string)    ; Tag string
			(cons 40 5.0)          ; Text height
			(cons 7 "ARIAL")       ; Text style
			(cons 62 color-blue)   ; Color
			(cons 8 "JobData")     ; Layer
		)
	)
)

(defun job-data-attdef ( tag-string text-value attribute-prompt position / y-offset )
	(setq y-offset (* 10.0 position))
	(entmake
		(list 
			(cons 0 "ATTDEF") ; Entity type
			(cons 10 (list 0.0 (- 0.0 y-offset))) ; Label insertion point
			(cons 1 text-value)    ; Text value
			(cons 2 tag-string)    ; Tag string
			(cons 3 attribute-prompt) ; Prompt string
			(cons 40 5.0)          ; Text height
			(cons 7 "ARIAL")       ; Text style
			(cons 62 color-blue)   ; Color
			(cons 8 "JobData")     ; Layer
		)
	)
)