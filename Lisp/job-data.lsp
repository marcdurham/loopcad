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
			(command "-LAYER" "ON" "JobData" "")
			(command "-LAYER" "SET" "JobData" "")
			(prompt "\nClick a location, to insert job data.\n")
			(command "-INSERT" "JobData.dwg" 0 0 0 1.0 1.0 0
			  ; These attributes are in a strange, as-created, order
			  ""  ; JOB_NAME
			  ""  ; JOB_NUMBER
			  "0" ; SUPPLY_STATIC_PRESSURE
			  ""  ; JOB_SITE_ADDRESS
			  "0" ; SUPPLY_RESIDUAL_PRESSURE
			  "0" ; SUPPLY_AVAILABLE_FLOW
			  "0" ; SUPPLY_ELEVATION
			  "0" ; SUPPLY_PIPE_LENGTH
			  "0" ; SUPPLY_PIPE_INTERNAL_DIAMETER
			  "X-Fire"  ; CALCULATED_BY_COMPANY
			  ""    ; SPRINKLER_PIPE_TYPE
			  ""    ; SPRINKLER_FITTING_TYPE
			  ""    ; SUPPLY_PIPE_TYPE
			  "0"   ; SUPPLY_PIPE_SIZE
			  "MTR" ; SUPPLY_NAME
			  "0"   ; DOMESTIC_FLOW_ADDED
			  ""    ; WATER_FLOW_SWITCH_MAKE_MODEL
			  ""    ; SUPPLY_PIPE_FITTINGS_SUMMARY
			  "0"   ; SUPPLY_PIPE_FITTINGS_EQUIV_LENGTH
			  "0"   ; SUPPLY_PIPE_ADD_PRESSURE_LOSS
			  "0"   ; WATER_FLOW_SWITCH_PRESSURE_LOSS
			)
		)
	)
	(vl-vbarun "ScanJobData")
	(vl-vbarun "EditJobData")
)