(defun job-data-insert ( / old-osmode)
    (setq acadObj (vlax-get-acad-object))
    (setq doc (vla-get-ActiveDocument acadObj))
    (setq modelSpace (vla-get-ModelSpace doc))
    
    ;; Define the mtext object
    (setq p (vlax-3d-point 0 0 0)
          width 100
          text (strcat 
                    "job_number: L-0000\n"
                    "job_name: \n"
                    "job_site_address: \n"
                    "calculated_by_company: \n"
                    "sprinkler_pipe_type: PEX\n"
                    "sprinkler_fitting_type: Brass\n"
                    "supply_static_pressure: 0.0\n"
                    "supply_residual_pressure: 0.0\n"
                    "supply_available_flow: 0.0\n"
                    "supply_elevation: 0\n"
                    "supply_pipe_type: CPVC\n"
                    "supply_pipe_size: 1.0\n"
                    "supply_pipe_internal_diameter: 1.5\n"
                    "supply_pipe_length: 10.0\n"
                    "supply_name: SUPPLY\n"
                    "domestic_flow_added: 0.0\n"
                    "water_flow_switch_make_model: \n"
                    "water_flow_switch_pressure_loss: 0.0\n"
                    "supply_pipe_fittings_summary: 0.0\n"
                    "supply_pipe_fittings_equiv_length: 0.0\n"
                    "supply_pipe_add_pressure_loss: 0.0\n"
               )
    )
  
    ; Insert the MText object
    (setq MTextObj (vla-AddMText modelSpace p width text))
    (vla-put-height MTextObj 10.0)
    (vla-put-layer MTextObj "JobData")
)

(defun job-data-insert-old ( / old-osmode)
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
    ;(vl-vbarun "ScanJobData")
    ;(vl-vbarun "EditJobData")
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
    (job-data-attribute "JOB_NUMBER" "" 1)
    (job-data-attribute "JOB_NAME" "" 2)
    (job-data-attribute "JOB_SITE_ADDRESS" "" 3)
    (job-data-attribute "CALCULATED_BY_COMPANY" "" 4)
    (job-data-attribute "SPRINKLER_PIPE_TYPE" "" 5)
    (job-data-attribute "SPRINKLER_FITTING_TYPE" "" 6)
    (job-data-attribute "SUPPLY_STATIC_PRESSURE" "0" 7)
    (job-data-attribute "SUPPLY_RESIDUAL_PRESSURE" "0" 8)
    (job-data-attribute "SUPPLY_AVAILABLE_FLOW" "0" 9)
    (job-data-attribute "SUPPLY_ELEVATION" "0" 10)
    (job-data-attribute "SUPPLY_PIPE_TYPE" "" 11)
    (job-data-attribute "SUPPLY_PIPE_SIZE" "0" 12)
    (job-data-attribute "SUPPLY_PIPE_INTERNAL_DIAMETER" "0" 13)
    (job-data-attribute "SUPPLY_PIPE_LENGTH" "0" 14)
    (job-data-attribute "SUPPLY_NAME" "MTR" 15)
    (job-data-attribute "DOMESTIC_FLOW_ADDED" "0" 16)
    (job-data-attribute "WATER_FLOW_SWITCH_MAKE_MODEL" "" 17)
    (job-data-attribute "WATER_FLOW_SWITCH_PRESSURE_LOSS" "0" 18)
    (job-data-attribute "SUPPLY_PIPE_FITTINGS_SUMMARY" "" 19)
    (job-data-attribute "SUPPLY_PIPE_FITTINGS_EQUIV_LENGTH" "0" 20)
    (job-data-attribute "SUPPLY_PIPE_ADD_PRESSURE_LOSS" "0" 21)
    (entmake
        (list 
            (cons 0 "SEQEND") 
            ;(cons -2 e)
        )
    )
    ;(entupd e)
    (princ "\nBlock inserted\n")
    (princ)
    e
)

; Test
(defun test-job-data-attribute ()
    (princ "Should return nil\n")
    (job-data-attribute "TEST_TAG" "TEST_VALUE" 0)
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

(defun test-job-data-dialog ()
    (setq id (load_dialog "job_data.dcl"))
    (princ (strcat "\nJob Data CDL ID:" (itoa id)))
    (if (not (new_dialog "job_data" id))
        (princ "\nError loading job_data.dcl\n")
        (exit)
    )
    (unload_dialog id)
)

(vlax-ldata-put "job_data" "job_site_address" "123 Main St")

(defun job-data-dialog ( / id result key value block-name new-block-name )
    (setq id (load_dialog "job_data.dcl"))
    (new_dialog "job_data" id)
  
    ; Set tiles values from job_data values
    (foreach key job_data:keys 
        (progn
            (princ (strcat "\nLoading Job Data Key: " key))
            (setq value (get-job-data key))
            (if (not (null value))
                (progn
                  (princ (strcat "\n    Value: " value))
                  (set_tile key value)
                )
                (progn
                  (princ (strcat "\n    Value: EMPTY"))
                  (set_tile key "") ; set_tile does not accept nil as a value
                )
            )
        )
    )
    
    (setq result (start_dialog))
    (if (= result 1) ; 1 = User clicked 'OK'
        (set-job-data-attributes)
        (princ "\nCancelled. Job data not set.\n")
    )
    (unload_dialog id)
)

(defun job-data-dialog-old ( / id result key value block-name new-block-name )
    (setq id (load_dialog "job_data.dcl"))
    (new_dialog "job_data" id)
    
    (setq block-name (get-job-data-block-name))
    (if (null block-name)
        (insert-job-data-block '(0 0 0))
        (if (job-data-block-is-v1 block-name)
            (progn
                ; Convert v1 to v2
                (alert "An older version of job data was found, it will be converted to the new format.  If you do not want it converted, do not save the file.")
                (setq new-block-name (insert-job-data-block '(0 0 0)))
                (set-attribute new-block-name "JOB_NUMBER" (get-attribute-value block-name "LEAD_NUMBER"))
                (set-attribute new-block-name "JOB_NAME"(get-attribute-value block-name "JOB_NAME"))
                (set-attribute new-block-name "SUPPLY_STATIC_PRESSURE"(get-attribute-value block-name "STATIC_PRESSURE"))
                (set-attribute new-block-name "JOB_SITE_ADDRESS" (get-attribute-value block-name "SITE_LOCATION"))
                (set-attribute new-block-name "SUPPLY_RESIDUAL_PRESSURE" (get-attribute-value block-name "RESIDUAL_PRESSURE"))
                (set-attribute new-block-name "SUPPLY_AVAILABLE_FLOW" (get-attribute-value block-name "AVAILABLE_FLOW"))
                (set-attribute new-block-name "SUPPLY_ELEVATION" (get-attribute-value block-name "METER_ELEVATION"))
                (set-attribute new-block-name "SUPPLY_PIPE_LENGTH" (get-attribute-value block-name "METER_PIPE_LENGTH"))
                (set-attribute new-block-name "SUPPLY_PIPE_INTERNAL_DIAMETER" (get-attribute-value block-name "METER_PIPE_INTERNAL_DIAMETER"))
                (set-attribute new-block-name "CALCULATED_BY_COMPANY" (get-attribute-value block-name "CALCULATED_BY_COMPANY"))
                ; Delete old block
                (entdel block-name)
                (load-job-data-attributes new-block-name)
            )
            (load-job-data-attributes block-name)
        )
    )    
    
    ; Set tiles values from job_data values
    (foreach key job_data:keys 
        (progn
            (setq value (get-job-data key))
            (if (not (null value))
                (set_tile key value)
                (set_tile key "") ; set_tile does not accept nil as a value
            )
        )
    )
    
    (setq result (start_dialog))
    (if (= result 1) ; 1 = User clicked 'OK'
        (set-job-data-attributes)
        (princ "\nCancelled. Job data not set.\n")
    )
    (unload_dialog id)
)

(defun job-data-block-is-v1 ( block-name )
    (not (null (get-attribute block-name "LEAD_NUMBER")))
)

(defun set-job-data-attributes( / key value job-data-block-name )
    (setq job-data-block-name (get-job-data-block-name))
    (foreach key job_data:keys 
        (progn
            (setq value (get-job-data-var key))
            (set-attribute job-data-block-name (strcase key) value)
        )
    )
    (princ)
)

(defun load-job-data-attributes ( block-name / key value )
    (foreach key job_data:keys 
        (progn
            (setq value (get-attribute-value block-name (strcase key)))
            (set-job-data-var (strcase key T) value)
        )
    )
    (princ)
)

(defun get-job-data-block-name ( )
    (cdr 
        (assoc 
            -1 
            (car (get-blocks (list "JobData" "Job Data")))
        )
    )
)

(setq job_data:keys 
    ; Each of these keys is a key in the job_data.dcl dialog file.
    ; They must match exactly.
    (list 
        "job_number"
        "job_name"
        "job_site_address"
        "calculated_by_company"
        "sprinkler_pipe_type"
        "sprinkler_fitting_type"
        "supply_static_pressure"
        "supply_residual_pressure"
        "supply_available_flow"
        "supply_elevation"
        "supply_pipe_type"
        "supply_pipe_size"
        "supply_pipe_internal_diameter"
        "supply_pipe_length"
        "supply_name"
        "domestic_flow_added"
        "water_flow_switch_make_model"
        "water_flow_switch_pressure_loss"
        "supply_pipe_fittings_summary"
        "supply_pipe_fittings_equiv_length"
        "supply_pipe_add_pressure_loss"
    )
)

; Called only by job_data dialog.
(defun set-job-data ( key value )
  (vlax-ldata-put "job_data" key value)
)
(defun get-job-data ( key )
    (vlax-ldata-get "job_data" key)
)

(defun set-job-data-old ( key value / lav item-list n)
    (setq lav (get_attr key "list")) ; job_data dialog needs to be open for this line
    (if (> (strlen lav) 0) ; The "list" attribute has a list
        (progn 
            (setq item-list (string-split lav ","))
            (setq n (atoi value))
            (setq value (nth n item-list))
        )
    )
    (set-job-data-var key value)
)

(defun set-job-data-var-old ( key value )
    (if (null job_data) (setq job_data '()))
    (if (null (assoc key job_data))
        (setq job_data (append job_data (list (cons key value))))
        (setq job_data 
            (subst 
                (cons key value) 
                (assoc key job_data) 
                job_data
            )
        )
    )
)

; Call only when a dialog with this key is open
(defun get-job-data-old ( key / lav value index popup-list )
    (setq value (get-job-data-var key))
    (setq lav (get_attr key "list")) ; From job_data dialog
    (if (> (strlen lav) 0) ; The "list" attribute contains a list
        (progn 
            (setq popup-list (string-split lav ","))
            (setq index (index-of value popup-list))
            (itoa index)
        )
        value
    )
)

(defun get-job-data-var ( key )
    (cdr (assoc key job_data))
)

(defun string-split ( source target / len lst i )
    (setq len (1+ (strlen target)))
    (while (setq i (vl-string-search target source))
        (setq lst (cons (substr source 1 i) lst)
              source (substr source (+ i len))
        )
    )
    (reverse (cons source lst))
)