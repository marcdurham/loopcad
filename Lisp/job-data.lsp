(defun test-job-data-dialog ()
    (setq id (load_dialog "job_data.dcl"))
    (princ (strcat "\nJob Data CDL ID:" (itoa id)))
    (if (not (new_dialog "job_data" id))
        (princ "\nError loading job_data.dcl\n")
        (exit)
    )
    (unload_dialog id)
)

(defun job-data-dialog ( / id result key value block-name new-block-name )
    (setq id (load_dialog "job_data.dcl"))
    (new_dialog "job_data" id)
  
    ; Set tiles values from job_data values saved to the DWG file
    (foreach key job_data:keys 
        (progn
            (setq value (load-job-data key ""))
            (if (not (null value))
                (set_tile key value)
                (set_tile key "") ; set_tile does not accept nil as a value
            )
        )
    )
    
    (setq result (start_dialog))
    (if (= result 1) ; 1 = User clicked 'OK'
        (foreach key job_data:keys 
            (progn
                (save-job-data key (get-dict-data "job_data_temp" key))
            )
        )
        (princ "\nCancelled. Job data not set.\n")
    )
    (unload_dialog id)
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
        "head_model_default"
        "head_coverage_default"
    )
)

; JOB-DATA functions, called only by job_data dialog.
(defun set-job-data ( key value )
    (set-dict-data "job_data_temp" key value)
)

; Data saved to the file
; Both dictionaries are actually saved, but job_data_temp is ignored
; There's no reason to delete it
(defun load-job-data ( key default / a b)
    (setq a (get-dict-data "job_data" key))
    (setq b (getcfg (strcat "AppData/LoopCAD/" key)))
    (if a a (if b b default))
)

(defun save-job-data ( key value )
    (set-dict-data "job_data" key value)
    t
)


; Generic DICT-DATA functions
; ****************************************************************
; Get a text value from the named dictionary
(defun get-dict-data (dict-name key)
    (cdr (assoc 1 (dictsearch (get-data-dict dict-name) key)))
)

; Set a text value in the named dictionary
(defun set-dict-data (dict-name key value / data-dict xrecord) 
    (setq data-dict (get-data-dict dict-name))
    ; If the XRecord does exist
    (if (setq xrecord (dictsearch data-dict key))
        (progn ; It does exist
            (dictremove (get-data-dict dict-name) key)
            (setq xrecord (entmakex 
                (list 
                  '(0 . "XRECORD")
                  '(100 . "AcDbXrecord")
                  (cons 1 value) ; Text value
                ))
            )
            ; If creation succeeded then add it to the dictionary
            (if xrecord (setq xrecord (dictadd data-dict key xrecord)))
        )
        (progn ; It does not exist
            (setq xrecord (entmakex 
                (list 
                  '(0 . "XRECORD")
                  '(100 . "AcDbXrecord")
                  (cons 1 value) ; Text value
                ))
            )
            ; If creation succeeded then add it to the dictionary
            (if xrecord (setq xrecord (dictadd data-dict key xrecord)))
        )
    )
)

; Get or create and get the named dictionary
(defun get-data-dict (dict-name / data-dict)
    ; If "data-dict" is already present in the main dictionary
    (if (not (setq data-dict (dictsearch (namedobjdict) dict-name)))
        ; Create the "data-dict" dictionary set the main dictionary as owner
        (progn
            (setq data-dict (entmakex '((0 . "DICTIONARY")(100 . "AcDbDictionary"))))
            ; if succesfully created, add it to the main dictionary
            (if data-dict (setq data-dict (dictadd (namedobjdict) dict-name data-dict)))
        )
        ; If "data-dict" exists then return its entity name
        (setq data-dict (cdr (assoc -1 data-dict)))
    )
)