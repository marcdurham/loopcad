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
            (setq value (load-job-data key))
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
                (save-job-data key (vlax-ldata-get "job_data_temp" key))
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
    )
)

; Called only by job_data dialog.
; Live data in the dialog
(defun set-job-data ( key value )
  (vlax-ldata-put "job_data_temp" key value)
)
(defun load-job-data ( key )
    (vlax-ldata-get "job_data_temp" key)
)

; Data saved to the file
; Both dictionaries are actually saved, but job_data_temp is ignored
; There's no reason to delete it
(defun load-job-data ( key )
    (vlax-ldata-get "job_data" key)
)
(defun save-job-data ( key value )
  (vlax-ldata-put "job_data" key value)
)
