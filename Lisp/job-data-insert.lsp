(defun job-data-insert ( / old-osmode)
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
    (command "-INSERT" "JobData.dwg" pause 1.0 1.0 0)
)