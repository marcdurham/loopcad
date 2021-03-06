(if (not (data-request "DefaultHeadModel"))
    (data-submit "DefaultHeadModel" "RFC67")
)
(if (not (data-request "DefaultHeadSlope"))
    (data-submit "DefaultHeadSlope" "1")
)
(if (not (data-request "DefaultHeadTemperature"))
    (data-submit "DefaultHeadTemperature" "166")
)

; TODO: Maybe have two parameters: model-code and filename, use a separate function to determine them.
(defun head-insert (coverage / model-code pt)
       ;(model coverage slope temperature / model-code pt)
  (setq old-osmode (getvar "OSMODE"))
  (setq temperror *error*)
  (defun *error* (message)
      (princ)
      (princ message)
      (princ)
    (setvar "OSMODE" old-osmode)
    (command-s "-LAYER" "OFF" "HeadCoverage" "")
    (setvar "LWDISPLAY" 1)
    (setq *error* temperror)
  )
  (setvar "INSUNITS" 2) ; 0 = not set, 1 = inches, 2 = feet
                        ; This line prevents inserted block refs from having a
                        ; different scale, being 12 times bigger than they should be.
  (setvar "OSMODE" osmode-snap-ins-pts)
  ;(command "-LAYER" "NEW" "Heads" "")
  ;(command "-LAYER" "NEW" "HeadCoverage" "")
  ;(command "-LAYER" "COLOR" "Red" "Heads" "")
  ;(command "-LAYER" "COLOR" "Yellow" "HeadCoverage" "")
  (command-s "-LAYER" "ON" "HeadCoverage" "")
  (setvar "LWDISPLAY" 0)
  (command-s "-LAYER" "SET" "Heads" "")
  (setq model-code "HEAD-X")
  ;;(while T
;
;   This section is for heads that you already know the coverage
;   Now the default is that you don't know. 
;   See the head-insert.lsp file and the head-insert function. 
;
;   (setq model-code (model-code-from model coverage slope temperature))
;    (prompt (strcat "\nInserting Head Model Code: " model-code "\n"))
;   (prompt "\nPress Esc to quit inserting heads.\n")
;    (command "-INSERT" (strcat "Head" coverage ".dwg") pause 1.0 1.0 0 model-code)
;
    (command-s "-INSERT" (strcat "Head" coverage) pause 1.0 1.0 0 "")
  (prompt "Prese the [ENTER] key to repeat")
  ;;)
    ; (setq pt (cdr (assoc 10 (entget (entlast)))))
    ; (if (null global:head-coverage)
    ;     (setq global:head-coverage "16")
    ; )
    ; (initget "12 14 16 18 20")
    ; (if (setq tmp (getkword (strcat "\nHead Coverage [12/14/16/18/20] <" global:head-coverage ">: ")))
    ;     (progn
    ;       (princ (strcat "\nHead Coverage Selected: " tmp))
    ;       (setq global:head-coverage tmp)
    ;     )
    ; )
    ; (entdel (entlast))
    ; (princ "\nLast entity (temp) deleted\n")
    ; (setq model-code (strcat model "-" (itoa global:head-coverage)))
    ; (prompt (strcat "\nInserting Head Model Code: " model-code "\n"))
    ; (prompt "\nPress Esc to quit inserting heads.\n")
    ; ;;;(command "-INSERT" (strcat "Head" (itoa global:head-coverage) ".dwg") pt 1.0 1.0 0 model-code)
    
    ; (setq acadObj (vlax-get-acad-object))
    ; (setq doc (vla-get-ActiveDocument acadObj))
    
    ; ; Insert the block
    ; (setq insertionPoint (vlax-3d-point pt))
    ; (setq modelSpace (vla-get-ModelSpace doc))
    ; (setq headBlockName (strcat "Head" global:head-coverage))
    ; (princ (strcat "\n Head Block Name: " headBlockName))
    ; (setq block (vla-InsertBlock modelSpace insertionPoint headBlockName 1 1 1 0 "X-2"))
      
    ; ; get the block attributes
    ; (setq attributes (vlax-safearray->list (vlax-variant-value (vla-getAttributes block))))
    
    ; ; Set attribute values by the attribute position
    ; (vla-put-TextString (nth 0 attributes) model-code)
    
 ; )
  
  (setvar "OSMODE" old-osmode)
    (command "-LAYER" "OFF" "HeadCoverage" "")
    (setvar "LWDISPLAY" 1)
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
  (setvar "OSMODE" osmode-snap-ins-pts)
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
        0 ; TODO: pause here for Rotation 
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
;        (setq global:head-coverage "16")
;    )
;    (initget "12 14 16 18 20")
;    (if (setq tmp (getkword (strcat "\nHead Coverage [12/14/16/18/20] <" global:head-coverage ">: ")))
;        (setq global:head-coverage tmp)
;    )
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
    ;    (setq global:head-spray-direction tmp)
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
       ; (data-request "DefaultHeadModel")
        coverage
       ; (data-request "DefaultHeadSlope")
       ; (data-request "DefaultHeadTemperature")
    )
)