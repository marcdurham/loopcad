; DXF Code Reference Document
; https://images.autodesk.com/adsk/files/autocad_2012_pdf_dxf-reference_enu.pdf
; Head Label Properties
(setq head-label:tag-string "HEADNUMBER")
(setq head-label:prompt "Head number label")
(setq head-label:label-color color-blue)
(setq head-label:layer "HeadLabels")
(setq head-label:x-offset 9.132)
(setq head-label:y-offset 0.0)

; Tee Label Properties
(setq tee-label:tag-string "TEENUMBER")
(setq tee-label:prompt "Tee number label")
(setq tee-label:label-color color-green)
(setq tee-label:layer "TeeLabels")
(setq tee-label:x-offset 4.0)
(setq tee-label:y-offset 4.0)

; Riser Label Properties
(setq riser-label:tag-string "RISERNUMBER")
(setq riser-label:prompt "Riser number label")
(setq riser-label:layer "RiserLabels")
(setq riser-label:label-color color-green)

; Head Model Properties
; Insert Point: 9.132, 8.395 copied from old block so it looks the same
(setq head-block:model-x-offset 9.132)
(setq head-block:model-y-offset 8.395)

(defun define-labelsx ()
    (princ "\nDoing nothing\n")
)
(defun define-labels ()
    ; These two block definitions are not used by any functions but they are defined so that
    ; a user can use the "INSERT" command to insert them manually if they want.
    (define-label-block 
        "HeadLabel" 
        head-label:tag-string 
        head-label:prompt 
        "H.0" 
        head-label:label-color 
        head-label:layer
        head-block:model-x-offset    ; Label X Offset
        0                              ; Label Y Offset
    )
    (define-label-block 
        "TeeLabel" 
        tee-label:tag-string
        tee-label:prompt 
        "T.0" 
        tee-label:label-color
        tee-label:layer
        head-block:model-x-offset    ; Label X Offset
        0                              ; Label Y Offset
    )
    (define-label-block 
        "RiserLabel" 
        riser-label:tag-string
        riser-label:prompt 
        "R.0.X" 
        riser-label:label-color
        riser-label:layer
        head-block:model-x-offset    ; Label X Offset
        0                              ; Label Y Offset
    )
    (define-head-coverage 12)
    (define-head-coverage 14)
    (define-head-coverage 16)
    (define-head-coverage 18)
    (define-head-coverage 20)
    (define-sw-head-coverage 12 "U") ; Sprays Up
    (define-sw-head-coverage 14 "U")
    (define-sw-head-coverage 16 "U")
    (define-sw-head-coverage 18 "U")
    (define-sw-head-coverage 20 "U")
    (define-sw-head-coverage 12 "D") ; Sprays Down
    (define-sw-head-coverage 14 "D")
    (define-sw-head-coverage 16 "D")
    (define-sw-head-coverage 18 "D")
    (define-sw-head-coverage 20 "D")
    (define-sw-head-coverage 12 "L") ; Sprays Left
    (define-sw-head-coverage 14 "L")
    (define-sw-head-coverage 16 "L")
    (define-sw-head-coverage 18 "L")
    (define-sw-head-coverage 20 "L")
    (define-sw-head-coverage 12 "R") ; Sprays Right
    (define-sw-head-coverage 14 "R")
    (define-sw-head-coverage 16 "R")
    (define-sw-head-coverage 18 "R")
    (define-sw-head-coverage 20 "R")    
    (define-floor-tag)
    (define-riser)
    (princ "\nLabels defined.\n")
    (define-job-data)
    (princ "\nJob data block defined.\n")
    (princ)
)

(defun define-head-coverage (coverage)
    (define-head-block coverage "Head" "MODEL" "Head model" "MODEL" color-red "Heads")
)

(defun define-sw-head-coverage (coverage direction)
    (define-sw-head-block direction coverage "Head" "MODEL" "Head model" "MODEL" color-red "Heads")
)

(defun define-label-block (block-name tag-string prompt default label-color layer label-x-offset label-y-offset)
    (entmake 
        (list
            (cons 0 "BLOCK")
            (cons 2 block-name) ; Block name
        )
    )
    (entmake 
        (list
            (cons 0 "ATTDEF")
            (cons 10 (list label-x-offset label-y-offset 0))
            (cons 1 default)      ; Text value
            (cons 2 tag-string)   ; Tag string
            (cons 3 prompt)       ; Prompt string
            (cons 40 5.0)         ; Text height
            (cons 7 "ARIAL")      ; Text style
            (cons 62 color-yellow) ; Color
            (cons 8 layer)        ; Layer
        )
    )
    (entmake 
        (list
            (cons 0 "ENDBLK")
        )
    )
)

; Side-wall Head Block
(defun define-sw-head-block (direction coverage block-name tag-string prompt default label-color layer / span halfway quarter coverage-text t-left-x t-left-y t-right-x t-right-y t-top-x t-top-y t-vertical c-ll-x c-ll-y c-lr-x c-lr-y c-ur-x c-ur-y c-ul-x c-ul-y text-center-x text-center-y model-label-x model-label-y)
    (entmake 
        (list
            (cons 0 "BLOCK")
            (cons 2 (strcat "Sw" block-name (itoa coverage) direction)) ; Block name
        )
    )
    
    ; Head Model Number
    (setq model-label-x head-block:model-x-offset)
    (setq model-label-y head-block:model-y-offset)
    ;(cond ((= direction "D")
    ;        (setq model-label-y (- 0 model-label-y)))
    ;    ((= direction "L")
    ;        (setq model-label-x (- 0 model-label-x)))
    ;)
    (entmake 
        (list
            (cons 0 "ATTDEF")
            (cons 10 
                (list 
                    model-label-x
                    model-label-y
                    0.0
                )
            )
            (cons 1 default)      ; Text value
            (cons 2 tag-string)   ; Tag string
            (cons 3 prompt)       ; Prompt string
            (cons 40 5.0)         ; Text height
            (cons 7 "ARIAL")      ; Text style
            (cons 62 label-color) ; Color
            (cons 8 layer)        ; Layer
        )
    )
    
    ; Head
    ; Triangle
    ; Default direction "U" (not in the cond below)
    (setq t-left-x -6)
    (setq t-right-x 6)            
    (setq t-left-y 12)
    (setq t-right-y 12)        
    (cond ((= direction "D")
            (setq t-left-x 6)
            (setq t-right-x -6)
            (setq t-left-y -12)
            (setq t-right-y -12)                    
        )
        ((= direction "L")
            (setq t-left-x -12)
            (setq t-right-x -12)                    
            (setq t-left-y 6)
            (setq t-right-y -6)        
        )
        ((= direction "R")
            (setq t-left-x 12)
            (setq t-right-x 12)
            (setq t-left-y 6)
            (setq t-right-y -6)        
        )
    )
    (entmake
        (list
            (cons 0 "POLYLINE")
            (cons 10 (list 0 0 0))  ; Point is always zero
            (cons 70 1)             ; 1 = Closed Polyline
            (cons 62 color-red)  ; Color
            (cons 8 layer) ; Layer
        )
    )
    (entmake
        (list
            (cons 0 "VERTEX")
            (cons 10 (list t-left-x t-left-y 0)) ; Left
        )
    )
    (entmake
        (list
            (cons 0 "VERTEX")
            (cons 10 (list t-right-x t-right-y 0)) ; Right
        )
    )
    (entmake
        (list
            (cons 0 "VERTEX")
            (cons 10 (list 0 0 0))    ; Top
        )
    )
    (entmake
        (list
            (cons 0 "SEQEND")
        )
    )
    
    ; Head Coverage Box
    (setq span (feet->inches coverage))
    (setq -span (- 0 span))
    (setq halfway (/ span 2))
    (setq -halfway (- 0 halfway))
    (setq quarter (/ span 4))

    (setq c-ll-x -halfway)
    (setq c-ll-y 0)
    (setq c-lr-x halfway)
    (setq c-lr-y 0)
    (setq c-ur-x halfway)
    (setq c-ur-y span)
    (setq c-ul-x -halfway)
    (setq c-ul-y span)
    (setq text-center-x 0)
    (setq text-center-y halfway)
    
    (cond ((= direction "D")
            (setq c-ll-x -halfway)
            (setq c-ll-y -span)
            (setq c-lr-x halfway)
            (setq c-lr-y -span)
            (setq c-ur-x halfway)
            (setq c-ur-y 0)
            (setq c-ul-x -halfway)
            (setq c-ul-y 0)
            (setq text-center-x 0)
            (setq text-center-y -halfway))
        ((= direction "L")
            (setq c-ll-x -span)
            (setq c-ll-y -halfway)
            (setq c-lr-x 0)
            (setq c-lr-y -halfway)
            (setq c-ur-x 0)
            (setq c-ur-y halfway)
            (setq c-ul-x -span)
            (setq c-ul-y halfway)
            (setq text-center-x -halfway)
            (setq text-center-y 0))
        ((= direction "R")
            (setq c-ll-x 0)
            (setq c-ll-y -halfway)
            (setq c-lr-x span)
            (setq c-lr-y -halfway)
            (setq c-ur-x span)
            (setq c-ur-y halfway)
            (setq c-ul-x 0)
            (setq c-ul-y halfway)
            (setq text-center-x halfway)
            (setq text-center-y 0))
    )
    (entmake
        (list
            (cons 0 "POLYLINE")
            (cons 10 (list 0 0 0))  ; Point is always zero
            (cons 70 1)             ; 1 = Closed Polyline
            (cons 62 color-yellow)  ; Color
            (cons 8 "HeadCoverage") ; Layer
        )
    )
    (entmake
        (list
            (cons 0 "VERTEX")
            (cons 10 (list c-ll-x c-ll-y 0)) ; Lower Left
        )
    )
    (entmake
        (list
            (cons 0 "VERTEX")
            (cons 10 (list c-lr-x c-lr-y 0)) ; Lower Right
        )
    )
    (entmake
        (list
            (cons 0 "VERTEX")
            (cons 10 (list c-ur-x c-ur-y 0)) ; Upper Right
        )
    )
    (entmake
        (list
            (cons 0 "VERTEX")
            (cons 10 (list c-ul-x c-ul-y 0)) ; Upper Left
        )
    )
    (entmake
        (list
            (cons 0 "SEQEND")
        )
    )
    
    ; Coverage Text 
    ; Example: 12' x 12'
    (setq coverage-text 
        (strcat (itoa coverage) "'  X  " (itoa coverage) "'")
    )
    (entmake
        (list
            (cons 0 "TEXT")
            (cons 10 (list span span 0)) ; Upper left corner
            (cons 11 (list text-center-x text-center-y 0)) ; Second alignment point, center of text
            (cons 40 16.0)         ; Text height
            (cons 1 coverage-text) ; Text value
            (cons 72 1) ; Horizontal text justification: 1 = Center, 4 = Middle
            (cons 73 2) ; Vertical text justification: 2 = Middle
            (cons 62 color-yellow)  ; Color
            (cons 8 "HeadCoverage") ; Layer
        )
    )
    (entmake 
        (list
            (cons 0 "ENDBLK")
        )
    )
)

; Head Block (Normal)
(defun define-head-block (coverage block-name tag-string prompt default label-color layer / span halfway quarter coverage-text)
    (entmake 
        (list
            (cons 0 "BLOCK")
            (cons 2 (strcat block-name (itoa coverage))) ; Block name
        )
    )
    
    ; Head Model Number
    (entmake 
        (list
            (cons 0 "ATTDEF")
            (cons 10 
                (list 
                    head-block:model-x-offset 
                    head-block:model-y-offset 
                    0.0
                )
            )
            (cons 1 default)      ; Text value
            (cons 2 tag-string)   ; Tag string
            (cons 3 prompt)       ; Prompt string
            (cons 40 5.0)         ; Text height
            (cons 7 "ARIAL")      ; Text style
            (cons 62 label-color) ; Color
            (cons 8 layer)        ; Layer
        )
    )
    
    ; Head
    ; Inner Circle
    (entmake
        (list
            (cons 0 "CIRCLE")     
            (cons 10 (list 0 0 0)) ; Center Point
            ; Radius: 2.278 copeid from old block so it looks the same
            (cons 40 2.278)        ; Radius
            (cons 62 color-red)    ; Color
            (cons 8 layer)         ; Layer
        )
    )
    ; Outer Circle
    (entmake
        (list
            (cons 0 "CIRCLE")      
            (cons 10 (list 0 0 0)) ; Center Point
            ; Radius: 6.653 copeid from old block so it looks the same
            (cons 40 6.653)        ; Radius
            (cons 62 color-red)    ; Color
            (cons 8 layer)         ; Layer
        )
    )
    
    ; Head Coverage Box
    (setq span (feet->inches coverage))
    (setq -span (- 0 span))
    (setq halfway (/ span 2))
    (setq -halfway (- 0 halfway))
    (setq quarter (/ span 4))
    (entmake
        (list
            (cons 0 "POLYLINE")
            (cons 10 (list 0 0 0))  ; Point is always zero
            (cons 70 1)             ; 1 = Closed Polyline
            (cons 62 color-yellow)  ; Color
            (cons 8 "HeadCoverage") ; Layer
        )
    )
    (entmake
        (list
            (cons 0 "VERTEX")
            (cons 10 (list -halfway -halfway 0)) ; Lower Left
        )
    )
    (entmake
        (list
            (cons 0 "VERTEX")
            (cons 10 (list halfway -halfway 0))    ; Lower Right
        )
    )
    (entmake
        (list
            (cons 0 "VERTEX")
            (cons 10 (list halfway halfway 0))    ; Upper Right
        )
    )
    (entmake
        (list
            (cons 0 "VERTEX")
            (cons 10 (list -halfway halfway 0))    ; Upper Left
        )
    )
    (entmake
        (list
            (cons 0 "SEQEND")
        )
    )
    
    ; Coverage Text 
    ; Example: 12' x 12'
    (setq coverage-text 
        (strcat (itoa coverage) "'  X  " (itoa coverage) "'")
    )
    (entmake
        (list
            (cons 0 "TEXT")
            (cons 10 (list span span 0)) ; Upper left corner
            (cons 11 (list 0 quarter 0)) ; Second alignment point, center of text
            (cons 40 16.0)         ; Text height
            (cons 1 coverage-text) ; Text value
            (cons 72 1) ; Horizontal text justification: 1 = Center, 4 = Middle
            (cons 73 2) ; Vertical text justification: 2 = Middle
            (cons 62 color-yellow)  ; Color
            (cons 8 "HeadCoverage") ; Layer
        )
    )
    (entmake 
        (list
            (cons 0 "ENDBLK")
        )
    )
)

; Floor Tag
(defun define-floor-tag ( / label-color layer )
   (defun *error* (message)
      (princ)
      (princ message)
      (princ)
    )
    (setq layer "FloorTags")
    (entmake 
        (list
            '(0 . "BLOCK")
            '(2 . "FloorTag") ; Block name
            (cons 8 layer)      ; Layer (recommended)
            '(10 0.0 0.0 0.0)         ; required
            '(70 . 2)                 ; required [NOTE 0 if no attributes]
            '(100 . "AcDbEntity")     ; recommended
            '(100 . "AcDbBlockBegin") ; recommended
        )
    )

    (entmake 
        '(
            (0 . "ATTDEF")
            (1 . "Main Floor")   ; Default value
            (2 . "NAME")         ; Tag name
            (3 . "Enter floor name")
            (8 . "FloorTag")     ; Layer
            (10 10.0 -10.0 0.0)  ; Coordinates
            (40 . 9.0)   ; Text Size (KEEP)
            (62 . 4)     ; Color (4 = Cyan)
            (70 . 0)     ; Attribute Flags (KEEP)
        )
    )
    
    (entmake 
        '(
            (0 . "ATTDEF")
            (1 . "100")         ; Default value
            (2 . "ELEVATION")   ; Tag name
            (3 . "Enter elevation in feet")
            (8 . "FloorTag")    ; Layer
            (10 10.0 -20.0 0.0) ; Coordinates
            (40 . 9.0)    ; Text Size (KEEP)
            (62 . 4)      ; Color (4 = Cyan)
            (70 . 0)      ; Attribute Flags (KEEP)
        )
    )
    
    ; Outer Circle
    (entmake
        (list
            '(0 . "CIRCLE")      
            '(10 0 0 0)       ; Center Point
            ; Radius: 7.71 copied from old block so it looks the same
            '(40 . 7.71)      ; Radius
            '(62 . 4)         ; Color (4 = Cyan)
            (cons 8 layer)    ; Layer
        )
    )
    
    ; Vertical Line
    (entmake
        (list
            '(0 . "POLYLINE")
            '(10 0 0 0)              ; Point is always zero
            '(70 . 1)                ; 1 = Closed Polyline
            (cons 62 color-cyan)     ; Color
            (cons 8 layer)           ; Layer
        )
    )
    (entmake
        '(
            (0 . "VERTEX")
            (10 7.71 0 0) ; Top
        )
    )
    (entmake
        '(
            (0 . "VERTEX")
            (10 -7.71 0 0)    ; Bottom
        )
    )
    (entmake
        '(
            (0 . "SEQEND")
        )
    )

    ; Horizontal Line
    (entmake
        '(
            (0 . "POLYLINE")
            (10 0 0 0)           ; Point is always zero
            (70 . 1)             ; 1 = Closed Polyline
            (62 . 4)             ; 4: color-cyan; Color
            (8 . "FloorTags")    ; Layer
        )
    )
    (entmake
        '(
            (0 . "VERTEX")
            (10 0 -7.71 0) ; Left
        )
    )
    (entmake
        '(
            (0 . "VERTEX")
            (10 0 7.71 0)    ; Right
        )
    )
    (entmake
        '(
            (0 . "SEQEND")
        )
    )

    (entmake 
        '(
            (0 . "ENDBLK")
            (100 . "AcDbBlockEnd") ; recommended
            (8 . "0")              ; recommended
        )
    )
)

; Riser
(defun define-riser ( / label-color layer )
    (setq layer "Risers")
    (entmake 
        (list
            (cons 0 "BLOCK")
            (cons 2 "Riser") ; Block name
        )
    )
    
    ; Outer Circle
    (entmake
        (list
            (cons 0 "CIRCLE")      
            (cons 10 (list 0 0 0)) ; Center Point
            ; Radius: 5.0 copied from old block so it looks the same
            (cons 40 4.5)        ; Radius
            (cons 62 color-cyan)  ; Color
            (cons 8 layer)           ; Layer
        )
    )
    
    (entmake 
        (list
            (cons 0 "ENDBLK")
        )
    )
    (princ)
)

; New Job Data
; 2021-03-05 Trying a different approach
(defun define-job-data-old ( / label-color layer )
    (setq layer "JobData")
  
    (setq acadObj (vlax-get-acad-object))
    (setq doc (vla-get-ActiveDocument acadObj))
    
    ; Insert the block
    (setq insertionPoint (vlax-3d-point 0 0 0))
    (setq modelSpace (vla-get-ModelSpace doc))
  
  
  ;; Create the block
    (setq blockObj (vla-Add (vla-get-Blocks doc) insertionPnt "JobData"))
    
     ;; Add a circle to the block
    (setq center (vlax-3d-point 0 0 0)
          radius 1)
    (setq circleObj (vla-AddCircle blockObj center radius))
  
  
   ;; Define the attribute definition
    (setq insertionPoint (vlax-3d-point 5 5 0) 
          attHeight 1
          attMode acAttributeModeVerify
          attPrompt "New Prompt"
          attTag "NEW_TAG"
          attValue "New Value")
    
    ;; Create the attribute definition object in model space    
    ;(setq attributeObj 
    ;    (vla-AddAttribute blockObj attHeight attMode attPrompt insertionPoint attTag attValue))
  
  (setq blockRefObj (vla-InsertBlock modelSpace (vlax-3d-point 100 100 0) "JobData" 1 1 1 0))
)

; Job Data
(defun define-job-data-old ( / label-color layer )
    (setq layer "JobData")
  
    (setq acadObj (vlax-get-acad-object))
    (setq doc (vla-get-ActiveDocument acadObj))
    
    ; Insert the block
    (setq insertionPoint (vlax-3d-point p))
    (setq modelSpace (vla-get-ModelSpace doc))
  
  
  ;; Create the block
    (setq blockObj (vla-Add (vla-get-Blocks doc) insertionPnt "JobData"))
    
    ;; Add a circle to the block
   ;; Define the attribute definition
    (setq insertionPoint (vlax-3d-point 5 5 0) 
          attHeight 1
          attMode acAttributeModeVerify
          attPrompt "New Prompt"
          attTag "NEW_TAG"
          attValue "New Value")
    
    ;; Create the attribute definition object in model space    
    ;;;;;;(setq attributeObj (vla-AddAttribute blockObj attHeight attMode attPrompt insertionPoint attTag 
   
  
    ;;(setq block (vla-InsertBlock modelSpace insertionPoint "FloorTag" 1 1 1 0))
  
    ; get the block attributes
    (setq attributes (vlax-safearray->list (vlax-variant-value (vla-getAttributes block))))
    
    ;(get_attr key attribute)
    (foreach attribute attributes
      (progn  
          (setq tag (val-get-TagString attribute))
          (princ (strcat "\nTag: " tag))
        (vla-addattribute block)
      )
        
    )
    ; Set attribute values by the attribute position
    ;(vla-put-TextString (nth 0 attributes) floor-name)
    ;(vla-put-TextString (nth 1 attributes) elevation)
  
    (job-data-attdef "JOB_NUMBER" "" "" 1)
    (job-data-attdef "JOB_NAME" "" "" 2)
    (job-data-attdef "JOB_SITE_ADDRESS" "" "" 3)
    (job-data-attdef "CALCULATED_BY_COMPANY" "" "" 4)
    (job-data-attdef "SUPPLY_STATIC_PRESSURE" "0" "" 5)
    (job-data-attdef "SUPPLY_RESIDUAL_PRESSURE" "0" "" 6)
    (job-data-attdef "SUPPLY_AVAILABLE_FLOW" "0" "" 7)
    (job-data-attdef "SUPPLY_ELEVATION" "0" "" 8)
    (job-data-attdef "SUPPLY_PIPE_LENGTH" "0" "" 9)
    (job-data-attdef "SUPPLY_PIPE_INTERNAL_DIAMETER" "0" "" 10)
    (job-data-attdef "SPRINKLER_PIPE_TYPE" "" "" 11)
    (job-data-attdef "SPRINKLER_FITTING_TYPE" "" "" 12)
    (job-data-attdef "SUPPLY_PIPE_TYPE" "" "" 13)
    (job-data-attdef "SUPPLY_PIPE_SIZE" "0" "" 14)
    (job-data-attdef "SUPPLY_NAME" "MTR" "" 15)
    (job-data-attdef "DOMESTIC_FLOW_ADDED" "0" "" 16)
    (job-data-attdef "WATER_FLOW_SWITCH_MAKE_MODEL" "" "" 17)
    (job-data-attdef "SUPPLY_PIPE_FITTINGS_SUMMARY" "" "" 18)
    (job-data-attdef "SUPPLY_PIPE_FITTINGS_EQUIV_LENGTH" "0" "" 19)
    (job-data-attdef "SUPPLY_PIPE_ADD_PRESSURE_LOSS" "0" "" 20)
    (job-data-attdef "WATER_FLOW_SWITCH_PRESSURE_LOSS" "0" "" 21)
    
)

; Convert feet to inches
(defun feet->inches (feet)
    (* feet 12)
)

(defun alt_def ()
  
  ;BLOCK Header definition:
  ;Code 70 shows that attributes follow. Code 10 contains the
  ;insertion point.
  
  (entmake '((0 . "BLOCK")(2 . "ALT_ID")(70 . 2)(10 0.0 0.0 0.0)))
  
  ;Text ATTRIBUTE definition:
  
  (entmake 
    '((0 . "ATTDEF")(8 . "0")(10 0.0 0.0 0.0)(1 . "I")(2 . "NUM_ALT")
    (3 . "Alturas")(40 . 2.0)(41 . 1.0)(50 . 0.0)(70 . 0)(71 . 0)(72 . 4)(73 . 2))
  )
  
  ;BLOCK's ending definition:
  
  (entmake '((0 . "ENDBLK")))
)


; New Job Data
; 2021-03-05 Trying a different approach
(defun define-job-data-old ( / label-color layer )
    (setq layer "JobData")
  
    (setq acadObj (vlax-get-acad-object))
    (setq doc (vla-get-ActiveDocument acadObj))
    
    ; Insert the block
    (setq insertionPoint (vlax-3d-point 0 0 0))
    (setq modelSpace (vla-get-ModelSpace doc))
  
  
  ;; Create the block
    (setq blockObj (vla-Add (vla-get-Blocks doc) insertionPnt "JobData"))
    
     ;; Add a circle to the block
    (setq center (vlax-3d-point 0 0 0)
          radius 1)
    (setq circleObj (vla-AddCircle blockObj center radius))
  
  
   ;; Define the attribute definition
    (setq insertionPoint (vlax-3d-point 5 5 0) 
          attHeight 1
          attMode acAttributeModeVerify
          attPrompt "New Prompt"
          attTag "NEW_TAG"
          attValue "New Value")
    
    ;; Create the attribute definition object in model space    
    ;(setq attributeObj 
    ;    (vla-AddAttribute blockObj attHeight attMode attPrompt insertionPoint attTag attValue))
  
  (setq blockRefObj (vla-InsertBlock modelSpace (vlax-3d-point 100 100 0) "JobData" 1 1 1 0))
)
