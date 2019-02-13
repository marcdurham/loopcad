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
(setq tee-label:label-color color-blue)
(setq tee-label:layer "TeeLabels")
(setq tee-label:x-offset 4.0)
(setq tee-label:y-offset 4.0)

; Riser Label Properties
(setq riser-label:tag-string "RISERNUMBER")
(setq riser-label:layer "RiserLabels")

; Head Model Properties
; Insert Point: 9.132, 8.395 copied from old block so it looks the same
(setq head-block:model-x-offset 9.132)
(setq head-block:model-y-offset 8.395)

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
		0 							 ; Label Y Offset
	)
	(define-label-block 
		"TeeLabel" 
		tee-label:tag-string
		tee-label:prompt 
		"T.0" 
		tee-label:label-color
		tee-label:layer
		head-block:model-x-offset    ; Label X Offset
		0 							 ; Label Y Offset
	)
	(define-head-coverage 12)
	(define-head-coverage 14)
	(define-head-coverage 16)
	(define-head-coverage 18)
	(define-head-coverage 20)
	(define-floor-tag)
	(princ "\nLabels defined.\n")
	(princ)
)

(defun define-head-coverage (coverage)
    (define-head-block coverage "Head" "MODEL" "Head model" "MODEL" color-red "Heads")
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
			(cons 10 (list halfway -halfway 0))	; Lower Right
		)
	)
	(entmake
		(list
			(cons 0 "VERTEX")
			(cons 10 (list halfway halfway 0))	; Upper Right
		)
	)
	(entmake
		(list
			(cons 0 "VERTEX")
			(cons 10 (list -halfway halfway 0))	; Upper Left
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
	(setq layer "FloorTags")
	(entmake 
		(list
			(cons 0 "BLOCK")
			(cons 2 "FloorTag") ; Block name
		)
	)
	
	; Name
	(entmake 
		(list
			(cons 0 "ATTDEF")
			(cons 10 (list 10.0 -10.0 0.0))
			(cons 1 "")       ; Text value
			(cons 2 "NAME")   ; Tag string
			(cons 3 "Enter floor name") ; Prompt string
			(cons 40 5.0)         ; Text height
			(cons 7 "ARIAL")      ; Text style
			(cons 62 color-cyan)  ; Color
			(cons 8 layer) ; Layer
		)
	)
	
	; Elevation
	(entmake 
		(list
			(cons 0 "ATTDEF")
			(cons 10 (list 10.0 -20.0 0.0))	
			(cons 1 "0")           ; Text value
			(cons 2 "ELEVATION")   ; Tag string
			(cons 3 "Enter elevation") ; Prompt string
			(cons 40 5.0)         ; Text height
			(cons 7 "ARIAL")      ; Text style
			(cons 62 color-cyan)  ; Color
			(cons 8 layer) ; Layer
		)
	)
	
	; Outer Circle
	(entmake
		(list
			(cons 0 "CIRCLE")      
			(cons 10 (list 0 0 0)) ; Center Point
			; Radius: 7.71 copied from old block so it looks the same
			(cons 40 7.71)        ; Radius
			(cons 62 color-cyan)  ; Color
			(cons 8 layer) 		  ; Layer
		)
	)
	
	; Vertical Line
	(entmake
		(list
			(cons 0 "POLYLINE")
			(cons 10 (list 0 0 0))  ; Point is always zero
			(cons 70 1)             ; 1 = Closed Polyline
			(cons 62 color-cyan)  	; Color
			(cons 8 layer) 			; Layer
		)
	)
	(entmake
		(list
			(cons 0 "VERTEX")
			(cons 10 (list 7.71 0 0)) ; Top
		)
	)
	(entmake
		(list
			(cons 0 "VERTEX")
			(cons 10 (list -7.71 0 0))	; Bottom
		)
	)
	(entmake
		(list
			(cons 0 "SEQEND")
		)
	)
	
	; Horizontal Line
	(entmake
		(list
			(cons 0 "POLYLINE")
			(cons 10 (list 0 0 0))  ; Point is always zero
			(cons 70 1)             ; 1 = Closed Polyline
			(cons 62 color-cyan)  ; Color
			(cons 8 "FloorTags") ; Layer
		)
	)	
	(entmake
		(list
			(cons 0 "VERTEX")
			(cons 10 (list 0 -7.71 0)) ; Left
		)
	)
	(entmake
		(list
			(cons 0 "VERTEX")
			(cons 10 (list 0 7.71 0))	; Right
		)
	)
	(entmake
		(list
			(cons 0 "SEQEND")
		)
	)
	
	(entmake 
		(list
			(cons 0 "ENDBLK")
		)
	)
)
; Convert feet to inches
(defun feet->inches (feet)
    (* feet 12)
)
