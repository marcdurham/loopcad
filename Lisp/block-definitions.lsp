; Head label properties
(setq head-label:tag-string "HEADNUMBER")
(setq head-label:prompt "Head number label")
(setq head-label:label-color color-blue)
(setq head-label:layer "HeadLabels")

; Tee label properties
(setq tee-label:tag-string "TEENUMBER")
(setq tee-label:prompt "Tee number label")
(setq tee-label:label-color color-blue)
(setq tee-label:layer "TeeLabels")

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
	)
	(define-label-block 
		"TeeLabel" 
		tee-label:tag-string
		tee-label:prompt 
		"T.0" 
		tee-label:label-color
		tee-label:layer
	)
	(define-head-coverage 12)
	(define-head-coverage 14)
	(define-head-coverage 16)
	(define-head-coverage 18)
	(define-head-coverage 20)
	(princ "\nLabels defined.\n")
	(princ)
)

(defun define-head-coverage (coverage)
    (define-head-block coverage "Head" "MODEL" "Head model" "MODEL" color-red "Heads")
)

(defun define-label-block (block-name tag-string prompt default label-color layer)
	(entmake 
		(list
			(cons 0 "BLOCK")
			(cons 2 block-name) ; Block name
		)
	)
	(entmake 
		(list
			(cons 0 "ATTDEF")
			(cons 10 (list 9.132 0 0))
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
			; Insert Point: 9.132, 8.395 copied from old block so it looks the same
			(cons 10 (list 9.132 8.395 0.0))
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

; Convert feet to inches
(defun feet->inches (feet)
    (* feet 12)
)
