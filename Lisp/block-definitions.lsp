; Head label properties
(setq head-label:tag-string "HEADNUMBER")
(setq head-label:prompt "Head number label")
(setq head-label:label-color color-green)
(setq head-label:layer "HeadLabels")

; Tee label properties
(setq tee-label:tag-string "TEENUMBER")
(setq tee-label:prompt "Tee number label")
(setq tee-label:label-color color-green)
(setq tee-label:layer "TeeLabels")

(defun define-labels ()
	; These two block definitions are not used by any functions but they are defined so that
	; a user can use the "INSERT" command to insert them manually if they want.
	(define-label-block "HeadLabel" "HEADNUMBER" "Head number label" "H.0" color-blue "HeadLabels")
	(define-label-block "TeeLabel" "TEENUMBER" "Tee number label" "T.0" color-blue "TeeLabels")
	(define-head-block 12 "Head" "MODEL" "Head model" "MODEL" color-red "Heads")
	(define-head-block 14 "Head" "MODEL" "Head model" "MODEL" color-red "Heads")
	(define-head-block 16 "Head" "MODEL" "Head model" "MODEL" color-red "Heads")
	(define-head-block 18 "Head" "MODEL" "Head model" "MODEL" color-red "Heads")
	(define-head-block 20 "Head" "MODEL" "Head model" "MODEL" color-red "Heads")
	(princ "\nLabels defined.\n")
	(princ)
)

(defun define-label-block (block-name tag-string prompt default label-color layer)
	(entmake 
		(list
			(cons 0 "BLOCK")
			(cons 2 block-name) ; Block name
		)
	)
	(entmake 
		(append
			(list
				(cons 0 "ATTDEF")
				(cons 10 (list 2.0 2.0 0.0))
			)
			(node-label-props 
				block-name
				tag-string
				prompt
				default ; Text
				label-color
				layer
			)
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
	(entmake 
		(append
			(list
				(cons 0 "ATTDEF")
				(cons 10 (list 8.0 8.0 0.0))
			)
			(node-label-props 
				block-name
				tag-string
				prompt
				default ; Text
				label-color
				layer
			)
		)
	)
	(entmake
		(list
			(cons 0 "CIRCLE")      ; Inner Circle
			(cons 10 (list 0 0 0)) ; Center Point
			(cons 40 3.0)          ; Radius
			(cons 62 color-red)    ; Color
			(cons 8 layer)         ; Layer
		)
	)
	(entmake
		(list
			(cons 0 "CIRCLE")      ; Outer Circle
			(cons 10 (list 0 0 0)) ; Center Point
			(cons 40 6.0)          ; Radius
			(cons 62 color-red)    ; Color
			(cons 8 layer)         ; Layer
		)
	)
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
	(setq coverage-text 
		(strcat 
			(itoa coverage)
			"'  X  "
			(itoa coverage)
			"'"
		)
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

(defun feet->inches (feet)
    (* feet 12)
)
(defun inches->feet (inches)
	(/ inches 12)
)


; Append this list to an (0 . "ATTRIB") or (0 . "ATTDEF") 
; Example:  (append (list (cons 0 "ATTRIB") point) (head-label-props text))
(defun node-label-props (block-name tag-string prompt text label-color layer-name)
	(list 
		(cons 1 text) 
		(cons 2 tag-string) ; Tag string
		(cons 3 prompt) ; Prompt string
		(cons 40 5.0) ; Text height
		(cons 7 "ARIAL") ; Text style
		(cons 62 label-color) ; Color
		(cons 8 layer-name) ; Layer
	)
)

(defun head-props (block-name tag-string prompt text label-color layer-name)
	(list 
		(cons 1 text) 
		(cons 2 tag-string) ; Tag string
		(cons 3 prompt) ; Prompt string
		(cons 40 5.0) ; Text height
		(cons 7 "ARIAL") ; Text style
		(cons 62 label-color) ; Color
		(cons 8 layer-name) ; Layer
	)
)

(setq head-label:tag-string "HEADNUMBER")
(setq head-label:prompt "")
(setq head-label:color color-green)
(setq head-label:layer "HeadLabels")

