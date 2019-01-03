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
	(define-label-block "HeadLabel" "HEADNUMBER" "Head number label" "H.0" color-green "HeadLabels")
	(define-label-block "TeeLabel" "TEENUMBER" "Tee number label" "T.0" color-blue "TeeLabels")
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
			(head-label-props 
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

; Append this list to an (0 . "ATTRIB") or (0 . "ATTDEF") 
; Example:  (append (list (cons 0 "ATTRIB") point) (head-label-props text))
(defun head-label-props (block-name tag-string prompt text label-color layer)
	(list 
		(cons 1 text) 
		(cons 2 tag-string) ; Tag string
		(cons 3 prompt) ; Prompt string
		(cons 40 3.0) ; Text height
		(cons 7 "ARIAL") ; Text style
		(cons 62 label-color) ; Color
		(cons 8 layer) ; Layer
	)
)

(setq head-label:tag-string "HEADNUMBER")
(setq head-label:prompt "")
(setq head-label:color color-green)
(setq head-label:layer "HeadLabels")

