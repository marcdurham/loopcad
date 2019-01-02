(defun define-labels ()
	(define-label-block "HeadLabel" "HEADNUMBER" "Head number label" "H.0" color-green)
	(define-label-block "TeeLabel" "TEENUMBER" "Tee number label" "T.0" color-blue)
	(princ "\nLabels defined.\n")
	(princ)
)

(defun define-label-block (block-name tag-string prompt default label-color)
	(entmake 
		(list
			(cons 0 "BLOCK")
			(cons 2 block-name) ; Block name
		)
	)
	(entmake 
		(list
			(cons 0 "ATTDEF")
			(cons 10 (list 2.0 2.00 0.0))
			(cons 1 default) ; Default value
			(cons 7 "ARIAL") ; Text style
			(cons 62 label-color) ; Color
			(cons 40 3.0) ; Text height
			(cons 3 prompt) ; Prompt string
			(cons 2 tag-string) ; Tag string
			;(cons 70 2) ; Attribute flags: 
				; 1 = Attribute is invisible (does not appear)
				; 2 = This is a constant attribute
				; 4 = Verification is required on input of this attribute
				; 8 = Attribute is preset (no prompt during insertion)
		)
	)
	(entmake 
		(list
			(cons 0 "ENDBLK")
		)
	)
)