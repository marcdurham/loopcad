(defun defblock-head-label ()
	(entmake 
		(list
			(cons 0 "BLOCK")
			(cons 2 "HeadLabel") ; Block name
		)
	)
	(entmake 
		(list
			(cons 0 "POINT")
			(cons 10 (list 0.0 0.0 0.0))
		)
	)
	(entmake 
		(list
			(cons 0 "ATTDEF")
			(cons 10 (list 2.0 2.00 0.0))
			(cons 1 "H.0") ; Default value
			(cons 7 "ARIAL") ; Text style
			(cons 62 color-green) ; Color
			(cons 40 3.0) ; Text height
			(cons 3 "Head number label:") ; Prompt string
			(cons 2 "HEADNUMBER") ; Tag string
			(cons 70 2) ; Attribute flags: 
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

(defblock-head-label)