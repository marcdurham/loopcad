(setq color-red 1)
(setq color-yellow 2)
(setq color-green 3)
(setq color-cyan 4)
(setq color-blue 5)
(setq color-magenta 6)
(setq color-black-white 7)
(setq color-dark-gray 8)
(setq color-light-gray 9)

(setq size-color-list 
	(list 
		(cons "1/2" color-green)
		(cons "3/4" color-red)
		(cons "1" 150) ; Since this is listed first it is the default 1 inch color
		(cons "1" color-blue)
		(cons "1-1/4" color-magenta)
	)
)

(defun pipe-color-size (color / size)
	(assoc-cdr color size-color-list)
)

(defun pipe-size-color (size)
	(cdr (assoc size size-color-list))
)

; Return the first element from a list of dotted pairs by matching
; the second element.
; Like assoc but find the value by the second (cadr) value 
; insead of the first (car)
(defun assoc-cdr (val items / key)
	(foreach item items
	    (if (= val (cdr item))
			(setq key (car item))
		)
	)
	key
)