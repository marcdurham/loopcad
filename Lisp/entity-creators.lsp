; Quick ways to create entities without using commands

(defun make-circle (center radius color layer)
	(entmakex 
		(list 
			(cons 0 "CIRCLE")
			(cons 10 center)
			(cons 40 radius)
			(cons 8 layer) ; Layer 
			(cons 62 color) ; Color
		)
	)
)


(defun make-polyline (vertices color layer)
	(entmakex 
		(list 
			(cons 0 "POLYLINE")
			(cons 8 layer) ; Layer 
			(cons 62 color) ; Color
			(cons 40 pipe-width) ; Starting width
			(cons 41 pipe-width) ; Ending width
			(cons 10 '(0 0 0)) ; Always zero 'Dummy pont'
		)
	)
	(mapcar 
	    'make-vertex 
	    vertices
	)
	(entmakex (list (cons 0 "SEQEND")))
)
 