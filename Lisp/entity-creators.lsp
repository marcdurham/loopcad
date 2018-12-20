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

(defun make-text (point text height color layer)
	(entmakex 
		(list 
			(cons 0 "TEXT")
			(cons 10 point)
			(cons 1 text)
			(cons 40 height)
			(cons 62 color)
			(cons 8 layer) ; Layer 	
		)
	)
)
				 
(defun make-mtext (pt text)
	(entmakex 
		(list 
			(cons 0 "MTEXT")         
			(cons 100 "AcDbEntity")
			(cons 100 "AcDbMText")
			(cons 10 pt)
			(cons 1 text)
		)
	)
)

(defun make-block-insert (point block-name layer)
	(entmake ; Removed x
		(list 
			(cons 0 "INSERT")
			(cons 10 point)
			(cons 2 block-name)
			(cons 8 layer)
			(cons 66 1)
		)
	)
	(entlast)
)