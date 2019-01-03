(defun label-all-pipes ()
	(delete-all-pipe-labels)
	(make-pipe-labels)
)

(defun label-all-nodes ( / n node label)
	(setq n 1)
	(foreach label (get-all-head-labels)
		(entdel (cdr (assoc -1 label)))
	)
	(foreach node (get-all-heads)
		(progn	
			(insert-head-label 
				(get-ins-point node) 
				(strcat "H." (itoa n))
			)
			(setq n (1+ n))
		)
	)
	(foreach label (get-all-tee-labels)
		(entdel (cdr (assoc -1 label)))
	)
	(foreach node (get-all-tees)
		(progn		
			(insert-tee-label 
				(get-ins-point node) 
				(strcat "T." (itoa n))
			)
			(setq n (1+ n))
		)
	)
)

(defun insert-head-label (point text / e p)
	(insert-node-label 
		point 
		text 
		"HeadLabel"  ; block-name
		"HeadLabels" ; layer
		"HEADNUMBER" ; tag-string
		color-green ; label-color
	)
)

(defun insert-tee-label (point text / e p)
	(insert-node-label 
		point 
		text 
		"TeeLabel"  ; block-name
		"TeeLabels" ; layer
		"TEENUMBER" ; tag-string
		color-green ; label-color
	)
)

(defun insert-node-label (point text block-name layer-name tag-string label-color / e p)
	(entmake
		(list 
			(cons 0 "INSERT")
			(cons 10 point) ; Insertion point
			(cons 2 block-name) ; Block name
			(cons 8 layer-name) ; Layer
			(cons 66 1) ; Attributes follow
		)
	)
	(setq e (entlast))
	(entmake
		(append 
			(list 
				(cons 0 "ATTRIB") ; Entity type
				(cons 10 (point-offset point 3.0 4.0)) ; Insertion point
			)
			;block-name tag-string prompt text label-color layer)
			(node-label-props 
				block-name
				tag-string
				"Node Number:" ; prompt
				text 
				label-color 
				layer-name
			)
		)
	)
	(entmake
		(list 
			(cons 0 "SEQEND") 
			(cons -2 e)
		)
	)
	(entupd e)
	(princ)
)

(defun point-offset (point x y)
	(list (+ x (getx point)) (+ y (gety point)))
)
 
(defun make-pipe-labels ( / seg p v vertices label)	
	(setq p 0)	
	(foreach pipe  (reverse (get-all-pipes))
		(setq v 0)				
		(setq vertices (get-vertices pipe))
		(while (< v (length vertices))
			(setq seg (segment v vertices))
			(if (< v (1- (length vertices))) ; not the last vertex index
				(progn
					(setq label (strcat "P" (itoa (1+ p))))
					(insert-pipe-label (midpoint (car seg) (cadr seg)) label)
				)
			)
			(setq v (1+ v))
		)
		(setq p (1+ p))
	)
	(princ)
)

(defun insert-pipe-label (point text)
	(make-text point text 4.0 color-blue "Pipe Labels")
)

; Delete me
;(defun get-all-head-labels ( / en ent labels layer)
;	(setq labels '())
;	(setq en (entnext))
;   (while en
;		(if (and (str= "HeadLabels" (get-layer en))
;		        (str= (get-etype en) "INSERT")
;			)
;			(setq labels (cons en labels))
;		)
;		(setq en (entnext en))
;	)
;	labels
;)

(defun get-all-pipe-labels ( / en ent labels layer) 
	(setq labels '())
	(setq en (entnext))
    (while en
	    ;(setq ent (entget en))
		(if (and (str= "Pipe Labels" (get-layer en))
		        (str= (get-etype en) "TEXT")
			)
			;(setq labels (cons ent labels))
			(setq labels (cons en labels))
		)
		(setq en (entnext en))
	)
	labels
)

(defun delete-all-pipe-labels ( / label)
	(foreach label (get-all-pipe-labels)
		(entdel label)
	)
)

(defun get-all-head-labels ()
	(get-blocks (list "HeadLabels" "Head Labels"))
)

(defun get-all-tee-labels ()
	(get-blocks (list "TeeLabels" "Tee Labels"))
)
