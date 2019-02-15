(defun label-all-pipes ( / p )
	(princ "\nLabeling pipes...\n")
	(delete-all-pipe-labels)
	(setq p (make-pipe-labels))
	(princ (strcat "\n" (itoa p) " pipes were labeled.\n"))
	(princ)
)

(defun label-all-nodes ( / n node label)
	(princ "\nLabeling nodes...\n")
	(setq n 1)
    (delete-blockrefs (get-all-head-labels))
	(foreach node (get-all-heads)
		(progn	
			(insert-head-label 
				(get-ins-point node) 
				(strcat "H." (itoa n))
			)
			(setq n (1+ n))
		)
	)
	(delete-blockrefs (get-all-tee-labels))
	(foreach node (get-all-tees)
		(progn		
			(insert-tee-label 
				(get-ins-point node) 
				(strcat "T." (itoa n))
			)
			(setq n (1+ n))
		)
	)
	; Domestic tees are already deleted with Tees above
	(foreach node (get-all-domestic-tees)
		(progn		
			(insert-domestic-tee-label 
				(get-ins-point node) 
				(strcat "D.T." (itoa n))
			)
			(setq n (1+ n))
		)
	)
	(delete-blockrefs (get-all-riser-labels)) 
	(foreach node (get-all-risers)
		(progn	
			; Find riser friends
			(insert-riser-label 
				(get-ins-point node)
				; Risers must be manually re-labeled for now
				(strcat "R." (itoa n))
			)
			(setq n (1+ n))
		)
	)
	(princ (strcat "\n" (itoa (- n 1)) " nodes were labeled.\n"))
	(princ)
)

(defun delete-blockrefs (blockrefs)
	(foreach blockref blockrefs
		(entdel (cdr (assoc -1 blockref)))
	)
)

(defun insert-head-label (point text)
	(insert-node-label 
		point 
		text 
		"HeadLabel"            ; block-name
		head-label:layer       ; layer
		head-label:tag-string  ; tag-string
		head-label:label-color ; label-color
		head-label:x-offset
		head-label:y-offset
	)
)

(defun insert-tee-label (point text)
	(insert-node-label 
		point 
		text 
		"TeeLabel"             ; block-name
		tee-label:layer        ; layer
		tee-label:tag-string   ; tag-string
		tee-label:label-color  ; label-color
		tee-label:x-offset
		tee-label:y-offset
	)
)

(defun insert-domestic-tee-label (point text)
	(insert-node-label 
		point 
		text 
		"TeeLabel"            ; block-name
		tee-label:layer       ; layer
		tee-label:tag-string  ; tag-string
		tee-label:label-color ; label-color
		tee-label:x-offset
		tee-label:y-offset
	)
)

(defun insert-riser-label (point text)
	(insert-node-label 
		point 
		text 
		"RiserLabel"  ; block-name
		riser-label:layer       ; layer
		riser-label:tag-string  ; tag-string
		riser-label:label-color ; label-color
		tee-label:x-offset
		tee-label:y-offset
	)
)

(defun insert-node-label (point text block-name layer-name tag-string label-color label-x-offset label-y-offset / e p)
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
		(list 
			(cons 0 "ATTRIB") ; Entity type
			(cons 10 (add-point-offset point label-x-offset label-y-offset)) ; Label insertion point
			(cons 1 text)          ; Text value
			(cons 2 tag-string)    ; Tag string
			(cons 3 "Node number:")        ; Prompt string
			(cons 40 5.0)          ; Text height
			(cons 7 "ARIAL")       ; Text style
			(cons 62 label-color)  ; Color
			(cons 8 layer-name)    ; Layer
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
	p ; Return number of pipes labled
)

(defun insert-pipe-label (point text)
	(make-text point text 4.0 color-blue "Pipe Labels")
)

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

(defun get-all-riser-labels ()
	(get-blocks (list "RiserLabels" "Riser Labels"))
)
