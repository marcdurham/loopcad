(defun label-all-pipes ()
	(delete-all-pipe-labels)
	(make-pipe-labels)
)

(defun set-attribute (val tagname ename / ent) 
	(setq ent (entget (get-attribute tagname ename)))
	(setq ent 
		(subst 
			(cons 1 val)  ; New replacement value
			(assoc 1 ent) ; Old value
			ent           ; Entity list
		)
	)     
	(entmod ent)  
)

(defun get-attribute-value (tagname ename / attr-entity)
	(setq attr-entity (entget (get-attribute tagname ename)))
	; Return attribute value from: (1 . "VALUE HERE")
	(cdr (assoc 1 attr-entity))
)

(defun get-attribute (tagname ename / att)
	(foreach att (get-attributes ename)
		; Find by tag name: (2 . "TAG NAME")
		(if (str= tagname (cdr (assoc 2 (entget att)))) 
			att ; Returns entity name
		)
	)
)

; Returns a list of entity names that are owned by 'ename'
(defun get-attributes (ename / en attributes layer) 
	(setq attributes '())
	(setq en (entnext))
    (while en
		(if (and (= "ATTRIB" (get-etype en)) (= ename (get-owner-name en)))
			(progn
				(setq attributes (cons en attributes))
			)
		)
		(setq en (entnext en))
	)
	attributes
)

(defun make-pipe-labels ( / seg p v vertices label)	
	(setq p 0)	
	(foreach pipe (get-all-pipes)
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

(defun insert-head-label (point text)
	(insert-block point "HeadLabel.dwg" "HeadLabels")
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

