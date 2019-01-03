(defun label-all-pipes ()
	(delete-all-pipe-labels)
	(make-pipe-labels)
)

(defun label-all-nodes ( / n)
	(setq n 1)
	(foreach node (get-all-nodes)
		(insert-head-label 
			(get-ins-point node) 
			(strcat "H." (itoa n))
		)
	)
)

(defun insert-head-label (point text / e p)
	(entmake
		(list 
			(cons 0 "INSERT")
			(cons 10 point)
			(cons 2 "HeadLabel")
			(cons 8 "HeadLabels")
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
			(head-label-props text)
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

; Append this list to an (0 . "ATTRIB") or (0 . "ATTDEF") 
; Example:  (append (list (cons 0 "ATTRIB") point) (head-label-props text))
(defun head-label-props (text)
	(list 
		(cons 1 text) 
		(cons 2 "HEADNUMBER") ; Tag string
		(cons 40 3.0) ; Text height
		(cons 7 "ARIAL") ; Text style
		(cons 62 color-red) ; Color
		(cons 8 "HeadLabels") ; Layer
	)
)

(defun set-attribute (ename tagname val / em ent)
	(setq en (get-attribute ename tagname))
	(if en
		(progn
			(setq ent (entget en))
			(setq ent 
				(subst 
					(cons 1 val)  ; New replacement value
					(assoc 1 ent) ; Old value
					ent           ; Entity list
				)
			)     
			(entmod ent)
		)
		nil
	)
)

(defun get-attribute-value (ename tagname / attr-entity)
	(setq attr-entity (entget (get-attribute ename tagname)))
	; Return attribute value from: (1 . "VALUE HERE")
	(cdr (assoc 1 attr-entity))
)

; Returns entity name of attribute entity owned by 'ename'
(defun get-attribute (ename tagname / att)
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

(defun get-all-head-labels ( / en ent labels layer) 
	(setq labels '())
	(setq en (entnext))
    (while en
		(if (and (str= "HeadLabels" (get-layer en))
		        (str= (get-etype en) "INSERT")
			)
			(setq labels (cons en labels))
		)
		(setq en (entnext en))
	)
	labels
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

