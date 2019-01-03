
(defun get-blocks ( layers / en ent blocks) 
	(setq blocks '())
	(setq en (entnext))
    (while en
	    (setq ent (entget en))
		(if (and (list-contains (get-layer en) layers)
		        (str= (get-etype en) "INSERT")
			)
			(setq blocks (cons ent blocks))
		)
		(setq en (entnext en))
	)
	blocks
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

; Returns entity name of attribute entity owned by 'ename'
(defun get-attribute (ename tagname / att)
	(foreach att (get-attributes ename)
		; Find by tag name: (2 . "TAG NAME")
		(if (str= tagname (cdr (assoc 2 (entget att)))) 
			att ; Returns entity name
		)
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
