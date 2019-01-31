;(vl-registry-read "HKEY_CURRENT_USER\\Software\\LoopCalc\\ProgeCAD" "Test")

(defun greatest (a b)
  (if (> a b) a b)
)

(defun least (a b)
  (if (< a b) a b)
)

(defun average (a b)
	(/ (+ a b) 2)
)

(defun ents ( / en all all-lists)
	;(princ "\n\nEntity: ")
	;(princ (entget en))
	(setq all '())
	(setq en (entnext))
	(while en
		(setq all (cons en all))
		(setq en (entnext en))
	)
	(setq all-lists '())
	(foreach en all
		(setq all-lists (cons (entget en) all-lists))
	)
	all-lists
)

; Manual eyeball test
(defun test-midpoint ()
	(make-circle (midpoint (getpoint) (getpoint)) 10.0 color-green "Pipes")
)

(defun midpoint (a b)
	(list (average (getx a) (getx b)) (average (gety a) (gety b)) 0.0)
)

; Print the coordinates of the point.  For debugging.
(defun print-point (label point)
    (princ (strcat "\n" label ": "))
    (princ (car point))
	(princ ", ")
    (princ (cadr point))
	(princ)
)

; Get X coordinate from a point
(defun getx (point)
    (car point)
)

; Get Y coordinate from a point
(defun gety (point)
    (car (cdr point))
)

; Linear slope y/x of a line between points 'a' and 'b'
(defun slope (a b / xdiff ydiff)
	(setq xdiff (- (getx a) (getx b)))
	(setq ydiff (- (gety a) (gety b)))
	(if (= xdiff 0)
	    "Infinity"
		(if (= ydiff 0)
		    0
		    (/ ydiff xdiff)
		)
	)
)

; Negative reciprocal, used with slope to find perpendicular slope
(defun negative-reciprocal (x) (- 0 (/ 1 x)))

(setq dxf-point 10)

; Test with:
; (command "-PLINE" (get-vertices (car (get-all-pipes))) "")
(defun get-vertices (polyline / vertex remaining)
	(setq vertices '())
	(cond 
		((str= "LWPOLYLINE" (cdr (assoc 0 polyline)))
			(foreach property polyline
				(if (= 10 (car property)) 
					(setq vertices (cons (cdr property) vertices))
				)
			)
		)
		((str= "POLYLINE" (cdr (assoc 0 polyline)))
			(princ "\nNOT an LW Polyline box\n")
			(princ (entget (cdr (assoc -1 polyline))))
			(princ (entget (entnext (cdr (assoc -1 polyline)))))
			; TODO: Get the rest of the vertices
		)
	)
	(remove-repeated-points vertices)
)

(defun get-layer (entity-name)
	(cdr (assoc 8 (entget entity-name)))
)

(defun get-etype (entity-name)
	(cdr (assoc 0 (entget entity-name)))
)

(defun get-owner-name (entity-name)
	(cdr (assoc 330 (entget entity-name)))
)

(defun get-ins-point (entity)
    (if (= (type entity) "ENAME") 
	    (setq entity (entget entity))
	)
	(cdr (assoc 10 entity))
)

(defun ent-name (entity)
	(cdr (assoc -1 entity))
)

(defun get-x-scale (entity-name)
	(cdr (assoc 41 (entget entity-name)))
)

(defun get-x-scale (entity-name)
	(cdr (assoc 42 (entget entity-name)))
)

(defun get-z-scale (entity-name)
	(cdr (assoc 43 (entget entity-name)))
)

(defun get-block-name (entity-name)
	(cdr (assoc 2 (entget entity-name)))
)

(defun get-rotation-angle (entity-name)
	(cdr (assoc 50 (entget entity-name)))
)

(defun get-color (entity-name)
	(cdr (assoc 62 (entget entity-name)))
)

(defun ent-color (entity)
	(cdr (assoc 62 entity))
)

(defun str= (left right)
   (= (strcase left) (strcase right))
)

(defun strindexof (substring string)
   (vl-string-search (strcase substring) (strcase string))
)

(defun strcontains (substring string)
   (strindexof substring string)
)

(defun strstartswith (substring string)
   (= 0(strindexof substring string))
)

; Returns if 'items' contains 'item' works for strings only
; Uses case insenstive str= function.
(defun list-contains (item items / result)
	(setq result nil)
	(foreach i items 
		(if (str= i item)
			(setq result T)
		)
	)
	result
)



