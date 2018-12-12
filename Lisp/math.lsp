; Thse exist, replace them
(defun greatest (a b)
  (if (> a b) a b)
)
(defun least (a b)
  (if (< a b) a b)
)

;(vl-registry-read "HKEY_CURRENT_USER\\Software\\LoopCalc\\ProgeCAD" "Test")

(defun test-near-line ( / a b h)
	(setq a (getpoint))
	(setq b (getpoint))

	(print-point "a" a)
	(print-point "b" b)
	(while T
	    (setq h (getpoint))
		(print-point "h" h)

		(if (near-line h a b)
			(command "-COLOR" "green")
			(command "-COLOR" "red")
		)
		(command "-CIRCLE" h 1.0)
	)
)

; Is h near or on the line segment between a and b?
(defun near-line (p a b / int pp)
    ; Draw an imaginary line from p perpendicular to a-b
    (setq pp (perp-point (list a b) p))
	(if (in-box p a b)
	    (progn
		    ; Find the intersection between the imaginary perpendicular
			; line and a-b.
			(setq int (inters a b pp p nil))
			(if int
				(if (< (distance p int) 5.0)
					T ; near the line
					nil
				)
				(if (are-same-point p pp) 
					T ; on the line
					nil
				)
			)
		)
		nil
	)
)

(defun are-same-point (a b)
    (and (= (getx a) (getx b)) (= (gety a) (gety b)))
)

(defun test-in-box ( / )
    (and
		(not (in-box (list 1001 0 0) (list -500 -400 0) (list 1000 2000 0)))
		(not (in-box (list -401 0 0) (list -500 -400 0) (list 1000 2000 0)))
		(not (in-box (list 0 -401 0) (list -500 -400 0) (list 1000 2000 0)))
		(not (in-box (list 0 2001 0) (list -500 -400 0) (list 1000 2000 0)))
		(in-box (list 0 0 0) (list -500 -400 0) (list 1000 2000 0))
	)
)

; Is point 'p' in the box between the corners defined by 'a' and 'b'
(defun in-box (p a b / x y maxx maxy minx miny margin)
    (setq margin 5.0)
    (setq maxx (+ (max (getx a) (getx b)) margin))
	(setq maxy (+ (max (gety a) (gety b)) margin))
	(setq minx (- (min (getx a) (getx b)) margin))
	(setq miny (- (min (gety a) (gety b)) margin))
	(setq x (getx p))
	(setq y (gety p))
	(if (and (< x maxx)
			(< y maxy)
			(> x minx)
			(> y miny)
		)
	    T    ; p is in the box
		nil  ; p is not in the box
	)
)

(defun test-perp-point ( / )
    (perp-point (list (list 100 100 0) (list 2000 100 0)) (list 1000 1000 0))
	; should return '(1000 100)
)

; Perpendicular line from 'point' through 'line'
; If the line has a zero or infinite slope return 'point'
(defun perp-point (line point / x y s perp-slope)
	(setq s (slope (car line) (cadr line)))
	(cond ((= 0 s) ; Zero slope				
			(list (getx point) (gety point))
		)
		((= "Infinity" s) ; Infinite slope			
			(list (getx point) (gety point))
		)
		(T ; Normal slope
			(list  
			    (+ 100 (getx point)) 
				(+ (* 100 (negative-reciprocal s)) (gety point))
			)
		)
	)
)

(defun print-point (label point)
    (princ (strcat "\n" label ": "))
    (princ (car point))
	(princ ", ")
    (princ (cadr point))
	(princ)
)

(defun negative-reciprocal (x) (- 0 (/ 1 x)))

(defun getx (point)
    (car point)
)

(defun gety (point)
    (car (cdr point))
)

;(defun test-slope (a b)

;)

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

(setq dxf-point 10)

; Draw pipes from vertices 
; (foreach pipe (get-all-pipes) (command "-PLINE" (get-vertices pipe) ""))

; (foreach pipe (get-all-pipes) (pipe-draw "1/2" (get-vertices pipe)))
; Temp function
(defun test-pipe-draw ()
(foreach pipe (get-all-pipes) (pipe-draw "1/2" (get-vertices pipe)))
)

; Test with
;  (command "-PLINE" (get-vertices (car (get-all-pipes))) "")
(defun get-vertices (polyline / vertex remaining)
	(setq vertices '())
	(foreach property polyline
	    (if (= 10 (car property)) 
		    (setq vertices (cons (cdr property) vertices))
		)
	)
	vertices
)

(defun get-all-pipes () 
    (get-pipes (entnext))
)

(defun get-pipes (en / ent pipes layer) 
	(setq pipes '())
    (while en
	    (setq ent (entget en))
		(if (and (strstartswith "PIPES" (get-layer en))
		         (str= (get-etype en) "LWPOLYLINE")
			)
			(setq pipes (cons ent pipes))
		)
		(setq en (entnext en))
	)
	pipes
)

(defun get-layer (entity-name)
	(cdr (assoc 8 (entget entity-name)))
)

(defun get-etype (entity-name)
	(cdr (assoc 0 (entget entity-name)))
)

(defun get-ins-point (entity-name)
	(cdr (assoc 10 (entget entity-name)))
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

; Not tested yet, I think 62 is the color
(defun get-color (entity-name)
	(cdr (assoc 62 (entget entity-name)))
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





