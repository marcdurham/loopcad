; Thse exist, replace them
(defun greatest (a b)
  (if (> a b) a b)
)
(defun least (a b)
  (if (< a b) a b)
)

;(vl-registry-read "HKEY_CURRENT_USER\\Software\\LoopCalc\\ProgeCAD" "Test")
(defun test-thing ( / a b h)
    ;(setq a (car (get-vertices (car (get-all-pipes)))))
    ;(setq b (car (cdr (get-vertices (car (get-all-pipes))))))
	(setq a (getpoint))
	(setq b (getpoint))
	(setq h (getpoint))
	
	(print-point "a" a)
	(print-point "b" b)
	(print-point "h" h)


	(if (near-line h a b)
	    (command "-CIRCLE" h 5.0)
		;(command "-TEXT h 5.0 "NO" "")
	)
)

(defun near-line (h a b / int pp)
    (setq pp (perp-point (list a b) h))
	(if (in-box h a b)
	    (progn
		    (princ "\nIn the box\n")
			(setq int (inters a b pp h nil))
			(if int
                (progn
				    (princ "\ndistance: ")
					(princ (rtos (distance h int) 2 4))
					(princ "\n")
					(if (< (distance h int) 5.0)
						T ; near the line
						nil
					)
				)
				(if (same-point h pp) 
					T ; on the line
					nil
				)
			)
		)
		(progn
			(princ "\nOut of the box\n")
			nil
		)
	)
)

(defun same-point (a b)
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

; Perpendicular line from point through line 
(defun perp-point (line point / x y perp-slope)
    (setq x (getx point))
	(setq y (gety point))
	(princ "\n")
	(princ line)
	(setq sl (slope (car line) (cadr line)))
	
	(cond ((= 0 sl)
		    (princ "\nZERO Slope\n")
			;(command "-CIRCLE" point 10.0)
            ;(command "-CIRCLE" point 11.0) 	
            ;(command "-CIRCLE" point 12.0) 				
			(list (getx point) (gety point))
		)
		((= "Infinity" sl)
		    (princ "\nINFINITY Slope\n")
			;(command "-CIRCLE" point 9.0)
            ;(command "-CIRCLE" point 10.0) 			
			(list (getx point) (gety point))
		)
		(T 
		    (princ "\nNORMAL Slope\n")
		    (princ (strcat "\nslope: " (rtos sl 5 4)))
			(setq perp-slope (negative-reciprocal sl))
			(princ (strcat "\nperp-slope: " (rtos perp-slope 5 4)))
			(setq newX (* perp-slope (+ 1 y)))
			(princ (strcat "\nnewX: " (rtos newX 5 4)))
			(princ (strcat "\ny: " (rtos y 5 4)))
			;(list (+ 1 x) newY)
			(list  (+ 100 x) (+ (* 100 perp-slope) y))
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





