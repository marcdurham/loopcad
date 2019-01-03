; Functions called by the BREAK-PIPES command.
(defun break-pipes-delete-old ( / old-pipes)
	(setq old-pipes (get-all-pipes))
	(foreach pipe (break-all-pipes)
		(make-pipe (car pipe) (cdr pipe))
	)
	; Delete all old pipes
	(foreach pipe old-pipes (entdel (cdr (assoc -1 pipe))))
)
	
(defun break-all-pipes ( / size node-point all-nodes seg new-vertices new-pipes old-pipes pt start end vertex i vertices)
	(setq new-pipes '())
	(setq old-pipes (get-all-pipes))
	(setq all-nodes (get-all-nodes))
	(foreach pipe old-pipes
		(setq i 0)
		(setq size (get-pipe-size pipe))					
		(setq vertices (get-vertices pipe))
		(setq new-vertices '())
		(while (< i (length vertices))
			(setq vertex (nth i vertices))			
			(setq new-vertices (cons vertex new-vertices))
			(setq seg (segment i vertices))
			(if (and (> i 0) (< i (1- (length vertices)))) ; not the first or last vertex index
				(foreach node all-nodes
					(setq node-point (get-ins-point node))
					(setq dist (distance vertex node-point))
					(if (< dist near-line-margin)
						(progn 
							(if (> (length new-vertices) 0)
								(progn
									(setq new-vertices (cons size new-vertices))
									(setq new-pipes (cons new-vertices new-pipes))
								)
							)							

							(setq new-vertices '())
							(setq new-vertices (cons vertex new-vertices))
						)
					)
				)
			)
			(setq i (1+ i))
		)
		(if (> (length new-vertices) 0)
			(progn
				(setq new-vertices (cons size new-vertices))
				(setq new-pipes (cons new-vertices new-pipes))
			)
		)
	)
	new-pipes
)

; Should return five points 
(defun test-remove-repeated-points ()
	(remove-repeated-points 
		'(
			(1.0 1.0 0.0)
			
			(1.0 2.0 0.0)
			
			(2.0 3.0 0.0)
			(2.0 3.0 0.0)
			(2.0 3.0 0.0)
			
			(3.0 3.0 0.0)
			(3.0 3.0 0.0)
			
			(3.0 4.0 0.0)
		)
	)
)

(defun remove-repeated-points (points / last-pt output)
	(setq output '())
	(foreach point points
		(if (not (are-same-point last-pt point))
			(progn
				(setq output (cons point output))
				(setq last-pt point)
			)
			(progn
			
				(setq last-pt point)
			)
		)
	)
	(reverse output)
)


(defun test-segments ( / pipe)
    (foreach pipe (get-all-pipes)
	    (princ "\nPipe\n")
	    (princ (segments pipe))
	)
)

(defun segments (polyline / i next last-i seg vertices output)
    (setq output '())
	(setq vertices (get-vertices polyline))
    (setq last-i (- (length vertices) 2)) ; grabbing pairs, so don't grab the last one
	(setq i 0)
	(while (<= i last-i)
	    (setq seg (segment i vertices))
		(setq output (cons seg output))
	    (setq i (1+ i))
	)
	(reverse output)
)

(defun segment (index vertices / next output)
    (setq output '())
	(setq next (1+ index))    			
	(setq output (cons (nth next vertices) output))
	(setq output (cons (nth index vertices) output))
)

; Manual test, click the ends of the line, then click points near
; the line. Small red circles appear if they are not near the line
; and green circles appear that are near the line.
(defun test-near-line ( / a b p)
	(setq a (getpoint))
	(setq b (getpoint))

	(print-point "a" a)
	(print-point "b" b)
	(while T
	    (setq p (getpoint))
		(print-point "p" p)

		(if (near-line p a b)
			(command "-COLOR" "green")
			(command "-COLOR" "red")
		)
		(command "-CIRCLE" p 1.0)
	)
)

; How close a point has to be to a line to be considered near it.
(setq near-line-margin 5.0)

; *** This function not called from anywhere anymore ***
; *** I might keep it around because it could be handy ***

; Is 'p' near or on the line segment between 'a' and 'b'?
(defun near-line (p a b / int pp)
    ; Draw an imaginary line from p perpendicular to a-b
    (setq pp (perp-point (list a b) p))
	(if (in-box p a b)
	    (progn
		    ; Find the intersection between the imaginary perpendicular
			; line and a-b.
			(setq int (inters a b pp p nil))
			(if int
				(if (< (distance p int) near-line-margin)
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


; Compares only x and y but not z of 'a' and 'b' coordinates
(defun are-same-point (a b)
    (and (= (getx a) (getx b)) (= (gety a) (gety b)))
)

; Auto test 'in-box' function
(defun test-auto-in-box ( / a b)
    (princ "\ntest-auto-in-box: ")
	(setq a (list -500 -400 0) )
	(setq b (list 1000 2000 0))
    (if (and
			(not (in-box (list 1001 0 0) a b))
			(not (in-box (list -401 0 0) a b))
			(not (in-box (list 0 -401 0) a b))
			(not (in-box (list 0 2001 0) a b))
			(in-box (list 0 0 0) a b) ; This one should be in the box
		)
		(princ "PASS\n")
		(princ "FAIL\n")
	)
	(princ)
)

; Is point 'p' in the box between the corners defined by 'a' and 'b'
(defun in-box (p a b / x y maxx maxy minx miny)
    (setq maxx (+ (max (getx a) (getx b)) near-line-margin))
	(setq maxy (+ (max (gety a) (gety b)) near-line-margin))
	(setq minx (- (min (getx a) (getx b)) near-line-margin))
	(setq miny (- (min (gety a) (gety b)) near-line-margin))
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

; Should return '(1000 100)
(defun test-perp-point ( / )
    (perp-point (list (list 100 100 0) (list 2000 100 0)) (list 1000 1000 0))
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

(defun get-all-pipes ( / en ent pipes layer) 
	(setq pipes '())
	(setq en (entnext))
    (while en
	    (setq ent (entget en))
		(if (and (or (strstartswith "PIPES." (get-layer en))
					(strstartswith "PIPE." (get-layer en))
					(str= "PIPE" (get-layer en))
					(str= "PIPES" (get-layer en))
				)
		        (or (str= "LWPOLYLINE"(get-etype en))
					(str= "POLYLINE" (get-etype en))
				)
			)
			(setq pipes (cons ent pipes))
		)
		(setq en (entnext en))
	)
	pipes
)

(defun get-all-heads ()
	(get-nodes (list "HEADS"))
)

(defun get-all-tees ()
	(get-nodes (list "TEES"))
)

(defun get-nodes ( layers / en ent nodes) 
	(setq nodes '())
	(setq en (entnext))
    (while en
	    (setq ent (entget en))
		(if (and (list-contains (get-layer en) layers)
				;(or (str= "HEADS" (get-layer en))
		        ;    (str= "TEES" (get-layer en))
				;	(str= "FLOOR CONNECTORS" (get-layer en))
				;)
		        (str= (get-etype en) "INSERT")
			)
			(setq nodes (cons ent nodes))
		)
		(setq en (entnext en))
	)
	nodes
)

