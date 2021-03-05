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

; Area of a rectangle, input opposite corners a and b
(defun area (a b / xdiff ydiff)
    (setq xdiff (- (getx a) (getx b)))
    (setq ydiff (- (gety a) (gety b)))
    (* xdiff ydiff)
)

; Index of a member of a list
(defun index-of (item  lst / )
    (member item lst)
    (- (length lst) (length (member item lst)))
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

(defun get-ename (entity)
    (cdr (assoc -1 entity))
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

; Finds if two numbers are almost the same, within the margin
(defun approx (a b margin / )
    (< (abs (- a b)) margin)
)

; Are all items in both lists almost the same, within the margin
; Uses the approx function
(defun lists-approx (a-list b-list margin / a b i a-length different)
    (setq a-length (length a-list))
    (setq i 0)
    (if (not (= a-length (length b-list)))
        (setq different T) ; List are of different lengths
        (while (< i a-length)
            (if (not (approx (nth i a-list) (nth i b-list) margin))
                (setq different T)
            )
            (setq i (1+ i))
        )    
    )
    (not different)
)

(defun test-assoc-approx ( / )
    (princ "\ntest-assoc-approx: ")
    (if (= 
            (list (list 25 33.4 4) 66 77) ; Expected Output
            (assoc-approx                   ; Function Under Test
                (list 25.5 33.0 4.0) 
                (list 
                    (list (list 55 22 2.30) 55 66) 
                    (list (list 25.00 33.4 4.0) 66.0 77.0) 
                    (list (list 1.0 2.0) 11 22)
                ) 
                1.0
            )
        )
        (princ "PASS\n")
        (princ "FAIL\n")
    )
    (princ)
)

; Same as the assoc function, but uses lists-approx and approx instead of =
(defun assoc-approx (target items margin / item found)
    (foreach item items
        (if (lists-approx (car item) target margin)
            (setq found item)
        )
    )
    found
)
