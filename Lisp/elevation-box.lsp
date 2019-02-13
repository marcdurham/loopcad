(defun elevation-box-draw (/ a b p1 p2 p3 p4 top bottom left right)
	(setq temperror *error*)
	(defun *error* (message)
		(princ)
		(princ "Error")
		(princ message)
		(princ)
		(command "-COLOR" "BYLAYER")
		(command "-LAYER" "SET" "0" "")
		(setq *error* temperror)
	)

	(setq a (getpoint "\nElevation box first corner:"))
	(setq b (getcorner a))
		(if (null elevation) 
	    (setq elevation 100)
		(princ (strcat "Default elevation set to " (itoa elevation)))
	)
	(setq elevation (getint (strcat "\nEnter elevation (ft): <" (itoa elevation) ">")))
	
	(setq right (greatest (car a) (car b)))
	(setq top (greatest (cadr a) (cadr b)))
	(setq left (least (car a) (car b)))
	(setq bottom (least (cadr a) (cadr b)))
	(setq p1 (list left top))
	(setq p2 (list right top))
	(setq p3 (list right bottom))
	(setq p4 (list left bottom))
	
	;(command "-COLOR" "BYLAYER" "")
	;(command "-LINETYPE" "SET" "Continuous" "")
	;(command "-LAYER" "NEW" "ElevationBox" "")
	;(command "-LAYER" "COLOR" "Magenta" "ElevationBox" "")
	;(command "-LAYER" "SET" "ElevationBox" "")
	;(command "-PLINE" p1 p2 p3 p4 p1 "")
	;(command "-MTEXT" p1 p3 (strcat "Elevation " (itoa elevation)) "") 
	;(command "-COLOR" "BYLAYER")
	;(command "-LAYER" "SET" "0" "")
	
	(entmake
		(list
			(cons 0 "POLYLINE")
			(cons 10 (list 0 0 0))  ; Point is always zero
			(cons 70 1)             ; 1 = Closed Polyline
			(cons 62 color-magenta)  ; Color
			(cons 8 "ElevationBox") ; Layer
		)
	)
	(entmake
		(list
			(cons 0 "VERTEX")
			(cons 10 p1) ; Lower Left
		)
	)
	(entmake
		(list
			(cons 0 "VERTEX")
			(cons 10 p2)	; Lower Right
		)
	)
	(entmake
		(list
			(cons 0 "VERTEX")
			(cons 10 p3)	; Upper Right
		)
	)
	(entmake
		(list
			(cons 0 "VERTEX")
			(cons 10 p4)	; Upper Left
		)
	)
	(entmake
		(list
			(cons 0 "SEQEND")
		)
	)
	(entmake
		(list
			(cons 0 "MTEXT")
			(cons 10 p1)
			(cons 40 10.0) ; Text Height
			(cons 41 1000.0) ; Reference Width 
			(cons 11 (list 1.0 0.0 0.0))
			(cons 71 1)    ; Attachment point: 1 = Top left
			(cons 72 1)    ; Drawing direction: 1 = Left to right
			;(cons 73 1)    ; MText line spacing style: 1 = At least
			(cons 1 (strcat "Elevation " (itoa elevation))) ; Text Value
			(cons 62 color-magenta) ; Color
			(cons 8 "ElevationBox")  ; Layer
		)
	)
	
	(setq *error* temperror)
	(princ)
)

(defun get-elevation-boxes ( / en ent boxes layer) 
	(setq boxes '())
	(setq en (entnext))
    (while en
	    (setq ent (entget en))
		(if (and (or (str= "ElevationBox" (get-layer en))
					(str= "ElevationBoxes" (get-layer en))
					(str= "Elevation Box" (get-layer en))
					(str= "Elevation Boxes" (get-layer en))
				)
		        (or (str= "LWPOLYLINE"(get-etype en))
					(str= "POLYLINE" (get-etype en))
				)
			)
			(setq boxes (cons ent boxes))
		)
		(setq en (entnext en))
	)
	boxes
)

(defun get-elevation-text ( / en ent boxes layer) 
	(setq boxes '())
	(setq en (entnext))
    (while en
	    (setq ent (entget en))
		(if (and (or (str= "ElevationBox" (get-layer en))
					(str= "ElevationBoxes" (get-layer en))
					(str= "Elevation Box" (get-layer en))
					(str= "Elevation Boxes" (get-layer en))
				)
		        (or (str= "MTEXT"(get-etype en))
					(str= "TEXT" (get-etype en))
				)
			)
			(setq boxes (cons ent boxes))
		)
		(setq en (entnext en))
	)
	boxes
)

(defun get-polyline-vertices ( ent / en vertex vertices) 
	(setq vertices '())
	(setq en (cdr (assoc -1 ent)))
	(setq en (entnext en))
	(setq ent (entget en))
    (while en
		(cond ((str= "VERTEX" (cdr (assoc 0 ent)))
				(setq ent (entget en))
				(setq vertex (cdr (assoc 10 ent)))
				(setq vertices (cons vertex vertices))
				(setq en (entnext en))
				(setq ent (entget en))
			)
			(T (setq en nil))
		)
	)
	vertices
)

(defun test-ebox ( / )
	;(princ "\nTesting: test-ebox\n")
	; TODO: Add some elevation boxes, 2 at least
	(princ "\nShould return 102: ")
	(princ (find-ebox (list 4342.29 1633.89 0.000000)))
	(princ "\nShould return 109: ")
	(princ (find-ebox (list 4224.10 1672.70 0.000000)))
	(princ "\n")
)

(defun find-ebox ( p / box boxes result a b i ii ar areas all-areas m vertex vertices text-box text-boxes output)
	(setq areas '())
	;;(princ "\nGetting elevation boxes...")
	(setq boxes (get-elevation-boxes))
	;;(princ "\nElevation Boxes Count: ")
	;;(princ (itoa (length boxes)))
	;;(princ "\n")
	(foreach box boxes
		(progn 
			;;(princ "\nPoint: ")
			;;(princ p)
			;;(princ "\nBox: ")
			;;(princ box)
			;(princ "\na: ")
			;(princ (car (corners box)))
			;(princ "\nb: ")
			;(princ (cadr (corners box)))
			;;(princ "\nBox Area: ")
			(setq a (car (corners box)))
			(setq b (cadr (corners box)))
			(setq ar (area a b))
			;;(princ (rtos ar 2 1))
			;;(princ "\n")
			(setq all-areas (append all-areas (list ar)))
			
			(if (in-ebox p box)
				(progn 
					;;(princ "\np is in box: ")
					;;(princ box)
					;;(princ "\n")
					(setq result box)
					(setq areas (append areas (list ar)))
				)
				(progn 
					;;(princ "\np NOT in box: ")
					;;(princ box)
					;;(princ "\n")
				)
			)
		)
	)
	;;(princ "\nAll Areas: ")
	;;(princ (itoa (length all-areas)))
	;;(princ "  :")
	;;(princ all-areas)
	;;(princ "\n")
	
	;;(princ "\nIn Areas: ")
	;;(princ (itoa (length areas)))
	;;(princ "  :")
	;;(princ areas)
	;;(princ "\n")
	;(terpri)
	;(min areas)
	(setq m (apply 'min areas))
	(setq i (index-of m areas))
	;;(princ "\nSmallest In-Area Index: ")
	;;(princ (itoa i))
	;;(princ "\n")
	
	(setq ii (index-of m all-areas))
	;;(princ "\nSmallest All-Area Index: ")
	;;(princ (itoa ii))
	;;(princ "\n")
	
	;;(princ "\nSmallest Box: ")
	;;(princ (nth ii boxes))
	(setq vertices (get-polyline-vertices (nth ii boxes)))
	;;(princ "\nSmallest Box Vertices: ")
	;;(princ vertices)
	;;(princ "\n")
	
	;;(princ "\nFinding elevation text for smallest box...\n")
	(setq text-boxes (get-elevation-text))
	(foreach vertex vertices
		;;(princ "\nVertex: ")
		;;(princ vertex)
		(foreach text-box text-boxes
			;;(princ "\nBox Ins Pt: ")
			;;(princ (get-ins-point text-box))
			;;(princ " p: ")
			;;(princ p)
			;;(princ "\n")
			; TODO: check if box is on a vertex	
			(if (= vertex (get-ins-point text-box))
				(progn 
					;;(princ "\nMatch\n")			
					(setq output (elevation-from text-box))								     
				)
			)
		)
	)
	output
)

; Get numeric elevation value from MText entity
(defun elevation-from ( text-box / )
	; Input Example: "Elevation 999"
	; Digits start at position 11
	(substr (text-from text-box) 11)
)

; Get text from MText entity
(defun text-from ( text-box / )
	(cdr (assoc 1 text-box))					
)

; Is the point 'p' inside the polyline 'box'
(defun in-ebox ( p box / a b vertices )
	(setq a (car (corners box))) ; First corner
	(setq b (cadr (corners box))) ; Opposite corner
	(in-box p a b)
)

; Returns opposite corners of a box made of a four point polyline
(defun corners ( rectangle / a b vertices )
	(setq vertices (get-polyline-vertices rectangle))
	(setq a (nth 0 vertices)) ; First corner
	(setq b (nth 2 vertices)) ; Opposite corner	
	(list a b )
)