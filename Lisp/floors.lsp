(defun floor-tag-insert ( / p old-osmode elevation floor-name)
	(setq old-osmode (getvar "OSMODE"))
    (defun *error* (message)
        (princ)
        (princ message)
        (princ)
        (setvar "OSMODE" old-osmode)
        (setvar "LWDISPLAY" 1)
    )
	(setq p (getpoint))
	(setq floor-name (getstring "Enter floor name"))
	(setq elevation (get-elevation p))
    (setvar "INSUNITS" 0) ; This line prevents inserted block refs from having
						  ; a different scale, being 12 time bigger than they 
						  ; should be
    (setvar "OSMODE" 0)
    (command "-LAYER" "NEW" "FloorTags" "")
    (command "-LAYER" "COLOR" "White" "FloorTags" "")
    (setvar "LWDISPLAY" 0)
    (command "-LAYER" "SET" "FloorTags" "")
    (command "-INSERT" "FloorTag.dwg" p 1.0 1.0 0 floor-name elevation)  
)

(defun floor-connector-insert ( / p w old-osmode tag tags tag-elevation p-elevation offset tag-offset)
	(setq old-osmode (getvar "OSMODE"))
    (defun *error* (message)
        (princ)
        (princ message)
        (princ)
        (setvar "OSMODE" old-osmode)
        (setvar "LWDISPLAY" 1)
    )
    (setvar "INSUNITS" 0) ; This line prevents inserted block refs from having
						  ; a different scale, being 12 time bigger than they 
					 f	  ; should be
    (setvar "OSMODE" 0)
    (command "-LAYER" "NEW" "FloorConnectors" "")
    (command "-LAYER" "COLOR" "White" "FloorConnectors" "")
    (setvar "LWDISPLAY" 0)
    (command "-LAYER" "SET" "FloorConnectors" "")
    (command "-INSERT" "FloorConnector" pause 1.0 1.0 0) 
	
	(setq p (cdr (assoc 10 (entget (entlast)))))	
	(setq p-elevation (get-elevation p))
	(setq tags (get-floor-tags))
	
	(setq offset (floor-tag-offset p p-elevation tags))
	
	(insert-risers offset p-elevation tags)
)

; Get the x,y coordinates offset to the nearest floor tag, 
; the tag in the same elevation box.
(defun floor-tag-offset ( p p-elevation tags / tag tag-point tag-elevation offset )
	(foreach tag tags
		(progn
			(setq tag-point (get-ins-point tag))
			(setq tag-elevation (get-elevation (get-ins-point tag)))
			(if (= p-elevation tag-elevation)
				(setq offset (get-point-offset tag-point p))		
			)
		)
	)
	offset
)

; Insert corresponding risers (floor connectors)
(defun insert-risers ( offset p-elevation tags / tag-elevation tag-offset )
	(foreach tag tags
		(progn
			(setq tag-elevation (get-elevation (get-ins-point tag)))
			(if (not (= p-elevation tag-elevation))
				(progn 
					(setq tag-offset 
						(add-point-offset (get-ins-point tag) (- 0 (getx offset)) (- 0 (gety offset)))
					)
					(insert-flr-con tag-offset)
				)
			)
		)
	)
)

(defun invert-coordinates ( point )
	(list (- 0 (getx point)) (- 0 (gety point)))
)

(defun test-riser-tag-offset ( )
	(riser-tag-offset '(3906.95 1290.41 0.000000) (get-floor-tags))
)

; Group, using the lists-approx function, by the first item in a list of lists
(defun group-by ( items / item output offset offsets isinlist group)
	(setq offsets '())
	(foreach item items
		(progn 
			(setq isinlist (assoc-approx (car item) offsets 1.0))
			(if (not isinlist)
				(setq offsets (cons (list (car item)) offsets))
			)	
		)
	)
	(setq output '())
	(foreach offset offsets
		(setq group offset)
		(foreach item items	
			(if (lists-approx (car offset) (car item) 1.0)
				(setq group (append group (cdr item)))
			)
		)
		(setq output (cons group output))
	)
	output
)

; Riser offsets: Associative list with offsets of riser to floor tag 
; assocated with an entity name
; Example Output: (
;   ((-0.9906 56.8205 0) <Entity name: 354c6f30>) 
;   ((-2.3058 -56.2903 0) <Entity name: 354c70c0>)
;   ((-2.3058 -56.2903 0) <Entity name: 354c70e8>)
; )
(defun riser-offsets ( / tags riser-offset riser-name riser risers group groups)
	(setq groups '())
	(setq tags (get-floor-tags))	
	(setq risers (get-all-risers))	
	(foreach riser risers
		(progn
			(setq riser-point (get-ins-point riser))
			(setq riser-offset (floor-tag-offset riser-point tags))
			(setq group '())
			(setq group (append group (list riser-offset)))
			(setq riser-name (cdr (assoc -1 riser)))
			(setq group (append group (list riser-name)))
			(setq groups (cons group groups))
		)
	)
    groups
)

; Find x,y offset of riser from it's floor tag
; TODO: Optimize: Store these in a list?
(defun floor-tag-offset (riser-point tags / tag tag-offset tag-point riser-elevation)
	(setq riser-elevation (get-elevation riser-point))
	; TODO: Get offset, then check if any risers already have that same (approx)
	; offset, if yes then add to that riser/offset group;
	; if no then make a new offset group

	(foreach tag tags
		(setq tag-point (get-ins-point tag))
		; TODO: Optimize: get tag elevations and store them in a list?		
		(if (= riser-elevation (get-elevation tag-point)) ; Riser belongs to tag
			(setq tag-offset (get-point-offset riser-point tag-point))
		)
		tag-offset
	)

)
; TODO: Instead of find 'friends' of one point we need to group all risers 
(defun riser-friends ( p / p-elevation tag-elevation riser-elevation tag-point tag tags offset tag-offset riser risers friends)
	(princ "\nStarting riser-friends...\n")
	(setq friends '())
	(setq p-elevation (get-elevation p))
	(setq tags (get-floor-tags))
	(princ "\nTag Count: ")
	(princ (itoa (length tags)))
	
	(setq offset (floor-tag-offset p p-elevation tags))
	(princ "\nOffset: ")
	(princ offset)
	(princ "\nP-Elevation: ")
	(princ p-elevation)
	
	(princ "\nPoint: ")
	(princ p)
	
	(setq risers (get-all-risers))
	(princ "\nRiser Count: ")
	(princ (itoa (length risers)))
	(foreach tag tags
		(progn
			(setq tag-point (get-ins-point tag))
			(princ "\nTag Point: ")
			(princ tag-point)
			(setq tag-elevation (get-elevation  tag-point))
			;;;(if (not (= p-elevation tag-elevation))
			;;;	(progn 
					(foreach riser risers
						(setq riser-point (get-ins-point riser))
						; Which floor tag / elevation box does this riser-point belong to?
						; What is the offset from the floor tag that it belongs to?
						(princ "\n****Riser Point: ")
						(princ riser-point)
						(setq riser-elevation (get-elevation riser-point))
						(princ "\n    Riser Elevation: ")
						(princ riser-elevation)
						(if (= riser-elevation tag-elevation)
							(progn 
								(princ "\n    Elevation Match\n    Tag Offset: ")
								(setq tag-offset (get-point-offset riser-point tag-point))
								(princ tag-offset)
								(if (> 1.0 (distance tag-offset (invert-coordinates offset)))
									(progn 
										(princ "\n    Tag Offset MATCH\n")
										(setq friends (append riser friends))
									)
									(princ "\n    No tag offset match\n")
								)
							)
							(princ "\n    No elevation match\n")
						)
						;(setq tag-offset 
							;(add-point-offset riser-point (- 0 (getx offset)) (- 0 (gety offset)))
						;)
						;;(princ "\n    Tag Offset: ")
						;;(princ tag-offset)
						
						; Check distance of each tag-offset to original point 'p'
						; if distance is small then add each tag to a group
						; Assign a letter to each group chr(65 to 90)
						; A-Z (setq x 65)(while (<= x 90)(princ (chr x))(setq x (1+ x)))
						
						;;;(insert-flr-con tag-offset)
					)
			;;;	)
			;;;)
		)
	)
    friends
)

(defun insert-flr-con ( p / )
	(command "-LAYER" "NEW" "FloorConnectors" "")
    (command "-LAYER" "COLOR" "White" "FloorConnectors" "")
    (setvar "LWDISPLAY" 0)
    (command "-LAYER" "SET" "FloorConnectors" "")
    (command "-INSERT" "FloorConnector" p 1.0 1.0 0) 
)

(defun get-floor-tags ( / en ent tags layer) 
	(setq tags '())
	(setq en (entnext))
    (while en
	    (setq ent (entget en))
		(if (and (or (str= "FloorTags" (get-layer en))
					(str= "Floor Tags" (get-layer en))
				)
		        (str= "INSERT" (get-etype en))
				(str= "FloorTag" (get-block-name en))
			)
			(setq tags (cons ent tags))
		)
		(setq en (entnext en))
	)
	tags
)