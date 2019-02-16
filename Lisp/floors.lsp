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

; Get the nearest floor tag, the one in the same elevation box
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
	
	(setq risers (get-all-risers))
	(princ "\nRiser Count: ")
	(princ (itoa (length risers)))
	(foreach tag tags
		(progn
			(setq tag-point (get-ins-point tag))
			(princ "\nTag Point: ")
			(princ tag-point)
			(setq tag-elevation (get-elevation  tag-point))
			(if (not (= p-elevation tag-elevation))
				(progn 
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
									(princ "\n    Tag Offset MATCH\n")
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
				)
			)
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