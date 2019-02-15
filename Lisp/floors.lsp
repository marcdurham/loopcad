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
						  ; should be
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

; Get the nearest, in the same elevation box, floor tag
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