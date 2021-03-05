(defun elevation-box-draw (/ a b p1 p2 p3 p4 top bottom left right points)
    (setq old-osmode (getvar "OSMODE"))
    (setq temperror *error*)
    (defun *error* (message)
        (princ)
        (princ "Error")
        (princ message)
        (princ)
        (setvar "OSMODE" old-osmode)
        (command-s "-COLOR" "BYLAYER")
        (command-s "-LAYER" "SET" "0" "")
        (setq *error* temperror)
    )
    (setvar "OSMODE" 0)
    (setq a (getpoint "\nElevation box first corner:"))
    (setq b (getcorner a))
        (if (null elevation) 
        (setq elevation 100)
        (princ (strcat "Default elevation set to " (itoa elevation)))
    )
    (setq elevation (getint (strcat "\nEnter elevation (ft): <" (itoa elevation) ">")))
    
    (setq right (greatest (car a) (car b)))
    (setq top  (greatest (cadr a) (cadr b)))
    (setq left (least (car a) (car b)))
    (setq bottom (least (cadr a) (cadr b)))
    ; (setq p1 (list left top))
    ; (setq p2 (list right top))
    ; (setq p3 (list right bottom))
    ; (setq p4 (list left bottom))
 
    (command "-COLOR" "BYLAYER" "")
    (command "-LINETYPE" "SET" "Continuous" "")
    (command "-LAYER" "NEW" "ElevationBox" "")
    (command "-LAYER" "COLOR" "Magenta" "ElevationBox" "")
    (command "-LAYER" "SET" "ElevationBox" "")
    ;(command "-PLINE" p1 p2 p3 p4 p1 "")
    ;(command "-MTEXT" p1 p3 (strcat "Elevation " (itoa elevation)) "") 
    
    ;; Creates a lightweight polyline in model space.
    (setq acadObj (vlax-get-acad-object))
    (setq doc (vla-get-ActiveDocument acadObj))

    ;; Define the 2D polyline points
    (setq points (vlax-make-safearray vlax-vbDouble '(0 . 9)))
    (vlax-safearray-fill 
      points 
      (list 
        left top
        right top
        right bottom
        left bottom
        left top
      )
    )
        
    ;; Create a lightweight Polyline object in model space
    (setq modelSpace (vla-get-ModelSpace doc))
    (setq plineObj (vla-AddLightWeightPolyline modelSpace points))
    (vla-put-layer plineObj "ElevationBox")
   
    ;; Define the mtext object
    (setq corner (vlax-3d-point a)
          width (abs (- right left))
          text (strcat "Elevation " (itoa elevation))
    )

    ;; Creates the mtext object
    (setq modelSpace (vla-get-ModelSpace doc))
    (setq MTextObj (vla-AddMText modelSpace corner width text))
    (vla-put-height MTextObj 10.0)
    (vla-put-layer MTextObj "ElevationBox")

    ; Set things back
    (command "-COLOR" "BYLAYER")
    (command "-LAYER" "SET" "0" "")
    (setvar "OSMODE" old-osmode)
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
    (princ "\nGetting entity name...\n")
    (setq en (get-ename ent))
    (princ (strcat "\n Entity Name: " en))
    (setq en (entnext en))
   (princ (strcat "\n Next Entity Name: " en))
    (setq ent (entget en))
    (while en
        (cond ((str= "VERTEX" (get-etype ent))
                (setq ent (entget en))
                (setq vertex (get-ins-point ent))
                (setq vertices (cons vertex vertices))
                (setq en (entnext en))
                (setq ent (entget en))
            )
            (T (setq en nil))
        )
    )
    vertices
)

(defun test-get-elevation ( / )
    ;(princ "\nTesting: test-ebox\n")
    ; TODO: Add some elevation boxes, 2 at least
    (princ "\nShould return 102: ")
    (princ (get-elevation (list 4342.29 1633.89 0.000000)))
    (princ "\nShould return 109: ")
    (princ (get-elevation (list 4224.10 1672.70 0.000000)))
    (princ "\n")
)

; Find the smallets elevation box point 'p' is in, 
; return the elevation.
(defun get-elevation ( p / box boxes a b i ar in-areas all-areas m vertex vertices text-box text-boxes smallest-box elevation)
    (setq elevation "0")
    ; Get areas of all boxes that p is in, it may be in more than one
    (setq in-areas '())
  
    (princ "\nScan boxes\n")
  
    (setq boxes (get-elevation-boxes))
    
    (princ (strcat "\nElevation Boxes Found: " (itoa (length boxes)) "\n"))
  
    (foreach box boxes
        (progn 
            (setq a (car (corners box)))
            (setq b (cadr (corners box)))
            (setq ar (area a b))
            (setq all-areas (append all-areas (list ar)))    
            (if (in-ebox p box)
                (setq in-areas (append in-areas (list ar)))
            )
        )
    )
  
    (princ "\nBoxes done\n")
  
    (setq m (apply 'min in-areas))
    (setq i (index-of m all-areas))
    (setq smallest-box (nth i boxes))
    
    (princ "\nSmallest box found\n")
  
    (if (not (null smallest-box))
        (progn
            ; Match the smallest (elevation) box to it's MText containing the elevation text
            (setq vertices (get-vertices smallest-box))
            (princ (strcat "\nVertices found: " (itoa (length vertices)) "\n"))
            (setq text-boxes (get-elevation-text))
            (princ (strcat "\nText Boxes Count: " (itoa (length text-boxes)) "\n"))
            (foreach vertex vertices
                (foreach text-box text-boxes
                    
                  (setq insPoint (get-ins-point text-box))
                  (setq insPoint2D (list (nth 0 insPoint) (nth 1 insPoint)))
                  (princ (strcat "\nVertex length: " (itoa (length vertex))  (rtos (nth 0 vertex))  (rtos (nth 1 vertex)) "\n"))
                    (princ (strcat "\nInspoint2d length " (itoa (length insPoint2D))  (rtos (nth 0 insPoint2D))  (rtos (nth 1 insPoint2D))"\n"))
                    (princ (strcat "\nDistance: " (rtos (distance vertex insPoint)) "\n"))
                    (if (< (distance vertex insPoint) near-line-margin)
                        ; This text-box belongs to this elevation box
                        (setq elevation (elevation-from text-box))                                     
                    )
                )
            )
        )
    )
    (atoi elevation)
)

; Get numeric elevation value from MText entity
(defun elevation-from ( text-box / )
    ; Input Example: "Elevation 999"
    ; Digits start at position 11
    (substr (text-from text-box) 11)
)

; TODO: Try vla-get-text
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
    (setq vertices (get-vertices rectangle))
    (setq a (nth 0 vertices)) ; First corner
    (setq b (nth 2 vertices)) ; Opposite corner    
    (list a b )
)