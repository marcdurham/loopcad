
(defun floor-tag-insert ( / p old-osmode elevation floor-name)
    (setq old-osmode (getvar "OSMODE"))
    (defun *error* (message)
        (princ)
        (princ message)
        (princ)
        (setvar "OSMODE" old-osmode)
        (setvar "LWDISPLAY" 1)
    )
    
    (setvar "INSUNITS" 1) ; 0 = not set, 1 = inches, 2 = feet
    (setvar "OSMODE" osmode-snap-ins-pts)
    (command "-LAYER" "NEW" "Floor Tags" "")
    (command "-LAYER" "COLOR" "Cyan" "Floor Tags" "")
    (setvar "LWDISPLAY" 0)
    (command "-LAYER" "SET" "Floor Tags" "")
    (setq p (getpoint "Click insertion point for floor tag"))
    (setq floor-name (getstring T "Enter floor name")) ; getstring with T allows spaces
    (setq elevation (get-elevation p))
    (princ (strcat "Elevation: " (itoa elevation)))
  
    (setq acadObj (vlax-get-acad-object))
    (setq doc (vla-get-ActiveDocument acadObj))
    
    ; Insert the block
    (setq insertionPoint (vlax-3d-point p))
    (setq modelSpace (vla-get-ModelSpace doc))
    (setq block (vla-InsertBlock modelSpace insertionPoint "FloorTag" 1 1 1 0))
  
    ; get the block attributes
    (setq attributes (vlax-safearray->list (vlax-variant-value (vla-getAttributes block))))
    
    ; Set attribute values by the attribute position
    (vla-put-TextString (nth 0 attributes) floor-name)
    (vla-put-TextString (nth 1 attributes) elevation)
)

(defun riser-insert ( / p w old-osmode tag tags tag-elevation p-elevation offset tag-offset)
    (setq old-osmode (getvar "OSMODE"))
    (defun *error* (message)
        (princ)
        (princ message)
        (princ)
        (setvar "OSMODE" old-osmode)
        (setvar "LWDISPLAY" 1)
    )
    (setvar "INSUNITS" 1) ; 0 = not defined 1 = inches 2 = feet
    (setvar "OSMODE" osmode-snap-ins-pts)
    (command "-LAYER" "NEW" "Risers" "")
    (command "-LAYER" "COLOR" "White" "Risers" "")
    (setvar "LWDISPLAY" 0)
    (command "-LAYER" "SET" "Risers" "")
    (command "-INSERT" "FloorConnector" pause 1.0 1.0 0)
    
    (setq p (cdr (assoc 10 (entget (entlast))))) ; Get insertion point
    (setq p-elevation (get-elevation p))
    (setq tags (get-floor-tags))
    
    (setq offset (floor-tag-elevation-offset p p-elevation tags))
    
    (insert-risers offset p-elevation tags)
    (princ "\nInsert Riser: Done\n")
    (princ)
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
                    (command "-INSERT" "Riser" tag-offset 1.0 1.0 0)
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

(defun riser-labels-create ( last-i / offsets riser-groups r riser risers  riser-entity i leter letters )
    (princ "\nScanning risers and creating labels...\n")
    (setq letters (get-alphabet))
    (setq risers (get-all-risers))
    (setq offsets (riser-offsets))
    (setq riser-groups (group-by offsets))
    (setq i 0)
    (setq r 0)
    (while (< i (length riser-groups))
        (setq letter (nth i letters))
        (setq riser-group (nth i riser-groups))
        (foreach riser (cdr riser-group)
            (setq riser-entity (entget riser))
            (insert-riser-label 
                (get-ins-point riser-entity)
                (strcat "R." (itoa (+ r last-i)) "." letter)
            )
            (setq r (1+ r))
        )
        (setq i (1+ i))
    )
    (+ r last-i)
)

(defun get-alphabet ( / i letters)
    (setq letters '())
    (setq i 65)
    (while (<= i 90)
        (setq letters (append letters (list (chr i))))
        (setq i (1+ i))
    )
    letters
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
            (setq riser-offset (floor-tags-offset riser-point tags))
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
(defun floor-tags-offset (riser-point tags / tag tag-offset tag-point riser-elevation)
    (setq riser-elevation (get-elevation riser-point))
    (foreach tag tags
        (setq tag-point (get-ins-point tag))
        ; TODO: Optimize: get tag elevations and store them in a list?        
        (if (= riser-elevation (get-elevation tag-point)) ; Riser belongs to tag
            (setq tag-offset (get-point-offset riser-point tag-point))
        )
    )
    tag-offset
)

; Get the x,y coordinates offset to the nearest floor tag, 
; the tag in the same elevation box.
; Almost the same as floor-tags-offset, but elevation is input
(defun floor-tag-elevation-offset ( p p-elevation tags / tag tag-point tag-elevation offset )
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