; TODO: Finish this, it does not work yet.
; Select all PolyLines on the Pipes layer and then break them
(defun c:pipe-break ( / pipes)
    ;(setq lay_name "Pipe")
    (setq pipes
        (ssget "X"
            (list (cons 8 "Pipe"))
        )
    )
    ;(foreach var pipes
       ; (progn (princ "Item"))
    ;)
    (car pipes)
)

;This returns all data about the first entity
(entget (entnext))

; Gets the first item that starts with 8, the key & value
(assoc 8 entdata)

; Gets just the last part of the item, the value, 8 is the layer
(cdr (assoc 8 entdata))

; Gets the point of the last entity, 10 is the "insertion point"
(cdr (assoc 10 (entget (entlast))))

; Layer-of function
(defun layer-of (ent) (cdr (assoc 8 (entget (ent)))))

(entmake) ;make entity, add to drawing
(entmod) ;modify entity
(entupd) ;Redraw entity

;(subst a b) ;Like a search and replace in a list

;Example: Walks through entityes
(setq one_ent (entnext)) ; Gets name of first
entity.
(while one_ent
.
. ; Processes new entity.
.
(setq one_ent (entnext one_ent))
) ; Value of one_ent is
now nil.