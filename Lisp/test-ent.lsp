(defun layer-of (ent) (cdr (assoc 8 (entget ent))))

(defun vertices-of (ent)
    (foreach property (entget ent)
    ;(assoc 10 (entget ent))
        ;(princ " Property:")
        ;(princ (car property))
        (if (= 10 (car property))
            (princ property)
        )
    )
)

;Example: Walks through entities
(setq one_ent (entnext)) ; Gets name of firstentity.
(while one_ent
    (princ "Entity:")
    ;(princ (cdr (assoc 8 (entget one_ent))))
    (princ "Layer:")
    (setq e-layer (layer-of one_ent))
    (princ e-layer)
    (princ " : ")
    (princ (cdr (assoc 10 (entget one_ent))))
     (princ " : ")
    (princ (cdr (assoc 0 (entget one_ent))))
    (princ "\n")
    (if (= e-layer "Pipe")
        (progn
            (princ "YES: ")
            ;(princ (assoc 10 (entget one_ent)))
            ;(princ (layer-of one_ent))
            (vertices-of one_ent)
            (princ " :END: \n")
        )
    ;    (progn
    ;        (entget one_ent)
    ;        (princ "\n")
    ;    )
    )
    (setq one_ent (entnext one_ent))
) ; Value of one_ent is now nil.
