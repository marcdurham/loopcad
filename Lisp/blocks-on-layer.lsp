(defun blocks-on-layer 
    (   
        layers
        block-names 
        / entity         
        blocks
    )
    (foreach entity (entities-all)
        (progn
            (princ "\nLoop.Type:")
            (princ (type entity))
            (princ "\nLBlock.Type:")
            (princ (entity-type entity))
            (princ "\nBlock.Layer:")
            (princ (entity-layer entity))
            (princ "\nBlock.Name:")
            (princ (entity-name entity))
            ;(setq blocks (append blocks (list entity)))     
            (if     
                (and 
                    (contains (entity-layer entity) layers)
                    (= (entity-type entity) "INSERT")
                    (contains (entity-name entity) block-names)
                )
                (setq blocks (append blocks (list entity)))
            )
        )
    )
    blocks
)

(defun entity-name (entity)
    (cdr (assoc 2 entity))
)

(defun entity-type (entity)
    (cdr (assoc 0 entity))
)

(defun entity-layer (entity)
    (cdr (assoc 8 entity))
)

(defun entities-all ( / e entity entities)
    (setq entity (entnext))
    (while entity
        (setq e (entget entity))
        (setq entities (append entities (list e)))
        (setq entity (entnext entity))
    )
    entities
)
