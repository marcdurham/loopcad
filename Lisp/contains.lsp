(defun contains (target list / output result)
    (setq result nil)
    (foreach item list
        (if (eq (strcase item) (strcase target))
            (setq result T)
        )
    )
    result
)