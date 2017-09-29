(defun contains (target list / output result)
    (setq result nil)
    (foreach item list
        (if (eq item target)
            (setq result T)
        )
    )
    result
)