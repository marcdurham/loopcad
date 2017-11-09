(defun blocks-delete (layer-name block-name / block blocks)
    (setq blocks (blocks-on-layer (list layer-name) (list block-name)))
    (foreach block blocks
        (command)
    )
)