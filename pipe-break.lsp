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