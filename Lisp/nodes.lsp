(defun get-all-heads ()
    (get-blocks (list "Heads"))
)

(defun get-all-tees ()
    (get-blocks (list "Tees"))
)

(defun get-all-domestic-tees ()
    (get-blocks (list "DomesticTees"))
)

(defun get-all-risers ()
    (get-blocks (list "Floor Connectors" "FloorConnectors" "Risers"))
)
