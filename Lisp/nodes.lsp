(defun get-all-heads ()
	(get-blocks (list "HEADS"))
)

(defun get-all-tees ()
	(get-blocks (list "TEES"))
)

(defun get-all-risers ()
	(get-blocks (list "Floor Connectors" "FloorConnectors" "Risers"))
)
