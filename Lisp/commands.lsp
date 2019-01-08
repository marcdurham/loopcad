(defun C:SET-HEAD-DATA () (head-data-set))
(defun C:SET-HEAD-MODEL () (head-model-set))
(defun C:SET-HEAD-COVERAGE () (head-coverage-set))
(defun C:SET-HEAD-SLOPE () (head-slope-set))
(defun C:SET-HEAD-TEMP () (head-temperature-set))
(defun C:HEAD () (head-insert-user))
(defun C:HEAD-12 () (head-insert-coverage "12"))
(defun C:HEAD-14 () (head-insert-coverage "14"))
(defun C:HEAD-16 () (head-insert-coverage "16"))
(defun C:HEAD-18 () (head-insert-coverage "18"))
(defun C:HEAD-20 () (head-insert-coverage "20"))
(defun C:PIPE () (pipe-draw "?"))
(defun C:PIPE-12 () (pipe-draw "1/2"))
(defun C:PIPE-34 () (pipe-draw "3/4"))
(defun C:PIPE-1 () (pipe-draw "1"))
(defun C:PIPE-114 () (pipe-draw "1-1/4"))
(defun C:ELEVATION-BOX () (elevation-box-draw))
(defun C:TEE () (tee-insert))
(defun C:DTEE () (domestic-tee-insert))
(defun C:LABEL-NODES () (label-all-nodes))
(defun C:LABEL-PIPES () (label-all-pipes))
(defun C:BREAK-PIPES () (break-pipes-delete-old))
(defun C:JOIN-HEADS () (head-join))
(defun C:FLOOR-CONNECTOR () (floor-connector-insert))
(defun C:FLOOR-TAG () (floor-tag-insert))
(defun C:INSERT-JOB-DATA () (alert "This must be done manually using the JobData block."))

(princ) ; exit quietly
