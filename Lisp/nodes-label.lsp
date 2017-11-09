; Node Labeler
; First get all entities on a layer Tees, Heads DXF=8
; List
; defun heads-all
; defun entities-List*
; defun blocks-all
; defun on-layer

; get all entity types of "INSERT" DXF=0 ?
; Make a List
; Extract insertion points, (10 123.456 233.351 0) DXF=10, into a List
; defun points-from (blocks)
; insert blocks, TeeLabel.dwg at each points
; defun tee/head-label-insert-at (point-list)
; defun block-insert (name point-list)  ; Ignores scale and rotation
; set attribute of each block, NUMBER to a T.1...T.n sequential NUMBER
; Check to see if a TeeLabel is already there
; Version 2?
; Check the T.number of existing tees & nodes
(defun nodes-label-all ()
    (tees-label (heads-label 1))
)
(defun heads-label (node-number-start)
    (blocks-delete "HeadLabels" "HeadLabel")
    (nodes-label  
        (list "Heads") 
        (list "Head12" "Head14" "Head16" "Head18" "Head20")
        "HeadLabel"
        "H."
        node-number-start
    )
)
(defun tees-label (node-number-start)
    (blocks-delete "TeeLabels" "TeeLabel")
    (nodes-label  
        (list "Tees") 
        (list "Tee")
        "TeeLabels"
        "T."
        node-number-start
    )
)
(defun nodes-label ( layers block-names label-layer label-prefix node-number / block blocks point points props p proplist head-number block-point)
    (setq blocks 
        (blocks-on-layer layers block-names)
    )
    (foreach block blocks
        (progn
            (setq point (assoc 10 block))
            (setq points (append points (list point)))
        )
    )    
    (command "-LAYER" "SET" label-layer "")            
    (foreach p points
        (progn
            (setq proplist (append props (list p)))
            (setq block-point (cdr p))
            (setq head-label (strcat label-prefix (itoa node-number)))
            (command "-INSERT" label-layer block-point "1" "1" "0" head-label)
            (setq node-number (+ 1 node-number))
        )
    )
    node-number
)

; DXF Data Examples 
;((-1 . <Entity name: 21c78e18>) (0 . "INSERT") (5 . "299") (330 . <Entity name: b2884e0>) (100 . "AcDbEntity") (67 . 0) (410 . "Model") (8 . "HeadLabels") (62 . 5) (100 . "AcDbBlockReference") (2 . "HeadLabel") (10 -3.3056 147.1697 0) (41 . 1) (42 . 1) (43 . 1) (50 . 0) (70 . 1) (71 . 1) (44 . 0) (45 . 0) (210 0 0 1)) 
;((-1 . <Entity name: 21c78760>) (0 . "ATTRIB") (5 . "2A2") (330 . <Entity name: 21c7c9a0>) (100 . "AcDbEntity") (67 . 0) (410 . "Model") (8 . "HeadLabels") (6 . "Pex-1") (62 . 0) (100 . "AcDbText") (10 76.5173 186.8484 0) (40 . 8) (1 . "H.0") (50 . 0) (41 . 1) (51 . 0) (7 . "ARIAL") (71 . 0) (72 . 0) (11 0 0 0) (210 0 0 1) (100 . "AcDbAttribute") (2 . "HEADNUMBER") (70 . 0) (280 . 0))
;((-1 . <Entity name: 21c78850>) (0 . "SEQEND") (5 . "2A3") (330 . <Entity name: 21c7c9a0>) (100 . "AcDbEntity") (67 . 0) (410 . "Model") (8 . "Heads") (-2 . <Entity name: 21c7c9a0>))