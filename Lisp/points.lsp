
(defun add-point-offset (point x y)
    (list (+ x (getx point)) (+ y (gety point)) 0)
)
 
(defun get-point-offset (a b)
    (list (- (getx a) (getx b)) (- (gety a) (gety b)) 0)
)