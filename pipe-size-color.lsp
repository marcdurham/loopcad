(defun pipe-size-color (size)
	(if (= size "1/2") "Green"
	(if (= size "3/4") "Red"
	(if (= size "1") "Blue"
	(if (= size "1-1/4") "Magenta"))))
)