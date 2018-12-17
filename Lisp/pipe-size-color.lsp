(defun pipe-size-color (size)
	(if (= size "1/2") color-green
	(if (= size "3/4") color-red
	(if (= size "1") color-blue
	(if (= size "1-1/4") color-magenta))))
)

(setq color-red 1)
(setq color-yellow 2)
(setq color-green 3)
(setq color-cyan 4)
(setq color-blue 5)
(setq color-magenta 6)
(setq color-black-white 7)
(setq color-dark-gray 8)
(setq color-light-gray 9)

(setq color-list 
	(list 
		(cons "1/2" color-green)
		(cons "3/4" color-red)
		(cons "1" color-blue)
		(cons "1-1/4" color-magenta)
	)
)