; (pipe-draw) and (pipe-make) methods, a little different from each other.
; (pipe-draw) is the PIPE-DRAW command called by the user, 
;      by clicking a button, to pipe a system.
; (pipe-make) is used by the BREAK-PIPES command to re-draw the pipes.

(defun pipe-draw (size / old-osmode old-orthomode line-width)
	(setq old-osmode (getvar "OSMODE"))
	(setq old-orthomode (getvar "ORTHOMODE"))
	(defun *error* (message)
	  (princ)
	  (princ message)
	  (princ)
	  (setvar "OSMODE" old-osmode)
	  (setvar "ORTHOMODE" old-orthomode)
	  (command "-COLOR" "BYLAYER")
	  (command "-LAYER" "SET" "0" "")
	)
	(progn
		(if (null global:pipe-size)
			(setq global:pipe-size "05")
		)
		(initget "05 075 1 15")
		(if (setq tmp (getkword (strcat "\nPipe Size [05/075/1/15] <" global:pipe-size ">: ")))
			(setq global:pipe-size tmp)
		)
	)
	(setvar "OSMODE" osmode-snap-ins-pts)
	(setvar "ORTHOMODE" 1)
	(command "-LAYER" "NEW" layer-pipe "")
	(command "-LAYER" "COLOR" "White" layer-pipe "")
	(command "-LAYER" "SET" layer-pipe "")
	(command "-COLOR" (pipe-size-color global:pipe-size))
	(setq line-width "2\"")
	(prompt (strcat "\nPipe Size: " global:pipe-size "\n"))
	(prompt "\nDraw pipe to each head.\n")
	(command "PLINE" pause "Width" line-width line-width pause)
)

; 64 = OSMODE: Snap to insertion points
(setq osmode-snap-ins-pts 64)
; Layer for pipes is "Pipe" not "Pipes"
(setq layer-pipe "Pipe")

(setq pipe-width 2.0)

; Visual test, look at the output.
(defun test-make-pipe ()
    (make-pipe "1-1/4" 
		(list
			(10410.5 15273.4 0.000000) 
			(10386.9 15294.1 0.000000)
			(10797.7 15181.8 0.000000)
			(10856.8 15409.3 0.000000)
		)
	)
)

; Returns the pipe size, ex: "3/4" from a color input of 1 (which is red)
(defun get-pipe-size (polyline / vertex remaining)
	;(pipe-color-size (cdr (assoc 62 polyline)))
	(pipe-color-size (ent-color polyline))
)

(defun make-pipe (size vertices)
    (make-polyline vertices  (pipe-size-color size) "Pipes")
)

; Make a vertices for a polyline
; Also needs an 'SEQEND' entry
(defun make-vertex (p)
	(entmake 
		(list 
			(cons 0 "VERTEX") 
			(cons 10 p)
		)
	)			
)

; Name colors
(setq color-red 1)
(setq color-yellow 2)
(setq color-green 3)
(setq color-cyan 4)
(setq color-blue 5)
(setq color-magenta 6)
(setq color-black-white 7)
(setq color-dark-gray 8)
(setq color-light-gray 9)
(setq color-orange 30)

(setq size-color-list 
	(list 
		(cons "1/2" color-green) ; Domestic pipe
		(cons "3/4" color-red)
		(cons "1" 150) ; Since this is listed first it is the default 1 inch color
		(cons "1" color-blue)
		(cons "1" color-orange) ; Copper pipe
		(cons "1-1/2" color-magenta)
		(cons "05" color-green)
		(cons "075" color-red)
		(cons "15" color-magenta)
	)
)

(defun pipe-color-size (color / size)
	(assoc-cdr color size-color-list)
)

(defun pipe-size-color (size)
	(cdr (assoc size size-color-list))
)

; Return the first element from a list of dotted pairs by matching
; the second element.
; Like assoc but find the value by the second (cadr) value 
; insead of the first (car)
(defun assoc-cdr (val items / key)
	(foreach item items
	    (if (= val (cdr item))
			(setq key (car item))
		)
	)
	key
)