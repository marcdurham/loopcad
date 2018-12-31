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
	(if (or (not size) (= size "?"))
	    (progn
		    (princ "Pipe Sizes Available: 1/2, 3/4, 1, 1-1/4")
			(setq size (getstring "\nEnter Pipe Size:"))
		)
	)
	(setvar "OSMODE" osmode-snap-ins-pts)
	(setvar "ORTHOMODE" 1)
	(command "-LAYER" "NEW" "Pipe" "")
	(command "-LAYER" "COLOR" "White" "Pipe" "")
	(command "-LAYER" "SET" "Pipe" "")
	(command "-COLOR" (pipe-size-color size))
	(setq line-width "2\"")
	(prompt (strcat "\nPipe Size: " size "\n"))
	(prompt "\nDraw pipe to each head.\n")
	(command "-PLINE" pause "Width" line-width line-width pause)
)

; 64 = OSMODE: Snap to insertion points
(setq osmode-snap-ins-pts 64)

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

