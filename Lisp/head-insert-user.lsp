(defun head-insert-user ( / tmp) 
	(if (null global:head-model)
		(setq global:head-model "RFC43")
	)
	(if (setq tmp (getstring (strcat "\nHead Model <" global:head-model ">: ")))
		(setq global:head-model tmp)
	)
;
;   This section is for heads that you already know the coverage
;   Now the default is that you don't know. 
;   See the head-insert.lsp file and the head-insert function. 
;
;   (if (null global:head-coverage)
;		(setq global:head-coverage "16")
;	)
;	(initget "12 14 16 18 20")
;	(if (setq tmp (getkword (strcat "\nHead Coverage [12/14/16/18/20] <" global:head-coverage ">: ")))
;		(setq global:head-coverage tmp)
;	)
;
	(if (null global:head-slope)
		(setq global:head-slope "")
	)
	(if (setq tmp (getstring (strcat "\nHead Slope <" global:head-slope ">: ")))
		(setq global:head-slope tmp)
	)
	(if (null global:head-temperature)
		(setq global:head-temperature "")
	)
	(if (setq tmp (getstring (strcat "\nHead Temperature <" global:head-temperature ">: ")))
		(setq global:head-temperature tmp)
	)
    (head-insert
		global:head-model
		"20" ; global:head-coverage
		global:head-slope
		global:head-temperature
	)
)