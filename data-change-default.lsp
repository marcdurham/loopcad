(defun data-change-default (key message / value old-value)
	(setq old-value (data-request key))
	(setq value (getstring (strcat "\n" message " <" old-value "> :")))
	(if (not (= value ""))
		(data-submit key value)
	)
)