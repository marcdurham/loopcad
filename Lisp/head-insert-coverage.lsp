(defun head-insert-coverage (coverage) 
    (head-insert 
	    (data-request "DefaultHeadModel")
		coverage
		(data-request "DefaultHeadSlope")
		(data-request "DefaultHeadTemperature")
	)
)