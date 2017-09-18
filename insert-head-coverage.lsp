(defun insert-head-coverage (coverage) 
    (insert-head 
	    (request-data "DefaultHeadModel")
		coverage
		(request-data "DefaultHeadSlope")
		(request-data "DefaultHeadTemperature")
	)
)