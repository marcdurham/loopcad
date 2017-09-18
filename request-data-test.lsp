; Load head defaults
;(submit-data "DefaultHeadCoverage" "12")
(submit-data "DefaultHeadModel" "TEST")
;(submit-data "DefaultHeadSlope" "1")
;(submit-data "DefaultHeadTemperature" "190")

; Show error if cannot link to LoopCAD.dll
(if (= "TEST" (request-data "DefaultHeadModel"))
	(prompt "\nLinked to LoopCAD.dll successfully.\n")
	(progn 
	    (alert (strcat "DefaultHeadModel: " (request-data "DefaultHeadModel")))
		(setq link-message  "ERROR: Link to LoopCAD.dll failed!")
		(prompt (strcat "\n" link-message "\n"))
		(alert link-message)
	)
)