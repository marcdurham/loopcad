; Load head defaults
;(data-submit "DefaultHeadCoverage" "12")
(data-submit "DefaultHeadModel" "TEST")
;(data-submit "DefaultHeadSlope" "1")
;(data-submit "DefaultHeadTemperature" "190")

; Show error if cannot link to LoopCAD.dll
(if (= "TEST" (data-request "DefaultHeadModel"))
	(prompt "\nLinked to LoopCAD.dll successfully.\n")
	(progn 
	    (alert (strcat "DefaultHeadModel: " (data-request "DefaultHeadModel")))
		(setq link-message  "ERROR: Link to LoopCAD.dll failed!")
		(prompt (strcat "\n" link-message "\n"))
		(alert link-message)
	)
)