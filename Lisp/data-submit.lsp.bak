; The data-submit function uses two global system variables USERS1, and USERS2
; to request data from the LoopCAD.vbi module which passes the submission to the 
; LoopCAD.dll module.  

(defun data-submit (key val)
    (setvar "USERS1" key)
	(setvar "USERS2" val)
	(vl-vbarun "Controller.SubmitData")
	; Verify data submission was received
	(if (not 
			(and
				(= (getvar "USERS1") "DataSubmissionReceived")
				(= (getvar "USERS2") key)
			)
		)
		(alert (strcat "Error submitting data to LoopCAD.dll: Key: " key " Value: " val))
	)	
)

(defun send-controller-command (command)
    (data-submit "ControllerCommand" command)
)