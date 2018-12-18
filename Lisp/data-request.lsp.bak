; The request-data function uses two global system variables USERS1, and USERS2
; to request data from the LoopCAD.vbi module which passes the request to the 
; LoopCAD.dll module.  

(defun data-request (key)
    (setvar "USERS1" "RequestData")
	(setvar "USERS2" key)
	(vl-vbarun "Controller.RequestData")
	;Verify data was returned
	(if (= (getvar "USERS1") key)
		(getvar "USERS2")
		(alert (strcat "Error requesting data from LoopCAD.dll: Key: " key " Value: " (getvar "USERS2")))
	)
)