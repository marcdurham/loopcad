; These exist, replace them
(defun greatest (a b)
  (if (> a b) a b)
)
(defun least (a b)
  (if (< a b) a b)
)

(vl-registry-read "HKEY_CURRENT_USER\\Software\\LoopCalc\\ProgeCAD" "Test")

(defun get-pipes ()
	(prompt "Getting pipes...")
	(entget (tblobjname "LAYER" "Pipe"))
	(princ)
)