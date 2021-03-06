; Splits a string at all occurences of the delim into a list
; Example: (string-split "My,name,is,Marc") = (list "My" "name" "is" "Marc")
(defun string-split ( str delim / pos )
    (if (setq pos (vl-string-search delim str))
      (cons (substr str 1 pos) (string-split (substr str (+ pos 1 (strlen delim))) delim))
      (list str)
    )
)
