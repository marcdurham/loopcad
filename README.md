# LoopCAD 
LoopCAD is a set of commands for AutoCAD to assist drawing a looped PEX fire 
sprinkler system.

# Commands
- HEAD
- PIPE
- TEE
- DOMESTIC-TEE
- ELEVATION-BOX
- FLOOR-TAG
- JOB-DATA

# Handy LISP Functions for the Future
 (vla-startundomark (setq doc (vla-get-activedocument (vlax-get-acad-object))))
 (vla-endundomark doc)