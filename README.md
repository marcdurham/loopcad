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

# AutoCAD DXF Reference
[https://images.autodesk.com/adsk/files/autocad_2012_pdf_dxf-reference_enu.pdf]

# Handy LISP Functions for the Future
 (vla-startundomark (setq doc (vla-get-activedocument (vlax-get-acad-object))))
 (vla-endundomark doc)