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
- RISER
- JOB-DATA

# How to Install
- Click the AutoCAD logo button (farthest upper left)
- Click **Options**
- Click **Files** tab
- File Support Search Path
- Add LoopCAD folder, ex: `C:\LoopCAD`
- Add blocks folder, ex: `C:\LoopCAD\Blocks`

# AutoCAD Auto Start Help
(https://knowledge.autodesk.com/support/autocad/learn-explore/caas/sfdcarticles/sfdcarticles/Automatically-load-AutoLISP-routines.html)

# AutoCAD DXF Reference
[https://images.autodesk.com/adsk/files/autocad_2012_pdf_dxf-reference_enu.pdf]

# Handy LISP Functions for the Future
 (vla-startundomark (setq doc (vla-get-activedocument (vlax-get-acad-object))))
 (vla-endundomark doc)

# OSMODE
- 0 NONe
- 1 ENDpoint
- 2 MIDpoint
- 4 CENter
- 8 NODe
- 16 QUAdrant
- 32 INTersection
- 64 INSertion
- 128 PERpendicular
- 256 TANgent
- 512 NEArest
- 1024 Geometric CEnter
- 2048 APParent Intersection
- 4096 EXTension
- 8192 PARallel
- 16384 Suppresses the current running object snaps