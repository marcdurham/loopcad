VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkspace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workspace_EndOpen(ByVal DrawingName As String)
    Application.ActiveDocument.Utility.Prompt "Document open detected, loading LoopCAD LISP modules..."
    Application.LoadLISP "LoopCAD.lsp"
End Sub


