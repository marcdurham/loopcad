Attribute VB_Name = "ReferenceDlls"
Option Explicit

Sub AddReference(Path As String, Name As String)
    Dim VBAEditor As VBIDE.VBE
    Dim vbProj As VBIDE.VBProject
    Dim chkRef As VBIDE.Reference
    Dim BoolExists As Boolean

    Set VBAEditor = Application.VBE
    'Set vbProj = VBIDE.VBProject
    Set vbProj = VBAEditor.ActiveVBProject

    For Each chkRef In vbProj.References
        If chkRef.Name = Name Then
            BoolExists = True
            GoTo CleanUp
        End If
    Next

    vbProj.References.AddFromFile Path
    'vbProj.References.AddFromFile "LoopCAD.dll"
    'vbProj.References.AddFromFile "FileConverter.dll"
    'vbProj.References.AddFromFile "LoopCalc.Core.dll"
    'vbProj.References.AddFromFile "LoopFlow.Common.dll"

CleanUp:
    If BoolExists = True Then
        MsgBox "Reference already exists"
    Else
        MsgBox "Reference Added Successfully"
    End If

    Set vbProj = Nothing
    Set VBAEditor = Nothing
End Sub
