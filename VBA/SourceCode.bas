Attribute VB_Name = "SourceCode"
'***********************************************
'**** How to use this SourceCode.bas module ****
'**** ******************************************
' 1. Create a new module/project in VBA Editor
' 2. Import the SourceCode.bas file
' 3. Run the ImportCodeModules() method
'
' You do not need to use the RemoveCodeModules when
' you use SourceCode.bas
'
Public Sub SaveCodeModules()
'This code Exports all VBA modules
Dim i As Integer
Dim moduleName As String
    
    With VBE.ActiveVBProject
        For i = 1 To .VBComponents.Count
            If .VBComponents(i).CodeModule.CountOfLines > 0 Then
                moduleName = .VBComponents(i).CodeModule.Name
                .VBComponents(i).Export "C:\LoopCAD\VBA\" & moduleName & ".bas"
            End If
        Next i
    End With

End Sub

Public Sub RemoveCodeModules()
'Deletes all modules other than 'SourceCode' and 'ThisWorkspace'
Dim i As Integer
Dim moduleName As String
Dim moduleCount As Integer
    
    If MsgBox("Are you sure you want to delete all VBA modules?", vbYesNoCancel) = vbYes Then
        With VBE.ActiveVBProject
            moduleCount = .VBComponents.Count
            For i = moduleCount To 1 Step -1
                moduleName = .VBComponents(i).CodeModule.Name
                If moduleName <> "SourceCode" And moduleName <> "ThisWorkspace" Then
                    .VBComponents.Remove .VBComponents(moduleName)
                End If
            Next i
        End With
    End If

End Sub
Public Sub ImportCodeModules()
'Remove existing modules first
Dim i As Integer
Dim moduleName As String
Dim fileName As Variant

    ChDir "C:\LoopCAD\"
    fileName = VBA.Dir(".\VBA\")
    Do
        moduleName = Left(fileName, Len(fileName) - 4)
        If LCase(Right(fileName, 4)) = ".bas" And fileName <> "SourceCode.bas" Then
            VBE.ActiveVBProject.VBComponents.Import "C:\LoopCAD\VBA\" & moduleName & ".bas"
        End If
        fileName = VBA.Dir()
    Loop While Len(fileName) > 0

End Sub

