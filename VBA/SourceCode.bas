Attribute VB_Name = "SourceCode"
Public Sub SaveCodeModules()
'This code Exports all VBA modules
Dim i As Integer
Dim Name As String
    
    With VBE.ActiveVBProject
        For i% = 1 To .VBComponents.Count
            If .VBComponents(i%).CodeModule.CountOfLines > 0 Then
                Name = .VBComponents(i%).CodeModule.Name
                .VBComponents(i%).Export "C:\LoopCAD\VBA\" & Name & ".bas"
            End If
        Next i
    End With

End Sub

Public Sub RemoveCodeModules()
'Deletes all modules other than 'SourceCode' and 'ThisWorkspace'
Dim i As Integer
Dim Name As String
Dim moduleCount As Integer

    With VBE.ActiveVBProject
        moduleCount = .VBComponents.Count
        For i = moduleCount To 1 Step -1
            Name = .VBComponents(i).CodeModule.Name
            If Name <> "SourceCode" And Name <> "ThisWorkspace" Then
                .VBComponents.Remove .VBComponents(Name)
            End If
        Next i
    End With

End Sub
Public Sub ImportCodeModules()
Dim i As Integer
Dim Name As String
Dim d As Variant

    ChDir "C:\LoopCAD\"
    d = VBA.Dir(".\VBA\")
    Do
        Name = Left(d, Len(d) - 4)
        If LCase(Right(d, 4)) = ".bas" And d <> "SourceCode.bas" Then
            VBE.ActiveVBProject.VBComponents.Import "C:\LoopCAD\VBA\" & Name & ".bas"
        End If
        d = VBA.Dir()
    Loop While Len(d) > 0

End Sub

