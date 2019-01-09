Attribute VB_Name = "SourceCode"
Public Sub SaveCodeModules()
'This code Exports all VBA modules
Dim i As Integer
Dim name As String
    
    With VBE.ActiveVBProject
        For i% = 1 To .VBComponents.Count
            If .VBComponents(i%).CodeModule.CountOfLines > 0 Then
                name = .VBComponents(i%).CodeModule.name
                .VBComponents(i%).Export "C:\LoopCAD\VBA\" & name & ".bas"
            End If
        Next i
    End With

End Sub

Public Sub ImportCodeModules()

    With VBE.ActiveVBProject
        For i = 1 To .VBComponents.Count
    
            ModuleName = .VBComponents(i).CodeModule.name
    
            If ModuleName <> "VersionControl" Then
                If Right(ModuleName, 6) = "Macros" Then
                    .VBComponents.Remove .VBComponents(ModuleName)
                    .VBComponents.Import "C:\LoopCAD\VBA\" & ModuleName & ".bas"
               End If
            End If
        Next i
    End With

End Sub
