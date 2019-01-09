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
Dim i As Integer
Dim name As String

    With VBE.ActiveVBProject
        For i = 1 To .VBComponents.Count
    
            name = .VBComponents(i).CodeModule.name
    
            If ModuleName <> "VersionControl" Then
                If Right(ModuleName, 6) = "Macros" Then
                    .VBComponents.Remove .VBComponents(name)
                    .VBComponents.Import "C:\LoopCAD\VBA\" & name & ".bas"
               End If
            End If
        Next i
    End With

End Sub
