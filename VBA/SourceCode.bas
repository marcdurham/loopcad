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
    
            If name <> "SourceCode" And name <> "ThisWorkspace" Then
                .VBComponents.Remove .VBComponents(name)
                .VBComponents.Import "C:\LoopCAD\VBA\" & name & ".bas"
            End If
        Next i
    End With

End Sub
Public Sub ImportCodeModules2()
Dim i As Integer
Dim name As String
Dim d As Variant

    ChDir "C:\LoopCAD\"
    d = VBA.Dir(".\VBA\")
    Do While Len(d) > 0
        Debug.Print "File: " & d
        d = VBA.Dir()
        name = Left(d, Len(d) - 4)
        If LCase(Right(d, 4)) = ".bas" Then
            VBE.ActiveVBProject.VBComponents.Import "C:\LoopCAD\VBA\" & name & ".bas"
        End If
    Loop

End Sub

Public Sub TestSourceCode()
Dim d As Variant
    ChDir "C:\LoopCAD\"
    d = VBA.Dir(".\VBA\")
    Do While Len(d) > 0
        Debug.Print "File: " & d
        d = VBA.Dir()
    Loop
End Sub

