Attribute VB_Name = "Controller"
Dim Proc As LoopCAD.CommandProcessor
Public Sub RequestData()
    Processor().RequestData
End Sub
Public Sub SubmitData()
    Processor().SubmitData
End Sub
Private Function Processor() As LoopCAD.CommandProcessor

    If Proc Is Nothing Then
        Set Proc = New CommandProcessor
    End If
    
    If Proc.Application Is Nothing Then
        Set Proc.Application = Application
    End If
    
    Set Processor = Proc
    
End Function
Public Sub ReferenceAllDlls()
MsgBox "Referencing all Dlls..."
    AddReference "LoopCAD.dll", "LoopCAD"
    AddReference "FileConverter.dll", "FileConverter"
    AddReference "LoopCalc.Core.dll", "LoopCalc_Core"
    AddReference "LoopFlow.Common.dll", "LoopFlow_Common"
End Sub
Public Sub Test()
   Debug.Print "Test"
End Sub

