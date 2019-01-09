VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHeadModelBars 
   Caption         =   "Load Tool Bar for Head Models"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   OleObjectBlob   =   "frmHeadModelBars.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHeadModelBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





'Public lHeadCoverage(1 To 3) As String
'Public lHeads(1 To 3) As String

Private Sub buttonCancel_Click()
    Me.Hide
End Sub

Private Sub buttonInsertHead_Click()

    'Me.Hide
    'Scanner.InsertHeadModel Me.listHeadModels.Value, Me.listCoverages.Value, Me.listCoverages.Value, Me.listTemperatures.Value
    Scanner.LoadHeadToolbars Me.listHeadModels.value
End Sub

Private Sub CommandButton1_Click()
End Sub
Public Sub LoadList()
Dim xHeadModel As HeadModel
    
    For Each xHeadModel In Scanner.GetHeadModels()
        
        Me.listHeadModels.AddItem xHeadModel.ModelNumber
    
    Next xHeadModel
    
    
    
End Sub

Private Sub listHeadModelCoverage_AfterUpdate()
    Me.Hide
    Scanner.InsertHeadModel Me.listHeadModels.value & "-" & Me.listHeadModelCoverage.value
End Sub

Private Sub listHeadModelCoverage_Change()

End Sub

Private Sub listHeadModelCoverage_Click()

End Sub

Private Sub listHeadModels_Change()
Dim xHeadModel As HeadModel
Dim xCoverage As HeadModelCoverage
Dim xTemperature As HeadModelTemperature

    Set xHeadModel = Scanner.GetHeadModels(listHeadModels.value)
    
    Me.Label1.Caption = listHeadModels.value
    
    Me.listTemperatures.Clear
    Me.listCoverages.Clear
    If Not xHeadModel Is Nothing Then
        If Not xHeadModel.Temperatures Is Nothing Then
            For Each xTemperature In xHeadModel.Temperatures
                Me.listTemperatures.AddItem xTemperature.Temperature
            Next xTemperature

            If xHeadModel.Temperatures.Count > 0 Then
                If xHeadModel.DefaultTemperature <> 0 Then
                    Me.listTemperatures.value = xHeadModel.DefaultTemperature
                Else
                    Me.listTemperatures.value = xHeadModel.Temperatures(1).Temperature
                End If
            End If
        End If
        If Not xHeadModel.Coverages Is Nothing Then
            For Each xCoverage In xHeadModel.Coverages
                Me.listCoverages.AddItem xCoverage.Coverage
            Next xCoverage
            
            If xHeadModel.Coverages.Count > 0 Then Me.listCoverages.value = xHeadModel.Coverages(1).Coverage
        End If
        
    End If
    
    
    
End Sub

Private Sub listHeadModelTemperatures_Change()
Dim xHeadModel As HeadModel
Dim xTemperature As HeadModelTemperature
Dim xCoverage As HeadModelCoverage


    Set xHeadModel = Scanner.GetHeadModels(listHeadModels.value)
    If Not IsNull(Me.listHeadModelTemperatures.value) Then
        Set xTemperature = xHeadModel.Temperatures(Me.listHeadModelTemperatures.value)
        
        Me.listHeadModelTemperatures.Clear
        If Not xTemperature.Coverages Is Nothing Then
            For Each xCoverage In xTemperature.Coverages
                Me.listHeadModelCoverage.AddItem xCoverage.Coverage
                
            Next xCoverage
            If xTemperature.Coverages.Count > 0 Then Me.listHeadModelCoverage.value = xTemperature.Coverages(1).Coverage
        End If
    End If

End Sub

Private Sub listHeadModelTemperatures_Click()

End Sub

Private Sub UserForm_Activate()
    'Me.listHeadModels.RowSourceType
    LoadList
End Sub

