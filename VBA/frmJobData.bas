VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmJobData 
   Caption         =   "Job Data"
   ClientHeight    =   10215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7680
   OleObjectBlob   =   "frmJobData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmJobData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdCancel_Click()
    
    Me.Hide
    
End Sub

Private Sub cmdOK_Click()

    ScanJobData
    SaveJobData
    Me.Hide
    
End Sub
Public Sub SaveJobData()

    If Scanner.Job Is Nothing Then
        Set Scanner.Job = New Job
    End If
           
    Scanner.Job.CalculatedByCompany = Me.CalculatedByCompany
    Scanner.Job.JobNumber = Me.JobNumber
    Scanner.Job.JobName = Me.JobName
    Scanner.Job.JobSiteLocation = Me.JobSiteLocation
    Scanner.Job.SprinklerPipeType = Me.SprinklerPipeType
    Scanner.Job.SprinklerFittingType = Me.SprinklerFittingType
    Scanner.Job.SupplyName = Me.SupplyName
    Scanner.Job.SupplyAvailableFlow = Me.SupplyAvailableFlow
    Scanner.Job.SupplyStaticPressure = Me.SupplyStaticPressure
    Scanner.Job.SupplyResidualPressure = Me.SupplyResidualPressure
    Scanner.Job.SupplyElevation = Me.SupplyElevation
    Scanner.Job.SupplyPipeLength = Me.SupplyManifoldPipeLength
    Scanner.Job.SupplyPipeType = Me.SupplyManifoldPipeType
    Scanner.Job.SupplyPipeSize = Me.SupplyManifoldPipeSize
    Scanner.Job.SupplyPipeInternalDiameter = Me.SupplyPipeInternalDiameter
    Scanner.Job.SupplyPipeFittingsSummary = Me.SupplyPipeFittingsSummary
    Scanner.Job.SupplyPipeFittingsEquivLength = Me.SupplyPipeFittingsEquivLength
    Scanner.Job.SupplyPipeWaterFlowPressureLoss = Me.SupplyPipeWaterFlowPressureLoss
    Scanner.Job.WaterFlowMakeModel = Me.WaterFlowMakeModel
    Scanner.Job.DomesticFlowAdded = Me.DomesticFlowAdded
        
    JobData.SetJobData
    
End Sub

Private Sub txtLeadN_Change()

End Sub

Private Sub UserForm_Activate()

    LoadJobData
    
End Sub

Private Sub UserForm_Initialize()

    LoadJobData
    LoadCalculatedByCompanies
    
End Sub
Public Sub LoadJobData()

    JobData.ScanJobData
    
    If Scanner.Job Is Nothing Then
        Set Scanner.Job = New Job
    Else
        Me.CalculatedByCompany = Scanner.Job.CalculatedByCompany
        Me.JobNumber = Scanner.Job.JobNumber
        Me.JobName = Scanner.Job.JobName
        Me.JobSiteLocation = Scanner.Job.JobSiteLocation
        Me.SprinklerPipeType = Scanner.Job.SprinklerPipeType
        Me.SprinklerFittingType = Scanner.Job.SprinklerFittingType
        Me.SupplyName = Scanner.Job.SupplyName
        Me.SupplyAvailableFlow = Scanner.Job.SupplyAvailableFlow
        Me.SupplyStaticPressure = Scanner.Job.SupplyStaticPressure
        Me.SupplyResidualPressure = Scanner.Job.SupplyResidualPressure
        Me.SupplyElevation = Scanner.Job.SupplyElevation
        Me.SupplyPipeLength = Scanner.Job.SupplyManifoldPipeLength
        Me.SupplyPipeType = Scanner.Job.SupplyManifoldPipeType
        Me.SupplyPipeSize = Scanner.Job.SupplyManifoldPipeSize
        Me.SupplyPipeInternalDiameter = Scanner.Job.SupplyPipeInternalDiameter
        Me.SupplyPipeFittingsSummary = Scanner.Job.SupplyPipeFittingsSummary
        Me.SupplyPipeFittingsEquivLength = Scanner.Job.SupplyPipeFittingsEquivLength
        Me.SupplyPipeWaterFlowPressureLoss = Scanner.Job.SupplyPipeWaterFlowPressureLoss
        Me.WaterFlowMakeModel = Scanner.Job.WaterFlowMakeModel
        Me.DomesticFlowAdded = Scanner.Job.DomesticFlowAdded
       
    End If
    
End Sub
Public Sub LoadCalculatedByCompanies()

    Me.CalculatedByCompany.Clear
    Me.CalculatedByCompany.AddItem "X-Fire Security"
    Me.CalculatedByCompany.AddItem "13dpex.com"
    
End Sub
