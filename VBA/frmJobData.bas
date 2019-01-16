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
    
    Scanner.Job.JobName = Me.JobName
    Scanner.Job.JobSiteLocation = Me.JobSiteLocation
    Scanner.Job.JobNumber = Me.JobNumber
    Scanner.Job.SupplyStaticPressure = Me.SupplyStaticPressure
    Scanner.Job.SupplyResidualPressure = Me.SupplyResidualPressure
    Scanner.Job.SupplyAvailableFlow = Me.SupplyAvailableFlow
    Scanner.Job.SupplyElevation = Me.SupplyElevation
    Scanner.Job.SupplyManifoldPipeLength = Me.SupplyPipeLength
    Scanner.Job.SupplyManifoldPipeInternalDiameter = Me.SupplyPipeInternalDiameter
    Scanner.Job.CalculatedByCompany = Me.CalculatedByCompany
        
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
        Me.JobName = Scanner.Job.JobName
        Me.JobSiteLocation = Scanner.Job.JobSiteLocation
        Me.JobNumber = Scanner.Job.JobNumber
        Me.SupplyStaticPressure = Scanner.Job.SupplyStaticPressure
        Me.SupplyResidualPressure = Scanner.Job.SupplyResidualPressure
        Me.SupplyAvailableFlow = Scanner.Job.SupplyAvailableFlow
        Me.SupplyElevation = Scanner.Job.SupplyElevation
        Me.SupplyPipeLength = Scanner.Job.SupplyManifoldPipeLength
        Me.SupplyPipeInternalDiameter = Scanner.Job.SupplyManifoldPipeInternalDiameter
        Me.CalculatedByCompany = Scanner.Job.CalculatedByCompany
    End If
    
End Sub
Public Sub LoadCalculatedByCompanies()

    Me.CalculatedByCompany.Clear
    Me.CalculatedByCompany.AddItem "X-Fire Security"
    Me.CalculatedByCompany.AddItem "13dpex.com"
    
End Sub
