Attribute VB_Name = "basData"
' ---------------------------------------------------------------------
'   Win32 API declarations so that VBA can call
'   Windows functions directly
' ---------------------------------------------------------------------
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
' ---------------------------------------------------------------------
'   Win32 API structure definition as user-defined type
' ---------------------------------------------------------------------
Private Type OPENFILENAME
    lStructSize         As Long
    hWndOwner           As Long
    hInstance           As Long
    lpstrFilter         As String
    lpstrCustomFilter   As String
    nMaxCustFilter      As Long
    nFilterIndex        As Long
    lpstrFile           As String
    nMaxFile            As Long
    lpstrFileTitle      As String
    nMaxFileTitle       As Long
    lpstrInitialDir     As String
    lpstrTitle          As String
    flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As String
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As String
End Type
' ---------------------------------------------------------------------
'   Win32 API constants
' ---------------------------------------------------------------------
Private Const BIF_NEWDIALOGSTYLE         As Long = &H40
Private Const BIF_RETURNONLYFSDIRS       As Long = 1
Private Const MAX_PATH                   As Long = 260
Private Const OFN_OVERWRITEPROMPT        As Long = &H2
Private Const OFN_FILEMUSTEXIST          As Long = &H1000
Private Const OFN_PATHMUSTEXIST          As Long = &H800

Type Product
    Cost As Currency
    Labor As Double
End Type
Public Const DEFAULT_FOLDER = "X:"
Public Const MDB_FILE_NAME = DEFAULT_FOLDER & "\Autocad Export.mdb"
Public Const DATA_MDB_FILE_NAME = DEFAULT_FOLDER & "\LoopCalcData.mdb"
Public Const XML_FILE_NAME = DEFAULT_FOLDER & "\LoopCalcImport.xml"
Public Const LOOPCAD_VERSION = "0.2.3"
Public Const FSDX_VERSION = "0.2.3"


' ---------------------------------------------------------------------
'   ShowSave    Save As... common dialog
'   Arguments:  [in, String] dialog title,
'               [in, String] filter description, [optional]
'               [in, String] filter spec, [optional]
'               [in, String] default directory [optional]
'   Example call:
'   dgnFile = ShowSave ("Save Design File As...", "Microstation Files (*.dgn)", "*.dgn", "V:\shared")
'   Returns:    full path of file to be saved
' ---------------------------------------------------------------------
Public Function ShowSave( _
    ByVal strDialogTitle As String, _
    ByVal strProposed As String, _
    Optional ByVal strFilterDescr As String = "All files (*.*)", _
    Optional ByVal strFilterSpec As String = "*.*", _
    Optional ByVal strDefaultDir As String = vbNullString) As String
    On Error Resume Next
    Dim strFilter                           As String, _
        strFileSelected                     As String, _
        proposed                            As String
    Dim OFName                              As OPENFILENAME
    strFilter = strFilterDescr + Chr$(0) + strFilterSpec + Chr$(0)

    proposed = strProposed & Chr$(0) & Space$(254 - Len(strProposed)) 'Create a buffer
    Const Period                            As String = "."
    With OFName
        .lStructSize = Len(OFName) 'Set the structure size
        .hWndOwner = 0& 'Set the owner window
        .hInstance = 0& 'Set the application's instance
        .lpstrFilter = strFilter 'Set the filter
        .lpstrFile = proposed
        '.lpstrDefExt = Space$(Len(strFilterSpec))
        .lpstrDefExt = Mid$(strFilterSpec, 1 + InStr(strFilterSpec, Period))
        .nMaxFile = 255 'Set the maximum number of chars
        .lpstrFileTitle = Space$(254) 'Create a buffer
        .nMaxFileTitle = 255 'Set the maximum number of chars
        If (vbNullString <> strDefaultDir) Then _
            .lpstrInitialDir = strDefaultDir 'Set the initial directory
        .lpstrTitle = strDialogTitle 'Set the dialog title
        .flags = OFN_OVERWRITEPROMPT 'no extra flags
    End With
    If GetSaveFileName(OFName) Then 'Show the 'Save File' dialog
        strFileSelected = Trim$(OFName.lpstrFile)
        If (InStr(strFileSelected, Chr(0)) > 0) Then
            strFileSelected = Left(strFileSelected, InStr(strFileSelected, Chr(0)) - 1)
        End If
        ShowSave = Trim(strFileSelected)
    Else
        ShowSave = ""
    End If
End Function


Public Function GetConnection()
On Error GoTo ErrTrap
Dim strDSN As String
Dim strFileName As String

    strFileName = MDB_FILE_NAME
    'strDSN = "Provider=MSDASQL.1;Password=duttbutt;User ID=Dustin;Data Source=red_xfire_live;Extended Properties=""DSN=red_xfire_live;Description=xfire database;UID=marc;PWD=lukilu;APP=Microsoft Office XP;WSID=AQUA;DATABASE=xfire"""
    strDSN = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & strFileName

    If pCon Is Nothing Then
        Set pCon = New ADODB.Connection
        pCon.Open strDSN, "admin"
    End If
    If pCon.State = adStateClosed Then
        pCon.Open strDSN, "admin"
    End If
    Set GetConnection = pCon

ExitSub:
    Exit Function

ErrTrap:
    If Err.Number = -2147437259 Then
        MsgBox "The Microsoft Access file " & strFileName & " appears to be locked, you may need to close it.", vbCritical, "SprinklerScanner.basData.getConnection"
        Resume ExitSub
    Else
        MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical, "SprinklerScanner.basData.getConnection"
    End If
    
End Function

Public Function GetDataConnection()
On Error GoTo ErrTrap
Dim strDSN As String
Dim strFileName As String

    strFileName = DATA_MDB_FILE_NAME
    'strDSN = "Provider=MSDASQL.1;Password=duttbutt;User ID=Dustin;Data Source=red_xfire_live;Extended Properties=""DSN=red_xfire_live;Description=xfire database;UID=marc;PWD=lukilu;APP=Microsoft Office XP;WSID=AQUA;DATABASE=xfire"""
    strDSN = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & strFileName

    If pCon Is Nothing Then
        Set pCon = New ADODB.Connection
        pCon.Open strDSN, "admin"
    End If
    If pCon.State = adStateClosed Then
        pCon.Open strDSN, "admin"
    End If
    Set GetDataConnection = pCon

ExitSub:
    Exit Function

ErrTrap:
    If Err.Number = -2147437259 Then
        MsgBox "The Microsoft Access file " & strFileName & " appears to be locked, you may need to close it.", vbCritical, "SprinklerScanner.basData.GetDataConnection"
        Resume ExitSub
    Else
        MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical, "SprinklerScanner.basData.GetDataConnection"
    End If
    
End Function



Public Function getPipe(PipeNo As Long) As Pipe
    Dim rsItem As New ADODB.Recordset
    Dim newPipe As New Pipe
    
    
    With rsItem
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = GetConnection
        .Open "SELECT * FROM Pipes WHERE PipeNo = " & PipeNo
        If .RecordCount > 0 Then
            If Not IsNull(!Length) Then
                newPipe.ActualLength = !Length
            Else
                newPipe.ActualLength = 0
            End If
            If Not IsNull(!Diameter) Then
                newPipe.Diameter = !Diameter
            Else
                newPipe.Diameter = 0
            End If
            Set getPipe = newPipe
        Else
            'getPipe. = 0
            'getProduct.Labor = 0
            Err.Raise vbObjectError + 1, "basData.getPipe", "The Pipe Number """ & PipeNo & """ does not exist in the database."
        End If

        .Close
    End With
    
    Set rsItem = Nothing
End Function
Public Function AddNode(NodeNo As Long, CategoryID As Long, CategoryName As String)

    If ProposalID = 0 Then Exit Function
    
    Dim rsCategory As New ADODB.Recordset
    
    With rsCategory
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = GetConnection
        .Open "SELECT * FROM tblProposalCategory WHERE 0=1"
        .AddNew
        !lngProposalID = ProposalID
        !lngProposalCategoryID = CategoryID
        !strDescription = CategoryName
        
        .Update
        .Close

    End With
    

    Set rsCategory = Nothing
End Function

Public Function AddNodes(nodes As Collection, CalculationName As String, CalculationID As Long)
    Dim rsNode As New ADODB.Recordset
    Dim xNode As Node
    Dim manifoldDone As Boolean
    
    With rsNode
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = GetConnection
        .Open "SELECT * FROM Nodes"
        For Each xNode In nodes
            'Temporary section
            If xNode.NodeTypeID = NODE_TYPE_MANIFOLD_PORT Then
                If manifoldDone Then
                    GoTo NextxNode
                Else
                    manifoldDone = True
                End If
            End If
            
            .AddNew
            !NodeName = xNode.Name
            !NodeNo = xNode.NodeNumber
            !CoordinateX = xNode.insertionPoint.x
            !CoordinateY = xNode.insertionPoint.y
            !elevation = xNode.elevation
            !ModelNumber = xNode.model
            !CalculationName = CalculationName
            !CalculationID = CalculationID
            !NodeTypeID = xNode.NodeTypeID
            .Update
NextxNode:
        Next xNode
        .Close
        .ActiveConnection.Close
    End With
    

    Set rsNode = Nothing
    
End Function


Public Function AddResPipes(Pipes As Collection, CalculationName As String, CalculationID As Long)
    Dim rsPipe As New ADODB.Recordset
    Dim xPipe As ResPipe
    
    
    With rsPipe
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = GetConnection
        .Open "SELECT * FROM Pipes"
        For Each xPipe In Pipes
            .AddNew
            !PipeNo = xPipe.PipeNo
            !StartNodeNo = xPipe.StartNodeNo
            !EndNodeNo = xPipe.EndNodeNo
            !StartNodeName = xPipe.StartNodeName
            !EndNodeName = xPipe.EndNodeName
            !Diameter = xPipe.Diameter
            !Friction = xPipe.Friction
            !Length = xPipe.Length
            !CalculationName = CalculationName
            !CalculationID = CalculationID
            
            .Update
        Next xPipe
        .Close
        .ActiveConnection.Close
    End With
    

    Set rsPipe = Nothing
End Function
Public Function DeleteNodesPipes() As Boolean

    Dim lngPipeCount As Long
    Dim lngNodeCount As Long
    Dim lngHeadPairHeadCount As Long
    Dim lngHeadPairCount As Long
    Dim lngResult As Long
    Dim strCalculationName As String
    
    lngResult = MsgBox("Are you sure you want to export data?", vbYesNo, "Sprinkler System Scanner")
    
    
    
    If lngResult = VbMsgBoxResult.vbYes Then
        'Copy to archive first
        GetConnection.Execute "INSERT INTO Nodes_Archive SELECT * FROM Nodes"
        GetConnection.Execute "INSERT INTO Pipes_Archive SELECT * FROM Pipes"
        
        'Delete
        GetConnection.Execute "DELETE FROM Nodes ", lngNodeCount
        GetConnection.Execute "DELETE FROM Pipes ", lngPipeCount
        GetConnection.Execute "DELETE FROM HeadGroupHeads ", lngHeadPairHeadCount
        GetConnection.Execute "DELETE FROM HeadGroups ", lngHeadPairCount
        
        Debug.Print lngPipeCount & " Pipes Deleted"
        Debug.Print lngNodeCount & " Nodes Deleted"
        Debug.Print lngHeadPairHeadCount & " Head Pair Heads Deleted"
        Debug.Print lngHeadPairCount & " HeadPairs Deleted"
        
        DeleteNodesPipes = True
    Else
        Debug.Print "DeleteNodesPipes was cancelled!"
        DeleteNodesPipes = False
    End If
    
    
End Function



Public Function AddHeads(Heads As Collection, CalculationName As String, CalculationID As Long)
    Dim rsHead As New ADODB.Recordset
    Dim xHead As Head
    
    
    With rsHead
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = GetConnection
        .Open "SELECT * FROM Nodes"
        For Each xHead In Heads
            .AddNew
            !NodeName = xHead.Name
            !NodeNo = xHead.NodeNumber
            !CoordinateX = xHead.insertionPoint.x 'xHead.CoordinateX
            !CoordinateY = xHead.insertionPoint.y 'xHead.CoordinateY
            !elevation = xHead.elevation
            !CalculationName = CalculationName
            !CalculationID = CalculationID
            !ModelNumber = xHead.model
            !NodeTypeID = NODE_TYPE_HEAD
            .Update
        Next xHead
        .Close
        .ActiveConnection.Close
    End With
    

    Set rsHead = Nothing
    
End Function
Public Function AddHeadPairs(HeadPairs As Collection, CalculationID As Long)
    Dim rsHeadPair As New ADODB.Recordset
    Dim rsHeadPairHead As New ADODB.Recordset
    Dim xHeadPair As HeadPair
    Dim lngHeadPairNo As Long
    Dim lngHeadPairHeadNo As Long
    Dim strHeadAName As String
    Dim strHeadBName As String
    
    lngHeadPairNo = 0
    
    With rsHeadPair
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = GetConnection
        .Open "SELECT * FROM HeadGroups"
        For Each xHeadPair In HeadPairs
            xHeadPair.HeadPairNo = lngHeadPairNo
            .AddNew
            !HeadGroupNo = xHeadPair.HeadPairNo
            '!NodeName = xHeadPair.HeadA.Name & " " & xHeadPair.HeadB.Name
'            !CoordinateX = xHead.InsertionPoint.x 'xHead.CoordinateX
'            !CoordinateY = xHead.InsertionPoint.y 'xHead.CoordinateY
            !CalculationID = CalculationID
            .Update
            
            lngHeadPairNo = lngHeadPairNo + 1
        Next xHeadPair
        .Close
    End With
    
    lngHeadPairHeadNo = 0
    
    With rsHeadPairHead
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = GetConnection
        .Open "SELECT * FROM HeadGroupHeads"
        For Each xHeadPair In HeadPairs
            If xHeadPair.HeadA Is Nothing Then
                strHeadAName = "Error"
            Else
                strHeadAName = xHeadPair.HeadA.Name
            End If
            If xHeadPair.HeadB Is Nothing Then
                strHeadBName = "Error"
            Else
                strHeadBName = xHeadPair.HeadB.Name
            End If
        
        
            .AddNew
            !HeadGroupNo = xHeadPair.HeadPairNo
            !HeadGroupHeadNo = 0
            !NodeName = strHeadAName
            .Update
            
            .AddNew
            !HeadGroupNo = xHeadPair.HeadPairNo
            !HeadGroupHeadNo = 1
            !NodeName = strHeadBName
            .Update
        Next xHeadPair
        .Close
        .ActiveConnection.Close
    End With
    

    Set rsHeadPairHead = Nothing
    Set rsHeadPair = Nothing
    
End Function

Public Function AddCalculation(CalculationName As String) As Long
    Dim rsCalculations As New ADODB.Recordset
    Dim xHead As Head
    Dim lngCalculationID  As Long
    
    With rsCalculations
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = GetConnection
        .Open "SELECT * FROM Calculations"
        .AddNew
        !CalculationName = CalculationName
        .Update
        
        lngCalculationID = !CalculationID
        .Close
        .ActiveConnection.Close
    End With
    
    AddCalculation = lngCalculationID

    Set rsCalculations = Nothing
    
End Function


Public Function AddResNodes(nodes As Collection, CalculationName As String, CalculationID As Long)
    Dim rsNode As New ADODB.Recordset
    Dim xNode As ResNode
    Dim xHead As Head
    
    
    With rsNode
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = GetConnection
        .Open "SELECT * FROM Nodes"
        For Each xNode In nodes
            .AddNew
            !NodeName = xNode.NodeName
            !NodeNo = xNode.NodeNo
            !CoordinateX = xNode.CoordinateX
            !CoordinateY = xNode.CoordinateY
            !elevation = xNode.elevation
            !CalculationName = CalculationName
            !CalculationID = CalculationID
            If Mid(xNode.NodeName, 1, 1) = "H" Then
                'Set xHead = xNode
                '!ModelNumber = xHead.Model
            End If
            .Update
            'Exit For 'Do this only once!  Only one manifold
        Next xNode
        .Close
        .ActiveConnection.Close
    End With
    

    Set rsNode = Nothing
    
End Function

Public Function AddRisers(nodes As Collection, CalculationName As String, CalculationID As Long)
    Dim rsNode As New ADODB.Recordset
    Dim xNode As Node
    
    With rsNode
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = GetConnection
        .Open "SELECT * FROM Nodes"
        For Each xNode In nodes
            .AddNew
            !NodeName = xNode.Name
            !NodeNo = xNode.NodeNumber
            !CoordinateX = xNode.insertionPoint.x
            !CoordinateY = xNode.insertionPoint.y
            !elevation = xNode.elevation
            !model = xNode.model
            !CalculationName = CalculationName
            !CalculationID = CalculationID
            !NodeTypeID = xNode.NodeTypeID
            .Update
        Next xNode
        .Close
        .ActiveConnection.Close
    End With
    

    Set rsNode = Nothing
    
End Function

Public Function AddPipes(Pipes As Collection, CalculationID As Long)
    Dim rsPipe As New ADODB.Recordset
    Dim xPipe As Pipe
    Dim strStartNodeName As String
    Dim strEndNodeName As String
    
    
    
    With rsPipe
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = GetConnection
        .Open "SELECT * FROM Pipes"
        For Each xPipe In Pipes
            .AddNew
            !PipeNo = xPipe.PipeNumber
            '!StartNodeNo = xPipe.StartNodeNo
            '!EndNodeNo = xPipe.EndNodeNo
            If xPipe.StartNode Is Nothing Then
                strStartNodeName = "Error"
            Else
                strStartNodeName = xPipe.StartNode.Name
            End If
            !StartNodeName = strStartNodeName
            If xPipe.EndNode Is Nothing Then
                strEndNodeName = "Error"
            Else
                strEndNodeName = xPipe.EndNode.Name
            End If
            !EndNodeName = strEndNodeName
            !Diameter = xPipe.Diameter
            !Friction = 150 'it's all pex, to the C-Factor of 150 is correct
            !Length = xPipe.ActualLength / 12
            !CalculationName = "imported"
            !CalculationID = CalculationID
            
            .Update
        Next xPipe
        .Close
        .ActiveConnection.Close
    End With
    

    Set rsPipe = Nothing
End Function

Public Function SavePipesXML(Pipes As Collection)
    Dim xPipe As Pipe
    Dim strStartNodeName As String
    Dim strEndNodeName As String
    Dim strXMLDoc As String
    strXMLDoc = ""
    Dim strPipeRow As String
    
    
        For Each xPipe In Pipes
            strPipeRow = vbTab & "<pipes>" & vbCrLf
            strPipeRow = strPipeRow & vbTab & vbTab & "<PipeNo>" & xPipe.PipeNumber & "</PipeNo>" & vbCrLf
            If xPipe.StartNode Is Nothing Then
                strStartNodeName = "Error"
            Else
                strStartNodeName = xPipe.StartNode.Name
            End If
            strPipeRow = strPipeRow & vbTab & vbTab & "<StartNodeName>" & strStartNodeName & "</StartNodeName>" & vbCrLf
            If xPipe.EndNode Is Nothing Then
                strEndNodeName = "Error"
            Else
                strEndNodeName = xPipe.EndNode.Name
            End If
            strPipeRow = strPipeRow & vbTab & vbTab & "<EndNodeName>" & strEndNodeName & "</EndNodeName>" & vbCrLf
            strPipeRow = strPipeRow & vbTab & vbTab & "<Diameter>" & xPipe.Diameter & "</Diameter>" & vbCrLf
            'it's all pex, to the C-Factor of 150 is correct
            strPipeRow = strPipeRow & vbTab & vbTab & "<Friction>150</Friction>" & vbCrLf
            strPipeRow = strPipeRow & vbTab & vbTab & "<Length>" & xPipe.ActualLength / 12 & "</Length>" & vbCrLf
            strPipeRow = strPipeRow & vbTab & "</pipes>" & vbCrLf
            strXMLDoc = strXMLDoc & strPipeRow
        Next xPipe
        strXMLDoc = "<PipesList>" & vbCrLf & strXMLDoc & "</PipesList>" & vbCrLf
        SavePipesXML = strXMLDoc
    

End Function
Public Function SaveHeadPairsXML(HeadPairs As Collection)
    Dim xPipe As Pipe
    Dim strStartNodeName As String
    Dim strEndNodeName As String
    Dim xmlGroups As String
    Dim xmlHeads As String
    Dim strHeadGroupRow As String
    Dim strHeadRow As String
    xmlGroups = ""
    xmlHeads = ""
    
        For Each xHeadPair In HeadPairs
            xHeadPair.HeadPairNo = lngHeadPairNo
            strHeadGroupRow = vbTab & "<HeadGroups>" & vbCrLf
            strHeadGroupRow = strHeadGroupRow & vbTab & vbTab & "<HeadGroupNo>" & xHeadPair.HeadPairNo & "</HeadGroupNo>" & vbCrLf
            lngHeadPairNo = lngHeadPairNo + 1
            strHeadGroupRow = strHeadGroupRow & vbTab & "</HeadGroups>" & vbCrLf
            xmlGroups = xmlGroups & strHeadGroupRow
        Next xHeadPair
        xmlGroups = "<HeadGroupsList>" & vbCrLf & xmlGroups & "</HeadGroupsList>" & vbCrLf
        
        
        lngHeadPairHeadNo = 0
        
        For Each xHeadPair In HeadPairs
        
            If xHeadPair.HeadA Is Nothing Then
                strHeadAName = "Error"
            Else
                strHeadAName = xHeadPair.HeadA.Name
            End If
            If xHeadPair.HeadB Is Nothing Then
                strHeadBName = "Error"
            Else
                strHeadBName = xHeadPair.HeadB.Name
            End If
        
            strHead = vbTab & "<HeadGroupHeads>" & vbCrLf
            strHead = strHead & vbTab & vbTab & "<HeadGroupNo>" & xHeadPair.HeadPairNo & "</HeadGroupNo>" & vbCrLf
            strHead = strHead & vbTab & vbTab & "<HeadGroupHeadNo>0</HeadGroupHeadNo>" & vbCrLf
            strHead = strHead & vbTab & vbTab & "<NodeName>" & strHeadAName & "</NodeName>" & vbCrLf
            strHead = strHead & vbTab & "</HeadGroupHeads>" & vbCrLf
            
            strHead = strHead & vbTab & "<HeadGroupHeads>" & vbCrLf
            strHead = strHead & vbTab & vbTab & "<HeadGroupNo>" & xHeadPair.HeadPairNo & "</HeadGroupNo>" & vbCrLf
            strHead = strHead & vbTab & vbTab & "<HeadGroupHeadNo>1</HeadGroupHeadNo>" & vbCrLf
            strHead = strHead & vbTab & vbTab & "<NodeName>" & strHeadBName & "</NodeName>" & vbCrLf
            strHead = strHead & vbTab & "</HeadGroupHeads>" & vbCrLf
            xmlHeads = xmlHeads & strHead
            
        Next xHeadPair
        xmlHeads = "<HeadGroupHeadsList>" & vbCrLf & xmlHeads & "</HeadGroupHeadsList>" & vbCrLf
        
        SaveHeadPairsXML = xmlGroups & xmlHeads
    

End Function




Public Function SaveXMLFile(FilePath As String, InputString As String) As Boolean
    SaveXMLFile = SaveFile(FilePath, XMLFile(InputString))
End Function
Public Function XMLFile(InputString As String) As String
On Error GoTo ErrTrap
Dim strNewFilePath As String
Dim xmlString As String
    'That's the XML header
    
    xmlString = "<?xml version=""1.0"" encoding=""utf-8"" ?>" & vbCrLf & _
    "<!DOCTYPE fss [<!ENTITY version ""1.0.1"">]> " & vbCrLf & _
    "<FireSprinklerSystemsList>" & vbCrLf & _
    "<FireSprinklerSystems>" & vbCrLf & _
    InputString & vbCrLf & _
    "</FireSprinklerSystems>" & vbCrLf & _
    "</FireSprinklerSystemsList>" & vbCrLf
    XMLFile = xmlString

    
ExitProc:
    Exit Function
    
ErrTrap:
    MsgBox Err.Number & " : " & Err.Description, vbCritical, "XMLFile"
    Resume ExitProc
        

End Function

Public Function FSDXFile(InputString As String) As String
On Error GoTo ErrTrap

Dim xmlString As String
    'That's the XML header
    
    xmlString = "<?xml version=""1.0"" encoding=""utf-8"" ?>" & vbCrLf & _
    "<fsdx_file>" & vbCrLf & vbTab & _
        "<creation> " & vbCrLf & vbTab & vbTab & _
            "<created>" & Format(Now(), "YYYY-MM-DD") & "T" & Format(Now(), "hh:mm:ss") & "</created>" & vbCrLf & vbTab & vbTab & _
            "<software_name>LoopCad</software_name>" & vbCrLf & vbTab & vbTab & _
            "<software_version>" & LOOPCAD_VERSION & "</software_version>" & vbCrLf & vbTab & vbTab & _
            "<fsdx_version>" & FSDX_VERSION & "</fsdx_version>" & vbCrLf & vbTab & vbTab & _
            "<computer_name>" & VBA.Environ("COMPUTERNAME") & "</computer_name>" & vbCrLf & vbTab & vbTab & _
            "<dwg_file_path>" & ActiveDocument.FullName & "</dwg_file_path>" & vbCrLf & vbTab & vbTab & _
            "<dwg_folder_path>" & ActiveDocument.Path & "</dwg_folder_path>" & vbCrLf & vbTab & _
        "</creation> " & vbCrLf & vbTab & _
            InputString & vbCrLf & _
    "</fsdx_file>" & vbCrLf
    
    FSDXFile = xmlString
    
ExitProc:
    Exit Function
    
ErrTrap:
     MsgBox Err.Number & " : " & Err.Description, vbCritical, "FSDX File"
     Resume ExitProc

End Function
Public Function FSDXSystem(InputString As String) As String
On Error GoTo ErrTrap

Dim xmlString As String
    'That's the XML header
    
    xmlString = vbTab & _
        "<fire_sprinkler_system>" & vbCrLf & vbTab & vbTab & _
            InputString & vbCrLf & vbTab & _
        "</fire_sprinkler_system>" & vbCrLf
    
    FSDXSystem = xmlString
    
    
ExitProc:
    Exit Function
    
ErrTrap:
     MsgBox Err.Number & " : " & Err.Description, vbCritical, "FSDX System"
     Resume ExitProc
End Function
Public Function SaveFile(FilePath As String, InputString As String) As Boolean
On Error GoTo ErrTrap
Dim strNewFilePath As String
Dim strFileContents As String
    'That's the XML header
    
    strFileContents = InputString
    
    Open FilePath For Output Shared As #1
    Print #1, strFileContents
    Close #1
    
ExitProc:
    Exit Function
    
ErrTrap:
    If Err.Number = 76 Then
        strNewFilePath = InputBox("Please try a different file name or folder.", "Save File", FilePath)
        If strNewFilePath <> "" Then
            SaveFile strNewFilePath, InputString
        Else
            Resume ExitProc
        End If
    Else
        MsgBox Err.Number & " : " & Err.Description, vbCritical, "Save File"
        Resume ExitProc
    End If
        

End Function





    
Public Function getConnectionShaped()
On Error GoTo ErrTrap
Dim strDSN As String
Dim strFileName As String
Dim sCon As New ADODB.Connection

    strPath = MDB_FILE_NAME

    sCon.Provider = "MSDataShape"
    sCon.Properties("Data Provider").value = "Microsoft.Jet.OLEDB.4.0"
    sCon.Properties("Data Source").value = strPath
    sCon.Open

    If sCon Is Nothing Then
        Set sCon = New ADODB.Connection
        sCon.Provider = "MSDataShape"
        sCon.Properties("Data Provider").value = "Microsoft.Jet.OLEDB.4.0"
        sCon.Properties("Data Source").value = strPath
        sCon.Open
    End If
    If sCon.State = adStateClosed Then
        sCon.Provider = "MSDataShape"
        sCon.Properties("Data Provider").value = "Microsoft.Jet.OLEDB.4.0"
        sCon.Properties("Data Source").value = strPath
        sCon.Open
    End If
    Set getConnectionShaped = sCon

ExitSub:
    Exit Function

ErrTrap:
    If Err.Number = -2147437259 Then
        MsgBox "The Microsoft Access file " & strPath & " appears to be locked, you may need to close it.", vbCritical, "SprinklerScanner.basData.getConnection"
        Resume ExitSub
    Else
        MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical, "SprinklerScanner.basData.getConnection"
    End If
    
End Function


Public Function SaveNodesXML(nodes As Collection) As String
Dim xNode As Node
Dim manifoldDone As Boolean
 Dim strXMLDoc As String
Dim strNodeRow As String
    strXMLDoc = ""
    
    
        For Each xNode In nodes
            strNodeRow = vbTab & "<nodes>" & vbCrLf
            

            'Temporary section
            If xNode.NodeTypeID = NODE_TYPE_MANIFOLD_PORT Then
                If manifoldDone Then
                    GoTo NextxNode
                Else
                    manifoldDone = True
                End If
            End If
            
            
            strNodeRow = strNodeRow & vbTab & vbTab & "<NodeName>" & xNode.Name & "</NodeName>" & vbCrLf
            strNodeRow = strNodeRow & vbTab & vbTab & "<NodeNo>" & xNode.NodeNumber & "</NodeNo>" & vbCrLf
            strNodeRow = strNodeRow & vbTab & vbTab & "<CoordinateX>" & xNode.insertionPoint.x & "</CoordinateX>" & vbCrLf
            strNodeRow = strNodeRow & vbTab & vbTab & "<CoordinateY>" & xNode.insertionPoint.y & "</CoordinateY>" & vbCrLf
            strNodeRow = strNodeRow & vbTab & vbTab & "<Elevation>" & xNode.elevation & "</Elevation>" & vbCrLf
            strNodeRow = strNodeRow & vbTab & vbTab & "<ModelNumber>" & xNode.model & "</ModelNumber>" & vbCrLf
            strNodeRow = strNodeRow & vbTab & vbTab & "<NodeTypeID>" & xNode.NodeTypeID & "</NodeTypeID>" & vbCrLf
            strNodeRow = strNodeRow & vbTab & "</nodes>" & vbCrLf
            strXMLDoc = strXMLDoc & strNodeRow
NextxNode:
        Next xNode
        strXMLDoc = "<NodesList>" & vbCrLf & strXMLDoc & "</NodesList>" & vbCrLf
        SaveNodesXML = strXMLDoc

    
End Function
Public Function SaveJobXML(InputJob As Job) As String
Dim xNode As Node
Dim manifoldDone As Boolean
 Dim strXMLDoc As String
Dim strJobDataRow As String
    strXMLDoc = ""
    
    
        If Not Job Is Nothing Then
        'strJobDataRow = vbTab & "<FireSprinklerSystems>" & vbCrLf
        strJobDataRow = strJobDataRow & vbTab & vbTab & "<JobName>" & InputJob.JobName & "</JobName>" & vbCrLf
        strJobDataRow = strJobDataRow & vbTab & vbTab & "<LeadNo>" & InputJob.LeadN & "</LeadNo>" & vbCrLf
        strJobDataRow = strJobDataRow & vbTab & vbTab & "<DWGFileFolder>" & ActiveDocument.Path & "</DWGFileFolder>" & vbCrLf
        strJobDataRow = strJobDataRow & vbTab & vbTab & "<DWGFilePath>" & ActiveDocument.FullName & "</DWGFilePath>" & vbCrLf
        'strJobDataRow = strJobDataRow & vbTab & "</FireSprinklerSystems>" & vbCrLf
        strXMLDoc = strJobDataRow
        
        'strXMLDoc = "<FireSprinklerSystemsList>" & vbCrLf & strXMLDoc & "</FireSprinklerSystemsList>" & vbCrLf
        SaveJobXML = strXMLDoc
        Else
            SaveJobXML = "<job><JobName>No Job Data Saved</JobName></job>"
        End If

    
End Function

Public Function FSDXNodes(nodes As Collection) As String
Dim xNode As Node
Dim manifoldDone As Boolean
 Dim strXMLDoc As String
Dim strNodeRow As String
    strXMLDoc = ""
    
        ''no''strXMLDoc = vbTab & vbTab & "<nodes>" & vbCrLf
        strXMLDoc = ""
        
        For Each xNode In nodes
            strNodeRow = vbTab & vbTab & vbTab & vbTab & "<node>" & vbCrLf
            

            'Temporary section, omits extra manifold ports
            If xNode.NodeTypeID = NODE_TYPE_MANIFOLD_PORT Then
                If manifoldDone Then
                    GoTo NextxNode
                Else
                    manifoldDone = True
                End If
            End If
            
            
            
            strNodeRow = strNodeRow & vbTab & vbTab & vbTab & vbTab & vbTab & "<name>" & xNode.Name & "</name>" & vbCrLf
            strNodeRow = strNodeRow & vbTab & vbTab & vbTab & vbTab & vbTab & "<node_no>" & xNode.NodeNumber & "</node_no>" & vbCrLf
            strNodeRow = strNodeRow & vbTab & vbTab & vbTab & vbTab & vbTab & "<insertion_point><x>" & xNode.insertionPoint.x & "</x><y>" & xNode.insertionPoint.y & "</y><z>" & xNode.insertionPoint.Z & "</z></insertion_point>" & vbCrLf
            strNodeRow = strNodeRow & vbTab & vbTab & vbTab & vbTab & vbTab & "<elevation>" & xNode.elevation & "</elevation>" & vbCrLf
            strNodeRow = strNodeRow & vbTab & vbTab & vbTab & vbTab & vbTab & "<model_number>" & xNode.model & "</model_number>" & vbCrLf
            strNodeRow = strNodeRow & vbTab & vbTab & vbTab & vbTab & vbTab & "<node_type_id>" & xNode.NodeTypeID & "</node_type_id>" & vbCrLf
            strNodeRow = strNodeRow & vbTab & vbTab & vbTab & vbTab & "</node>" & vbCrLf
            strXMLDoc = strXMLDoc & strNodeRow
NextxNode:
        Next xNode
        
        ''no''strXMLDoc = strXMLDoc & vbTab & vbTab & "</nodes>" & vbCrLf
        
        FSDXNodes = strXMLDoc

    
End Function

Public Function FSDXPipes(Pipes As Collection) As String
    Dim xPipe As Pipe
    Dim strStartNodeName As String
    Dim strEndNodeName As String
    Dim strXMLDoc As String
    strXMLDoc = ""
    Dim strPipeRow As String
    Dim v As Vertex
    Dim lastX As Long
    Dim strLineType As String
    Dim strLineLayer As String
    Dim intLineColor As Integer
    
        ''no''strXMLDoc = vbTab & vbTab & vbTab & "<pipes>" & vbCrLf
        strXMLDoc = ""
        For Each xPipe In Pipes
            strPipeRow = vbTab & vbTab & vbTab & "<pipe>" & vbCrLf
            strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & "<index>" & xPipe.PipeNumber & "</index>" & vbCrLf
            If xPipe.StartNode Is Nothing Then
                strStartNodeName = "Error"
                'Err.Raise 1, "FSDX Pipes", "Pipe " & xPipe.PipeNumber & " is missing a start node! File will not be saved. " '  The other node is " & strStartNodeName
                GoTo Failure
            Else
                strStartNodeName = xPipe.StartNode.Name
            End If
            strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & "<start_node_name>" & strStartNodeName & "</start_node_name>" & vbCrLf
            If xPipe.EndNode Is Nothing Then
                strEndNodeName = "Error"
                Err.Raise 1, "FSDX Pipes", "Pipe " & xPipe.PipeNumber & " is missing an end node!  The other node is " & strStartNodeName
                GoTo Failure
            Else
                strEndNodeName = xPipe.EndNode.Name
            End If
            If xPipe.polyline Is Nothing Then
                strLineType = ""
                Err.Raise 1, "FSDX Pipes", "Pipe " & xPipe.PipeNumber & " is missing a Polyline object and layer type!"
                GoTo Failure
            Else
                strLineType = xPipe.polyline.LineType
                strLineLayer = xPipe.polyline.layer
                intLineColor = xPipe.polyline.Color
            End If
            
            strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & "<end_node_name>" & strEndNodeName & "</end_node_name>" & vbCrLf
            strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & "<internal_diameter>" & xPipe.Diameter & "</internal_diameter>" & vbCrLf
            strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & "<line_type>" & strLineType & "</line_type>" & vbCrLf
            strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & "<color>" & intLineColor & "</color>" & vbCrLf
            strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & "<layer>" & strLineLayer & "</layer>" & vbCrLf
            'it's all pex, to the C-Factor of 150 is correct
            strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & "<friction>150</friction>" & vbCrLf
            strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & "<length>" & xPipe.ActualLength / 12 & "</length>" & vbCrLf
            
            If Not xPipe.polyline Is Nothing Then
                strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & "<vertexes>" & vbCrLf
                lastX = (xPipe.polyline.Coordinates.Count - 1)
                For x = 0 To lastX Step 2
                        strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & "<vertex>" & vbCrLf
                        strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & vbTab & "<x>" & xPipe.polyline.Coordinates(x) & "</x>" & vbCrLf
                        strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & vbTab & "<y>" & xPipe.polyline.Coordinates(x + 1) & "</y>" & vbCrLf
                        strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & vbTab & "<z>0</z>" & vbCrLf
                        strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & "</vertex>" & vbCrLf
                Next x
                strPipeRow = strPipeRow & vbTab & vbTab & vbTab & vbTab & "</vertexes>" & vbCrLf
            Else
                MsgBox "Pipe " & xPipe.PipeNumber & " does not have a Polyline attached to it!", vbCritical, "FSDX Pipes"
                Exit Function
            End If
            
            strPipeRow = strPipeRow & vbTab & vbTab & vbTab & "</pipe>" & vbCrLf
            strXMLDoc = strXMLDoc & strPipeRow
        Next xPipe
        ''no''strXMLDoc = strXMLDoc & vbTab & vbTab & "</pipes>" & vbCrLf
        FSDXPipes = strXMLDoc

ExitProc:
    Exit Function
    
Failure:
    MsgBox Err.Description & vbCrLf & "Export canceled", vbCritical, "FSDXPipes"
    GoTo ExitProc

End Function
Public Function FSDXJob(InputJob As Job) As String
Dim xNode As Node
Dim manifoldDone As Boolean
 Dim strXMLDoc As String
Dim strJobDataRow As String
    strXMLDoc = ""

    strJobDataRow = vbTab & vbTab & "<job>" & vbCrLf
    strJobDataRow = strJobDataRow & vbTab & vbTab & vbTab & "<calculated_by_company>" & InputJob.CalculatedByCompany & "</calculated_by_company>" & vbCrLf
    strJobDataRow = strJobDataRow & vbTab & vbTab & vbTab & "<job_name>" & InputJob.JobName & "</job_name>" & vbCrLf
    strJobDataRow = strJobDataRow & vbTab & vbTab & vbTab & "<lead_no>" & InputJob.LeadN & "</lead_no>" & vbCrLf
    strJobDataRow = strJobDataRow & vbTab & vbTab & vbTab & "<site_location>" & InputJob.SiteLocation & "</site_location>" & vbCrLf
    strJobDataRow = strJobDataRow & vbTab & vbTab & vbTab & "<static_pressure>" & InputJob.StaticPressure & "</static_pressure>" & vbCrLf
    strJobDataRow = strJobDataRow & vbTab & vbTab & vbTab & "<residual_pressure>" & InputJob.ResidualPressure & "</residual_pressure>" & vbCrLf
    strJobDataRow = strJobDataRow & vbTab & vbTab & vbTab & "<available_flow>" & InputJob.AvailableFlow & "</available_flow>" & vbCrLf
    strJobDataRow = strJobDataRow & vbTab & vbTab & vbTab & "<meter_elevation>" & InputJob.MeterElevation & "</meter_elevation>" & vbCrLf
    strJobDataRow = strJobDataRow & vbTab & vbTab & vbTab & "<meter_pipe_length>" & InputJob.MeterPipeLength & "</meter_pipe_length>" & vbCrLf
    strJobDataRow = strJobDataRow & vbTab & vbTab & vbTab & "<meter_pipe_internal_diameter>" & InputJob.MeterPipeInternalDiameter & "</meter_pipe_internal_diameter>" & vbCrLf
    strJobDataRow = strJobDataRow & vbTab & vbTab & "</job>" & vbCrLf
    strXMLDoc = strJobDataRow
    
    'strXMLDoc = "<FireSprinklerSystemsList>" & vbCrLf & strXMLDoc & "</FireSprinklerSystemsList>" & vbCrLf
    FSDXJob = strXMLDoc

    
End Function
Public Function FSDXHeadPairs(HeadPairs As Collection)
    Dim xPipe As Pipe
    Dim strStartNodeName As String
    Dim strEndNodeName As String
    Dim xmlGroups As String
    Dim xmlHeads As String
    Dim strHeadGroupRow As String
    Dim strHeadRow As String
    xmlGroups = ""
    xmlHeads = ""
    
        For Each xHeadPair In HeadPairs
            xHeadPair.HeadPairNo = lngHeadPairNo
            strHeadGroupRow = vbTab & vbTab & "<head_group>" & vbCrLf
            strHeadGroupRow = strHeadGroupRow & vbTab & vbTab & vbTab & "<index>" & xHeadPair.HeadPairNo & "</index>" & vbCrLf
            
            If xHeadPair.HeadA Is Nothing Then
                strHeadAName = "Error"
            Else
                strHeadAName = xHeadPair.HeadA.Name
            End If
            If xHeadPair.HeadB Is Nothing Then
                strHeadBName = "Error"
            Else
                strHeadBName = xHeadPair.HeadB.Name
            End If

            strHeadGroupRow = strHeadGroupRow & vbTab & vbTab & vbTab & "<heads>" & vbCrLf
            strHeadGroupRow = strHeadGroupRow & vbTab & vbTab & vbTab & vbTab & "<head>" & vbCrLf
            strHeadGroupRow = strHeadGroupRow & vbTab & vbTab & vbTab & vbTab & vbTab & "<name>" & strHeadAName & "</name>" & vbCrLf
            strHeadGroupRow = strHeadGroupRow & vbTab & vbTab & vbTab & vbTab & "</head>" & vbCrLf
            strHeadGroupRow = strHeadGroupRow & vbTab & vbTab & vbTab & vbTab & "<head>" & vbCrLf
            strHeadGroupRow = strHeadGroupRow & vbTab & vbTab & vbTab & vbTab & vbTab & "<name>" & strHeadBName & "</name>" & vbCrLf
            strHeadGroupRow = strHeadGroupRow & vbTab & vbTab & vbTab & vbTab & "</head>" & vbCrLf
            strHeadGroupRow = strHeadGroupRow & vbTab & vbTab & vbTab & "</heads>" & vbCrLf
            
            lngHeadPairNo = lngHeadPairNo + 1
            strHeadGroupRow = strHeadGroupRow & vbTab & "</head_group>" & vbCrLf
            xmlGroups = xmlGroups & strHeadGroupRow
        Next xHeadPair
        ''no''xmlGroups = vbTab & vbTab & "<head_groups>" & vbCrLf & xmlGroups & vbTab & vbTab & "</head_groups>" & vbCrLf
        
        
        lngHeadPairHeadNo = 0
        
'        For Each xHeadPair In HeadPairs
'
'            If xHeadPair.HeadA Is Nothing Then
'                strHeadAName = "Error"
'            Else
'                strHeadAName = xHeadPair.HeadA.Name
'            End If
'            If xHeadPair.HeadB Is Nothing Then
'                strHeadBName = "Error"
'            Else
'                strHeadBName = xHeadPair.HeadB.Name
'            End If
'
'            strHead = vbTab & "<HeadGroupHeads>" & vbCrLf
'            strHead = strHead & vbTab & vbTab & "<HeadGroupNo>" & xHeadPair.HeadPairNo & "</HeadGroupNo>" & vbCrLf
'            strHead = strHead & vbTab & vbTab & "<HeadGroupHeadNo>0</HeadGroupHeadNo>" & vbCrLf
'            strHead = strHead & vbTab & vbTab & "<NodeName>" & strHeadAName & "</NodeName>" & vbCrLf
'            strHead = strHead & vbTab & "</HeadGroupHeads>" & vbCrLf
'
'            strHead = strHead & vbTab & "<HeadGroupHeads>" & vbCrLf
'            strHead = strHead & vbTab & vbTab & "<HeadGroupNo>" & xHeadPair.HeadPairNo & "</HeadGroupNo>" & vbCrLf
'            strHead = strHead & vbTab & vbTab & "<HeadGroupHeadNo>1</HeadGroupHeadNo>" & vbCrLf
'            strHead = strHead & vbTab & vbTab & "<NodeName>" & strHeadBName & "</NodeName>" & vbCrLf
'            strHead = strHead & vbTab & "</HeadGroupHeads>" & vbCrLf
'            xmlHeads = xmlHeads & strHead
'
'        Next xHeadPair
'        xmlHeads = "<HeadGroupHeadsList>" & vbCrLf & xmlHeads & "</HeadGroupHeadsList>" & vbCrLf
        
        FSDXHeadPairs = xmlGroups '& xmlHeads
    

End Function

Public Sub LoadHeadModels()
Dim con As ADODB.Connection
Dim rsModels As ADODB.Recordset
Dim rsItems As ADODB.Recordset
Dim rsTemperatures As ADODB.Recordset
Dim rsSlopes As ADODB.Recordset

Dim model As HeadModel
Dim ModelItem As HeadModelItem
Dim Slope As HeadModelSlope
Dim Temperature As HeadModelTemperature
Dim tempSuffix As String

Dim vDefaultTemperature As Variant
Dim lngSlope As Long

    Set rsTemperatures = New ADODB.Recordset
    With rsTemperatures
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = GetDataConnection
        .Open "SELECT ht.* FROM head_temperature ht ORDER BY ht.temperature "
    End With

    Set rsItems = New ADODB.Recordset
    With rsItems
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = GetDataConnection
        .Open "SELECT hi.* FROM head_item hi ORDER BY hi.head_model_id, hi.sidewall, hi.slope, hi.temperature, hi.coverage_width, hi.coverage_depth "
    End With

    Set rsSlopes = New ADODB.Recordset
    With rsSlopes
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = GetDataConnection
        .Open "SELECT DISTINCT hi.head_model_id, hi.slope, hi.sidewall FROM head_item hi" ' ORDER BY hi.head_item, hi.slope "
    End With
    
        
    Set rsModels = New ADODB.Recordset
    With rsModels
        rsModels.CursorLocation = adUseClient
        rsModels.CursorType = adOpenDynamic
        rsModels.LockType = adLockOptimistic
        rsModels.ActiveConnection = GetDataConnection
        rsModels.Open "SELECT hm.*, m.name AS manufacturer_name FROM head_model hm INNER JOIN manufacturer m ON m.manufacturer_id = hm.manufacturer_id WHERE m.active <> 0 ORDER BY hm.sort, m.name, hm.model_no DESC "
        
        If rsModels.RecordCount > 0 Then
            Set Scanner.headModels = New Collection
            Do While Not rsModels.EOF
                Set model = New HeadModel
                model.HeadModelID = rsModels!head_model_id
                model.ModelNumber = rsModels!model_no
                Set model.Slopes = New Collection
                Set model.HeadModelItems = New Collection
                
                vDefaultTemperature = rsModels!default_temperature
                If Not IsNull(vDefaultTemperature) Then
                    model.DefaultTemperature = vDefaultTemperature
                Else
                    model.DefaultTemperature = 0
                End If
                Scanner.headModels.Add model, CStr(model.ModelNumber)
                
                rsTemperatures.Filter = "head_model_id = " & model.HeadModelID
                If rsTemperatures.RecordCount > 0 Then
                    rsTemperatures.MoveFirst
                    Set model.Temperatures = New Collection
                    Do While Not rsTemperatures.EOF
                        Set Temperature = New HeadModelTemperature
                        Temperature.Temperature = rsTemperatures!Temperature
                        model.Temperatures.Add Temperature, CStr(Temperature.Temperature)
                        rsTemperatures.MoveNext
                    Loop
                End If
                
                rsSlopes.Filter = "head_model_id = " & model.HeadModelID
                If rsSlopes.RecordCount > 0 Then
                    rsSlopes.MoveFirst
                    Do While Not rsSlopes.EOF
                        Set Slope = New HeadModelSlope
                        Slope.SlopeIndex = rsSlopes!Slope
                        Slope.Sidewall = rsSlopes!Sidewall

                        If Slope.Sidewall Then
                            Slope.SlopeName = "SW"
                        Else
                            Slope.SlopeName = "S" & rsSlopes!Slope
                        End If
                        Set Slope.HeadModel = model
                        Set Slope.HeadModelItems = New Collection
                        model.Slopes.Add Slope, Slope.SlopeName
                        
                        rsItems.Filter = "head_model_id = " & model.HeadModelID & " AND slope = " & Slope.SlopeIndex
                        If rsItems.RecordCount > 0 Then
                            rsItems.MoveFirst
                            
                            Do While Not rsItems.EOF
                                Set ModelItem = New HeadModelItem
                                ModelItem.Coverage = CStr(rsItems!coverage_depth) & "x" & CStr(rsItems!coverage_width)
                                Set ModelItem.Slope = Slope
                                ModelItem.Coverage = CStr(rsItems!coverage_depth) & "x" & CStr(rsItems!coverage_width)
                                ModelItem.CoverageDepth = rsItems!coverage_depth
                                ModelItem.CoverageWidth = rsItems!coverage_width
                                If IsNull(rsItems!Temperature) Then
                                    ModelItem.Temp = 155
                                Else
                                    ModelItem.Temp = rsItems!Temperature
                                End If
                                Slope.HeadModelItems.Add ModelItem, ModelItem.Coverage & "-" & ModelItem.Temp
                                model.HeadModelItems.Add ModelItem, CStr(ModelItem.Slope.SlopeIndex) & "-" & ModelItem.Coverage & "-" & ModelItem.Temp
                                
                                rsItems.MoveNext
                            Loop
                        End If
                        
                        rsSlopes.MoveNext
                    Loop
                End If
                rsModels.MoveNext
            Loop
        End If
    End With
            
    rsModels.Close
    rsItems.Close
    rsTemperatures.Close
    rsSlopes.Close
    
    Set rsModels = Nothing
    Set rsItems = Nothing
    Set rsTemperatures = Nothing
    Set rsSlopes = Nothing
            
End Sub
Public Sub TestXMLEncode()
    Debug.Print XMLEncode("""Hello"" & 'Goodbye' <world>")
End Sub
Public Function XMLEncodeAndrew()
' This procedure creates XML ActiveDocument
' and saves it to disk.
' Requires msxml.dll (Go to Project --> References and
' and choose Microsoft XML version 2.0, or whatever the
' current version you have installed)
' The example given below will write the following XML
' ActiveDocuments.
'
' <Family>
'    <Member Relationship="Father">
'       <Name>Some Guy</Name>
'    </Member>
' </Family>
'
'but it should be clear how to modify the code
'to create your own ActiveDocuments


   
   Dim objDom As DOMActiveDocument
   Dim objRootElem As IXMLDOMElement
   Dim objMemberElem As IXMLDOMElement
   Dim objMemberRel As IXMLDOMAttribute
   Dim objMemberName As IXMLDOMElement
   
   Set objDom = New DOMActiveDocument
   
   ' Creates root element
   Set objRootElem = objDom.createElement("Family")
   objDom.appendChild objRootElem
   
   ' Creates Member element
   Set objMemberElem = objDom.createElement("Member")
   objRootElem.appendChild objMemberElem
   
   ' Creates Attribute to the Member Element
   Set objMemberRel = objDom.createAttribute("Relationship")
   objMemberRel.nodeValue = "Father"
   objMemberElem.setAttributeNode objMemberRel
   
   ' Create element under Member element, and
   ' gives value "some guy"
   Set objMemberName = objDom.createElement("Name")
   objMemberElem.appendChild objMemberName
   objMemberName.text = "Some ""Guy"" & <hello> 'world'"

   ' Saves XML data to disk.
   objDom.Save ("\\TSCLIENT\C\temp\andrew.xml")
    
End Function
