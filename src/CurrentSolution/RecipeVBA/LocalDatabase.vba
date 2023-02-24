Option Explicit
Option Compare Text
'////////////////////////////////////////
'
'LocalDatabase Class Version 2.0.0
'v6.4.00_01
'Created by: Carey Warren
'Modified by: Carey Warren
'Date: 2015-10-14

'This object contains a local copy of the data in the TPMDB.  The Hmi references this copy and never the database directly.
Private Type tMaterial
    PLC_ID As Long
    Name As String
    Display As String
    SiteAlias As String 'Added AKB /10/21
    EU As String 'Added_ND 20230109
End Type

Private Type tMaterialClass
    id As Long
    Name As String
    MaterialList() As tMaterial
End Type

Private Type tUsers
    id As Long
    Name As String
End Type

Private Type tStatus
    id As Long
    Description As String
    Description2 As String
End Type

Private Type tMode
    id As Long
    Description As String
    phase As String
End Type

Private Type tCondition
    id As Long
    Description As String
End Type

Private Type tMaterialGroup
    id As Long
    Name As String
End Type

Public db As dbConnect
Private db1 As dbConnect
Private db2 As dbConnect

Private Tagdb As dbConnect
Private Tagdb1 As dbConnect
Private Tagdb2 As dbConnect

Private lUsers() As tUsers
Private lMaterialClass() As tMaterialClass
Private lMaterial() As tMaterial
Private lCstat() As tStatus
Private lEstat() As tStatus
Private lPhMode() As tMode
Private lEqMode() As tMode
Private lYesNo() As tStatus
Private lStepDescription As Collection
Private lQAResult() As tStatus
Private lActivity() As tStatus
Private lCCPhaseList() As tStatus
Private lCondition() As tCondition
Private lConditionList() As String
Private lCMIS As Collection
Private lCIPRecipe() As String
Private lCIPStandaloneRecipe() As String
Private lAreas As Collection

Private lDSNTagDB As String
Private lDSNTagDB2 As String
Private lDSNTPMDB As String
Private lDSNTPMDB2 As String
Private lconnectiontype As Integer
Private lAlternateModeTexts As Integer

Const DB_SingleSQL = 1 'standard install, 1 sql server
Const DB_RedundantSQL = 2 '2 SQL servers for tagdb and TPMDB

Private rs As adodb.Recordset
Private rs1 As adodb.Recordset
Private rs2 As adodb.Recordset
Private rs3 As adodb.Recordset
Private fld As adodb.Field
Private par As adodb.Parameter

Event StatusChanged(status As String, value As Integer)
Private lMaterialGroup() As tMaterialGroup
Private debugmode As Boolean


Private Sub LoadAreas()
    Dim temp As Collection
    Set lAreas = New Collection
    Dim rsa As adodb.Recordset
    'Execute Command
    Set rsa = Tagdb.executeCommand("ui_getAreas")
    'object type list returned from stored procedure
    While Not rsa.EOF
        Set temp = New Collection
        temp.Add rsa.Fields("area").value, "Name"
        temp.Add LoadUnits(rsa.Fields("area").value), "Units"
        lAreas.Add temp, rsa.Fields("area").value
        rsa.MoveNext
    Wend
    rsa.Close
    'LogDiagnosticsMessage "Object Types loaded from database", ftDiagSeverityInfo
End Sub

Public Sub MaterialGroupUpdate(id As Long, newname As String)
    'insert the note to db

    Dim par1 As adodb.Parameter
    Set par1 = New adodb.Parameter
    par1.Type = adInteger                      'Parameter type
    par1.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par1.value = id                 'Parameter value (MaterialClass)
                            'Max. size for Parameter value
    
    Dim par2 As adodb.Parameter
    Set par2 = New adodb.Parameter
    par2.Type = adVarChar                      'Parameter type
    par2.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par2.value = newname                 'Parameter value (MaterialClass)
    par2.Size = 50                             'Max. size for Parameter value
                            'Max. size for Parameter value

    Tagdb.ConnectDSN
    If Tagdb.getConnection.State = adStateOpen Then
        'Execute Command
        Tagdb.executeCommand2Par "[uiMaterialGroupUpdate]", par1, par2
        LoadMaterialGroupData
        Tagdb.Disconnect
    End If

End Sub


Private Sub LoadMaterialGroupData()
    Dim i As Integer
   
    'Execute Command
    Set rs = Tagdb.executeCommand("ui_getMaterialGroupNames")
    ReDim lMaterialGroup(rs.RecordCount)
    i = 1
    'material list returned from stored procedure
    While Not rs.EOF
        lMaterialGroup(i).id = rs.Fields("ID").value
        lMaterialGroup(i).Name = rs.Fields("Name").value
        i = i + 1
        rs.MoveNext
    Wend
    rs.Close
'    LogDiagnosticsMessage "Material Class loaded from database", ftDiagSeverityInfo
End Sub

Private Function LoadUnits(area As String) As Collection
    Dim temp As Collection
    Dim tempObj As New Collection
    Dim par As adodb.Parameter
    Dim rsu As adodb.Recordset
    Set par = New adodb.Parameter
    par.Type = adVarChar                      'Parameter type
    par.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par.value = area                 'Parameter value (MaterialClass)
    par.Size = 50                             'Max. size for Parameter value

    'Execute setup Command with Parameters
    Set rsu = Tagdb.executeCommand1Par("ui_getUnits", par)              'Execute command
    'Objects returned from stored procedure
    While Not rsu.EOF
        Set temp = New Collection
        temp.Add rsu.Fields("unit").value, "Name"
        temp.Add LoadUnitGroups(rsu.Fields("unit").value), "Groups"
        temp.Add LoadPhases(rsu.Fields("unit").value), "Phases"
        tempObj.Add temp, rsu.Fields("unit").value
        rsu.MoveNext
    Wend
    rsu.Close
    Set LoadUnits = tempObj
End Function

Private Function LoadPhases(unit As String) As Collection
    Dim temp As Collection
    Dim tempObj As New Collection
    Dim rsp As adodb.Recordset
    Dim par As adodb.Parameter
    Set par = New adodb.Parameter
    par.Type = adVarChar                      'Parameter type
    par.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par.value = unit                 'Parameter value (MaterialClass)
    par.Size = 50                             'Max. size for Parameter value

    'Execute setup Command with Parameters
    Set rsp = Tagdb.executeCommand1Par("[ui_getUnitPhases]", par)              'Execute command
    'Objects returned from stored procedure
    While Not rsp.EOF
        Set temp = New Collection
        temp.Add rsp.Fields("type").value, "Name"
        temp.Add LoadPhaseGroups(rsp.Fields("type").value, unit), "Groups"
        'temp.Add LoadPhaseSteps(rsp.Fields("type").value, Unit), "Steps"
        tempObj.Add temp, rsp.Fields("type").value
        rsp.MoveNext
    Wend
    rsp.Close
    Set LoadPhases = tempObj
End Function


Private Function LoadUnitGroups(unit As String) As Collection
    Dim temp As Collection
    Dim tempObj As New Collection
    Dim rsg As adodb.Recordset
    Dim par As adodb.Parameter
    Set par = New adodb.Parameter
    par.Type = adVarChar                      'Parameter type
    par.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par.value = unit                 'Parameter value (MaterialClass)
    par.Size = 50                             'Max. size for Parameter value

    'Execute setup Command with Parameters
    Set rsg = Tagdb.executeCommand1Par("ui_getUnitGroups", par)              'Execute command
    'Objects returned from stored procedure
    While Not rsg.EOF
        Set temp = New Collection
        temp.Add rsg.Fields("group").value, "Name"
        temp.Add LoadUnitParameters(unit, rsg.Fields("group").value), "Parameters"
        tempObj.Add temp, rsg.Fields("group").value
        rsg.MoveNext
    Wend
    rsg.Close
    Set LoadUnitGroups = tempObj
End Function

Private Function LoadPhaseGroups(phase As String, unit As String) As Collection
    Dim temp As Collection
    Dim tempObj As New Collection
    Dim rsg As adodb.Recordset
    Dim par As adodb.Parameter
    Set par = New adodb.Parameter
    par.Type = adVarChar                      'Parameter type
    par.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par.value = phase                 'Parameter value (MaterialClass)
    par.Size = 50                             'Max. size for Parameter value

    Dim par1 As adodb.Parameter
    Set par1 = New adodb.Parameter
    par1.Type = adVarChar                      'Parameter type
    par1.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par1.value = unit                 'Parameter value (MaterialClass)
    par1.Size = 50                             'Max. size for Parameter value

    'Execute setup Command with Parameters
    Set rsg = Tagdb.executeCommand2Par("[ui_getPhaseGroups]", par, par1)             'Execute command
    'Objects returned from stored procedure
    While Not rsg.EOF
        Set temp = New Collection
        temp.Add rsg.Fields("group").value, "Name"
        temp.Add LoadPhaseParameters(phase, rsg.Fields("group").value, unit), "Parameters"
        tempObj.Add temp, rsg.Fields("group").value
        rsg.MoveNext
    Wend
    rsg.Close
    Set LoadPhaseGroups = tempObj
End Function

Private Sub LoadPhaseSteps()

    Dim tsteps As Collection
    Dim tstep As Collection
    Dim ttag As Collection
    Dim rsg As adodb.Recordset
    Dim rsg1 As adodb.Recordset
    
    Set lStepDescription = New Collection
    
   
    
    Set rsg = Tagdb.executeCommand("[ui_getStepDescriptions_Tagname]")             'Execute command
    'Objects returned from stored procedure
    While Not rsg.EOF
        Set tsteps = New Collection
        Set ttag = New Collection
        Dim par1 As adodb.Parameter
        Set par1 = New adodb.Parameter
        par1.Type = adVarChar                      'Parameter type
        par1.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
        par1.Size = 100
        par1.value = rsg.Fields("Tagname").value
        Set rsg1 = Tagdb.executeCommand1Par("[ui_getStepDescriptions_Steps]", par1)
        While Not rsg1.EOF
            Set tstep = New Collection
            tstep.Add rsg1.Fields("Step").value, "Step"
            tstep.Add rsg1.Fields("Description").value, "Description"
            tsteps.Add tstep, CStr(rsg1.Fields("Step").value)
            rsg1.MoveNext
        Wend
        rsg1.Close
        ttag.Add tsteps, "Steps"
        ttag.Add CStr(rsg.Fields("Tagname").value), "Name"
        lStepDescription.Add ttag, CStr(rsg.Fields("Tagname").value)
        rsg.MoveNext
    Wend
    rsg.Close

End Sub
'
Private Function LoadUnitParameters(unit As String, group As String) As Collection
    Dim temp As Collection
    Dim tempPar As New Collection
    Dim rsp1 As adodb.Recordset
    Dim par As adodb.Parameter
    Set par = New adodb.Parameter
    par.Type = adVarChar                      'Parameter type
    par.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par.value = unit                 'Parameter value (MaterialClass)
    par.Size = 50                             'Max. size for Parameter value

    Dim par1 As adodb.Parameter
    Set par1 = New adodb.Parameter
    par1.Type = adVarChar                      'Parameter type
    par1.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par1.value = group                 'Parameter value (MaterialClass)
    par1.Size = 50                             'Max. size for Parameter value


    'Execute setup Command with Parameters
    Set rsp1 = Tagdb.executeCommand2Par("[ui_getUnitGroupsParameters]", par, par1)             'Execute command

    'parameters returned from stored procedure
    While Not rsp1.EOF
        Set temp = New Collection
        temp.Add rsp1.Fields("TagName").value, "TagName"
        temp.Add Trim(rsp1.Fields("Description").value), "Description"
        temp.Add Trim(rsp1.Fields("EU").value), "EU"
        temp.Add rsp1.Fields("Max").value, "Max"
        temp.Add rsp1.Fields("Min").value, "Min"
        temp.Add rsp1.Fields("value").value, "Value"
        temp.Add rsp1.Fields("datatype").value, "DataType"
        tempPar.Add temp, rsp1.Fields("TagName").value
        rsp1.MoveNext
    Wend
    rsp1.Close
    Set LoadUnitParameters = tempPar
End Function

Private Function LoadPhaseParameters(phase As String, group As String, unit As String) As Collection
    Dim temp As Collection
    Dim tempPar As New Collection
    Dim rsp1 As adodb.Recordset
    Dim par As adodb.Parameter
    Set par = New adodb.Parameter
    par.Type = adVarChar                      'Parameter type
    par.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par.value = phase                 'Parameter value (MaterialClass)
    par.Size = 50                             'Max. size for Parameter value

    Dim par1 As adodb.Parameter
    Set par1 = New adodb.Parameter
    par1.Type = adVarChar                      'Parameter type
    par1.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par1.value = group                 'Parameter value (MaterialClass)
    par1.Size = 50                             'Max. size for Parameter value

    Dim par3 As adodb.Parameter
    Set par3 = New adodb.Parameter
    par3.Type = adVarChar                      'Parameter type
    par3.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par3.value = unit                 'Parameter value (MaterialClass)
    par3.Size = 50                             'Max. size for Parameter value


    'Execute setup Command with Parameters
    Set rsp1 = Tagdb.executeCommand3Par("[ui_getPhaseGroupsParameters]", par, par1, par3)            'Execute command

    'parameters returned from stored procedure
    While Not rsp1.EOF
        Set temp = New Collection
        temp.Add rsp1.Fields("TagName").value, "TagName"
        temp.Add Trim(rsp1.Fields("Description").value), "Description"
        temp.Add Trim(rsp1.Fields("EU").value), "EU"
        temp.Add rsp1.Fields("Max").value, "Max"
        temp.Add rsp1.Fields("Min").value, "Min"
        temp.Add rsp1.Fields("value").value, "Value"
        temp.Add rsp1.Fields("datatype").value, "DataType"
        tempPar.Add temp, rsp1.Fields("TagName").value
        rsp1.MoveNext
    Wend
    rsp1.Close
    Set LoadPhaseParameters = tempPar
End Function

Private Sub LoadCMIS()
    Dim temp As Collection
    Set lCMIS = New Collection
    Dim rsa As adodb.Recordset
    'Execute Command
    Set rsa = Tagdb.executeCommand("[ui_getCMIS]")
    'object type list returned from stored procedure
    While Not rsa.EOF
        Set temp = New Collection
        temp.Add rsa.Fields("tagname").value, "Name"
        temp.Add rsa.Fields("display").value, "Display"
        temp.Add rsa.Fields("topic").value, "Topic"
        temp.Add rsa.Fields("displayname").value, "DisplayName"
        temp.Add rsa.Fields("description").value, "Description"
        
        lCMIS.Add temp, rsa.Fields("topic").value & rsa.Fields("tagname").value
        rsa.MoveNext
    Wend
    rsa.Close
End Sub

Private Sub LoadMaterialData()
    lMaterial = LoadMaterialList("%")
'    LogDiagnosticsMessage "Material loaded from database", ftDiagSeverityInfo
End Sub

Private Sub LoadMaterialClassData()
    Dim i As Integer
   
    'Execute Command
    Set rs = db.executeCommand("ui_getMaterialClassList")
    ReDim lMaterialClass(rs.RecordCount)
    i = 1
    'material list returned from stored procedure
    While Not rs.EOF
        lMaterialClass(i).id = rs.Fields("ID").value
        lMaterialClass(i).Name = rs.Fields("Name").value
        lMaterialClass(i).MaterialList = LoadMaterialList(lMaterialClass(i).Name)
        i = i + 1
        rs.MoveNext
    Wend
    rs.Close
'    LogDiagnosticsMessage "Material Class loaded from database", ftDiagSeverityInfo
End Sub

Private Function LoadMaterialList(Materialclass As String) As tMaterial()
    Dim temp() As tMaterial
    Dim i As Integer
    Set par = New adodb.Parameter
    par.Type = adVarChar                      'Parameter type
    par.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par.value = Materialclass                 'Parameter value (MaterialClass)
    par.Size = 30                             'Max. size for Parameter value
    
    'Execute setup Command with Parameters
    Set rs1 = db.executeCommand1Par("ui_getMaterialList", par)              'Execute command
    ReDim temp(rs1.RecordCount)
    i = 1
    'UserID returned from stored procedure
    While Not rs1.EOF
        temp(i).PLC_ID = rs1.Fields("PLC_ID").value
        temp(i).Name = rs1.Fields("Name").value
        temp(i).Display = temp(i).PLC_ID & Chr(9) & temp(i).Name
        temp(i).EU = rs1.Fields("EU").value 'ND_20220109
                
        'Added AKB 8/10/21
        If IsNull(rs1.Fields("SiteMaterialAlias").value) Then
            temp(i).SiteAlias = ""
        Else
            temp(i).SiteAlias = rs1.Fields("SiteMaterialAlias").value
        End If

        i = i + 1
        rs1.MoveNext
    Wend
    rs1.Close
    LoadMaterialList = temp
End Function

Private Sub LoadUserData()
    Dim i As Integer
    
    'Execute setup Command with Parameters
    Set rs = db.executeCommand("ui_getUsers")               'Execute command
    ReDim lUsers(rs.RecordCount)
    i = 1
    'UserID returned from stored procedure
    While Not rs.EOF
        lUsers(i).id = rs.Fields("PLC_ID").value
        lUsers(i).Name = rs.Fields("Name").value
        i = i + 1
        rs.MoveNext
    Wend
    rs.Close
'    LogDiagnosticsMessage "Users loaded from database", ftDiagSeverityInfo
End Sub

Private Sub LoadEstatData()
    Dim i As Integer
    
    'Execute setup Command with Parameters
    Set rs = db.executeCommand("ui_getEstatTexts")               'Execute command
    ReDim lEstat(rs.RecordCount)
    i = 1
    While Not rs.EOF
        lEstat(i).id = rs.Fields("ID").value
        lEstat(i).Description = rs.Fields("Description").value
        i = i + 1
        rs.MoveNext
    Wend
    rs.Close
'    LogDiagnosticsMessage "Estat loaded from database", ftDiagSeverityInfo
End Sub

Private Sub LoadCstatData()
    Dim i As Integer
    
    'Execute setup Command with Parameters
    Set rs = db.executeCommand("ui_getCstatTexts")               'Execute command
    ReDim lCstat(rs.RecordCount)
    i = 1
    While Not rs.EOF
        lCstat(i).id = rs.Fields("ID").value
        lCstat(i).Description = rs.Fields("Description").value
        i = i + 1
        rs.MoveNext
    Wend
    rs.Close
'    LogDiagnosticsMessage "Cstat loaded from database", ftDiagSeverityInfo
End Sub


Private Sub LoadPhModeData()
    Dim i As Integer
    
    'Execute setup Command with Parameters
    Set rs = db.executeCommand("ui_getPhaseModeTexts")               'Execute command
    ReDim lPhMode(rs.RecordCount)
    i = 1
    While Not rs.EOF
        lPhMode(i).id = rs.Fields("ID").value
        lPhMode(i).Description = rs.Fields("Description").value
        lPhMode(i).phase = rs.Fields("Phase").value
        i = i + 1
        rs.MoveNext
    Wend
    rs.Close

'    LogDiagnosticsMessage "Cstat loaded from database", ftDiagSeverityInfo
End Sub

Private Sub LoadEqModeData()
    Dim i As Integer
    
    'Execute setup Command with Parameters
    Set rs = db.executeCommand("ui_getEquipmentModeTexts")               'Execute command
    ReDim lEqMode(rs.RecordCount)
    i = 1
    While Not rs.EOF
        lEqMode(i).id = rs.Fields("ID").value
        lEqMode(i).Description = rs.Fields("Description").value
        lEqMode(i).phase = rs.Fields("Unit").value
        i = i + 1
        rs.MoveNext
    Wend
    rs.Close

'    LogDiagnosticsMessage "Cstat loaded from database", ftDiagSeverityInfo
End Sub


Private Sub LoadPhModeData_Alternate()
    Dim i As Integer
    
    'Execute setup Command with Parameters
    Set rs = Tagdb.executeCommand("ui_getPhaseModeTexts")               'Execute command
    ReDim lPhMode(rs.RecordCount)
    i = 1
    While Not rs.EOF
        lPhMode(i).id = rs.Fields("ID").value
        lPhMode(i).Description = rs.Fields("Description").value
        lPhMode(i).phase = rs.Fields("Phase").value
        i = i + 1
        rs.MoveNext
    Wend
    rs.Close

'    LogDiagnosticsMessage "Cstat loaded from database", ftDiagSeverityInfo
End Sub

Private Sub LoadEqModeData_Alternate()
    Dim i As Integer
    
    'Execute setup Command with Parameters
    Set rs = Tagdb.executeCommand("ui_getEquipmentModeTexts")               'Execute command
    ReDim lEqMode(rs.RecordCount)
    i = 1
    While Not rs.EOF
        lEqMode(i).id = rs.Fields("ID").value
        lEqMode(i).Description = rs.Fields("Description").value
        lEqMode(i).phase = rs.Fields("Unit").value
        i = i + 1
        rs.MoveNext
    Wend
    rs.Close

'    LogDiagnosticsMessage "Cstat loaded from database", ftDiagSeverityInfo
End Sub

Private Sub LoadYesNo()
    ReDim lYesNo(2)
    lYesNo(2).id = 1
    lYesNo(2).Description = "Yes"
    lYesNo(1).id = 0
    lYesNo(1).Description = "No"
End Sub

Private Sub LoadQAResultData()
    Dim i As Integer
    
    'Execute setup Command with Parameters
    Set rs = db.executeCommand("ui_getQAResults")               'Execute command
    ReDim lQAResult(rs.RecordCount)
    i = 1
    While Not rs.EOF
        lQAResult(i).id = rs.Fields("ID").value
        lQAResult(i).Description = rs.Fields("Description").value
        i = i + 1
        rs.MoveNext
    Wend
    rs.Close

'    LogDiagnosticsMessage "Cstat loaded from database", ftDiagSeverityInfo
End Sub

Private Sub LoadConditionTextsData()
    Dim i As Long
    
    'Execute setup Command with Parameters
    Set rs = db.executeCommand("ui_getConditionTexts")               'Execute command
    ReDim lCondition(rs.RecordCount)
    ReDim lConditionList(2, rs.RecordCount)
    i = 1
    While Not rs.EOF
        lCondition(i).id = rs.Fields("ID").value
        lCondition(i).Description = rs.Fields("Description").value
        lConditionList(1, i) = lCondition(i).id
        lConditionList(2, i) = lCondition(i).Description
        i = i + 1
        rs.MoveNext
    Wend
    rs.Close
'    LogDiagnosticsMessage "Condition Texts loaded from database", ftDiagSeverityInfo
End Sub

Private Sub LoadCIPRecipeData()
    Dim i As Integer
    Dim par As adodb.Parameter
    Set par = New adodb.Parameter
    par.Type = adVarChar                      'Parameter type
    par.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par.value = 2                             'Parameter value
    par.Size = 50
    
    Dim par1 As adodb.Parameter
    Set par1 = New adodb.Parameter
    par1.Type = adVarChar                      'Parameter type
    par1.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par1.value = "en"                           'Parameter value
    par1.Size = 10
    
    'Execute Command
    Set rs = Tagdb.executeCommand2Par("ui_CIPRecipeGet", par, par1)
    ReDim lCIPRecipe(2, rs.RecordCount)
    i = 1
    'cip recipe list returned from stored procedure
    While Not rs.EOF
        lCIPRecipe(1, i) = rs.Fields("PLC_ID").value
        lCIPRecipe(2, i) = rs.Fields("Description").value
        i = i + 1
        rs.MoveNext
    Wend
    rs.Close

End Sub

Private Sub LoadCIPStandaloneRecipeData()
    Dim i As Integer
    Dim par As adodb.Parameter
    Set par = New adodb.Parameter
    par.Type = adVarChar                      'Parameter type
    par.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par.value = 4                             'Parameter value
    par.Size = 50
    
    Dim par1 As adodb.Parameter
    Set par1 = New adodb.Parameter
    par1.Type = adVarChar                      'Parameter type
    par1.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par1.value = "en"                           'Parameter value
    par1.Size = 10
    
    'Execute Command
    Set rs = Tagdb.executeCommand2Par("ui_CIPRecipeGet", par, par1)
    ReDim lCIPStandaloneRecipe(2, rs.RecordCount)
    i = 1
    'cip recipe list returned from stored procedure
    While Not rs.EOF
        lCIPStandaloneRecipe(1, i) = rs.Fields("PLC_ID").value
        lCIPStandaloneRecipe(2, i) = rs.Fields("Description").value
        i = i + 1
        rs.MoveNext
    Wend
    rs.Close

End Sub


Private Sub LoadCIPActivity()
    Dim i As Integer
    
    'Execute Command
    Set rs = Tagdb.executeCommand("ui_getCleaningActivity")
    ReDim lActivity(rs.RecordCount)
    i = 1
    'cip recipe list returned from stored procedure
    While Not rs.EOF
        lActivity(i).id = rs.Fields("ID").value
        lActivity(i).Description = rs.Fields("Name").value
        lActivity(i).Description2 = rs.Fields("Phase_Name").value
        i = i + 1
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub LoadCCPhaseList()
    Dim i As Integer
    
    'Execute Command
    Set rs = Tagdb.executeCommand("ui_getCCPhaseList")
    ReDim lCCPhaseList(rs.RecordCount)
    i = 1
    'cip recipe list returned from stored procedure
    While Not rs.EOF
        lCCPhaseList(i).id = rs.Fields("ID").value
        lCCPhaseList(i).Description = rs.Fields("Name").value
        i = i + 1
        rs.MoveNext
    Wend
    rs.Close
End Sub


Public Sub Initialize()
On Error GoTo ErrHandler
    Select Case ConnectionType
    Case DB_SingleSQL
        'TPMDB
        Set db = New dbConnect
        db.DSN = DSNTPMDB
        'TagDB
        Set Tagdb = New dbConnect
        Tagdb.DSN = DSNTagDB
    Case DB_RedundantSQL
        'TPMDB Primary
        Set db1 = New dbConnect
        db1.DSN = DSNTPMDB
        'TPMDB Primary
        Set db2 = New dbConnect
        db2.DSN = DSNTPMDB2
        'default to primary
        Set db = db1
        'TagDB Primary
        Set Tagdb1 = New dbConnect
        Tagdb1.DSN = DSNTagDB
        'TagDB Primary
        Set Tagdb2 = New dbConnect
        Tagdb2.DSN = DSNTagDB2
        'default to primary
        Set Tagdb = Tagdb1
    End Select
    
Exit Sub
    'Error message
ErrHandler:
    
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on LocalDatabase Initialize", ftDiagSeverityError
End Sub
'
Public Function GetParameters(area As String, unit As String, Optional phase As String = "none") As Collection
    'returns a group collection that contains collections of parameters
    Dim temp As New Collection
    Dim carea As Collection
    Dim cunit As Collection
    Dim cphase As Collection

    If (UCase(phase) = "NONE") Or (UCase(phase) = "_UNIT") Then
        'unit parameters
        For Each carea In lAreas
            'If carea("Name") = Area Then
                For Each cunit In carea("Units")
                    If cunit("Name") = unit Then
                        Set temp = cunit("Groups")
                       GoTo found
                    End If
                Next
            'End If
        Next
    Else
        'phase parameters
         For Each carea In lAreas
            'If carea("Name") = Area Then
                For Each cunit In carea("Units")
                    If cunit("Name") = unit Then
                        For Each cphase In cunit("Phases")
                            If cphase("Name") = phase Then
                                Set temp = cphase("Groups")
                                GoTo found
                            End If
                        Next
                    End If
                Next
            'End If
        Next
    End If


found:
    Set GetParameters = temp
End Function


Public Function GetPhaseStep(Tagname As String, StepNumber As String) As String
    'returns a group collection that contains collections of parameters
    Dim temp As String
    Dim temp1 As String
    Dim lstep As Collection
    Dim tstep As Collection
    temp = "Not Defined"

    For Each tstep In lStepDescription
        If tstep("Name") = Tagname Then
            For Each lstep In tstep("Steps")
                If lstep("Step") = StepNumber Then
                    temp = lstep("Description")
                    GoTo found
                End If
            Next
        End If
    Next

found:
    temp1 = StepNumber & " - " & temp
    
    GetPhaseStep = temp1
End Function

'like search used by faceplatecipinfo for cip circuit step descriptions
Public Function GetPhaseStep_LikeSearch(Tagname As String, StepNumber As String) As String
    'returns a group collection that contains collections of parameters
    Dim temp As String
    Dim temp1 As String
    Dim lstep As Collection
    Dim tstep As Collection
    temp = "Not Defined"

    For Each tstep In lStepDescription
        If tstep("Name") Like Tagname Then
            For Each lstep In tstep("Steps")
                If lstep("Step") = StepNumber Then
                    temp = lstep("Description")
                    GoTo found
                End If
            Next
        End If
    Next

found:
    temp1 = StepNumber & " - " & temp
    
    GetPhaseStep_LikeSearch = temp1
End Function

'
Public Sub ReloadSelectedParameters(area As String, unit As String, group As String, Optional phase As String = "none")
    'returns a group collection that contains collections of parameters
    Dim temp As New Collection
    Dim carea As Collection
    Dim cunit As Collection
    Dim cphase As Collection
    Dim cgroup As Collection

    If phase = "none" Then
        'unit parameters
        For Each carea In lAreas
            If carea("Name") = area Then
                For Each cunit In carea("Units")
                    If cunit("Name") = unit Then
                        For Each cgroup In cunit("Groups")
                            If cgroup("Name") = group Then
                                Tagdb.ConnectDSN
                                If Tagdb.getConnection.State = adStateOpen Then
                                    cgroup.Remove ("Parameters")
                                    cgroup.Add LoadUnitParameters(unit, group), "Parameters"
                                    Tagdb.Disconnect
                                End If
                                GoTo found
                            End If
                        Next
                    End If
                Next
            End If
        Next
    Else
        'phase parameters
         For Each carea In lAreas
            If carea("Name") = area Then
                For Each cunit In carea("Units")
                    If cunit("Name") = unit Then
                        For Each cphase In cunit("Phases")
                            If cphase("Name") = phase Then
                                For Each cgroup In cphase("Groups")
                                    If cgroup("Name") = group Then
                                        Tagdb.ConnectDSN
                                        If Tagdb.getConnection.State = adStateOpen Then
                                            cgroup.Remove ("Parameters")
                                            cgroup.Add LoadPhaseParameters(phase, group, unit), "Parameters"
                                            Tagdb.Disconnect
                                        End If
                                        GoTo found
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        Next
    End If


found:

End Sub

Public Function GetUserID(Name As String) As Integer
    Dim i As Integer
    Dim temp As Integer
    'initialize to not found
    temp = 999
    For i = 1 To UBound(lUsers)
        If lUsers(i).Name = Name Then
            temp = lUsers(i).id
            GoTo found
        End If
    Next i
found:
    GetUserID = temp
End Function

Public Function GetMaterialName(plcid As Long) As String
    Dim i As Integer
    Dim temp As String
    'initialize to not found
    temp = "ID: " & CStr(plcid) & " missing"
    For i = 1 To UBound(lMaterial)
        If lMaterial(i).PLC_ID = plcid Then
            temp = lMaterial(i).Name
            GoTo found
        End If
    Next i
found:
    GetMaterialName = temp
End Function

Public Function GetSiteMaterialAlias2(sta As String) As Integer
    Dim i As Integer
    Dim temp As Integer
    'initialize to not found
    temp = 999
    For i = 1 To UBound(lMaterial)
        If lMaterial(i).SiteAlias = sta Then
            temp = lMaterial(i).PLC_ID
            GoTo found
        End If
    Next i
found:
    GetSiteMaterialAlias2 = temp
End Function

Public Function GetPhaseMode(plcid As Long, phase As String) As String
    Dim i As Integer
    Dim temp As String
    'initialize to not found
    temp = "Not Found"
    For i = 1 To UBound(lPhMode)
        If UCase(lPhMode(i).phase) = UCase(phase) Then
            temp = "Not Defined"
            If lPhMode(i).id = plcid Then
                temp = lPhMode(i).Description
                GoTo found
            End If
        End If
    Next i
found:
    GetPhaseMode = temp
End Function

Public Function GetEquipmentMode(plcid As Long, unit As String) As String
    Dim i As Integer
    Dim temp As String
    'initialize to not found
    'temp = "ID: " & CStr(PLCID) & " missing"
    temp = "Not Found"
    For i = 1 To UBound(lEqMode)
'        If UCase(lEqMode(i).phase) = UCase(unit) Then
'            temp = "Not Defined"
'            If lEqMode(i).ID = plcid Then
'                temp = lPhMode(i).Description
'                GoTo found
'            End If
'        End If

        If lEqMode(i).id = plcid And UCase(lEqMode(i).phase) = UCase(unit) Then
            temp = lEqMode(i).Description
            GoTo found
        End If
    Next i
found:
    GetEquipmentMode = temp
End Function

Public Function GetCIPRecipeName(plcid As Long) As String
    Dim i As Integer
    Dim temp As String
    'initialize to not found
    temp = ""
    For i = 1 To UBound(lCIPRecipe, 2)
        If lCIPRecipe(1, i) = plcid Then
            temp = lCIPRecipe(2, i)
            GoTo found
        End If
    Next i
found:
    GetCIPRecipeName = temp
End Function


Public Function GetCIPStandaloneRecipeName(plcid As Long) As String
    Dim i As Integer
    Dim temp As String
    'initialize to not found
    temp = ""
    For i = 1 To UBound(lCIPStandaloneRecipe, 2)
        If lCIPStandaloneRecipe(1, i) = plcid Then
            temp = lCIPStandaloneRecipe(2, i)
            GoTo found
        End If
    Next i
found:
    GetCIPStandaloneRecipeName = temp
End Function


Public Function GetStatusName(StatusID As Long, StatusType As String) As String
    Dim i As Integer
    Dim lStatus() As tStatus
    Dim temp As String
    temp = "Not defined"
    
    Select Case UCase(StatusType)
    Case UCase("cstat")
        lStatus = lCstat
    Case UCase("Estat")
        lStatus = lEstat
    Case UCase("Activity")
        lStatus = lActivity
    Case UCase("ActivityPhase")
        lStatus = lActivity
    Case UCase("CC")
        lStatus = lCCPhaseList
    Case UCase("YesNo")
        lStatus = lYesNo
    Case UCase("QA")
        lStatus = lQAResult
    Case Else
        lStatus = lCstat
    End Select
    
    For i = 1 To UBound(lStatus)
        If lStatus(i).id = StatusID Then
            If UCase(StatusType) = UCase("ActivityPhase") Then
                temp = lStatus(i).Description2
            Else
                temp = lStatus(i).Description
            End If
        End If
    Next i
    GetStatusName = temp
End Function

Public Function GetConditionList() As String()
    'return an array of condition texts and ids
    GetConditionList = lConditionList
End Function
'
Public Function GetAreasList() As Collection
    'return an collection of areas
    Set GetAreasList = lAreas
End Function

Public Function GetArea(area As String) As Collection

   'returns a area collection
    Dim temp As New Collection
    Dim carea As Collection
    Dim cunit As Collection
    Dim cphase As Collection

    'unit parameters
    For Each carea In lAreas
        If carea("Name") = area Then
            Set temp = carea
            GoTo found
        End If
    Next
found:
    Set GetArea = temp

End Function

Public Function GetUnit(unit As String) As Collection

   'returns a unit collection that contains collections of parameters
    Dim temp As New Collection
    Dim carea As Collection
    Dim cunit As Collection
    Dim cphase As Collection

    'unit parameters
    For Each carea In lAreas
        For Each cunit In carea("units")
            If cunit("Name") = unit Then
                Set temp = cunit
                GoTo found
            End If
        Next
    Next
found:
    Set GetUnit = temp

End Function

Public Function GetCIPRecipeList() As String()
    'return an array of CIP Recipes names and ids
    GetCIPRecipeList = lCIPRecipe
End Function


Public Function GetCIPStandaloneRecipeList() As Collection
    'return an array of CIP Recipes names and ids

    Dim temp As Collection
    Dim Col As New Collection
    Dim i As Integer

    For i = 1 To UBound(lCIPStandaloneRecipe, 2)
        Set temp = New Collection
        temp.Add lCIPStandaloneRecipe(1, i), "PLC_ID"
        temp.Add lCIPStandaloneRecipe(2, i), "Name"
        Col.Add temp
    Next i
    Set GetCIPStandaloneRecipeList = Col
End Function

Public Function getEnabledCIPRecipeList(EnabledRecipes As Long) As Collection

    Dim temp As Collection
    Dim Col As New Collection
    Dim i As Integer

    For i = 1 To UBound(lCIPRecipe, 2)
        If ((2 ^ (i - 1)) And EnabledRecipes) > 0 Then
            Set temp = New Collection
            temp.Add lCIPRecipe(1, i), "PLC_ID"
            temp.Add lCIPRecipe(2, i), "Name"
            Col.Add temp

        End If
    Next i
    Set getEnabledCIPRecipeList = Col
End Function


Public Sub GetStatusCombo(StatusComboBox As Object, StatusID As Long, StatusType As String)
    Dim i As Integer
    
    Dim lStatus() As tStatus

    Select Case UCase(StatusType)
    Case UCase("cstat")
        lStatus = lCstat
    Case UCase("Estat")
        lStatus = lEstat
    Case Else
        lStatus = lCstat
    End Select
    'Fill the combobox and select the match
    For i = 1 To UBound(lStatus)
        StatusComboBox.AddItem (CStr(lStatus(i).id) & Chr(9) & lStatus(i).Description)
        If lStatus(i).id = StatusID Then
            StatusComboBox.ListIndex = i - 1
        End If
    Next i
    
    If StatusComboBox.ListCount < 1 Then
        StatusComboBox.AddItem StatusID & Chr(9) & "Not defined"
    End If
End Sub

Public Function GetStatusList(StatusType As String) As Collection
    Dim i As Integer
    Dim Col As New Collection
    Dim temp As Collection
    Dim lStatus() As tStatus

    Select Case UCase(StatusType)
    Case UCase("cstat")
        lStatus = lCstat
    Case UCase("Estat")
        lStatus = lEstat
    Case UCase("Activity")
        lStatus = lActivity
    Case UCase("YesNo")
        lStatus = lYesNo
    Case UCase("QA")
        lStatus = lQAResult
    Case Else
        lStatus = lCstat
    End Select
    'Fill the combobox and select the match
    For i = 1 To UBound(lStatus)
        Set temp = New Collection
        temp.Add lStatus(i).id, "PLC_ID"
        temp.Add lStatus(i).Description, "Name"
        Col.Add temp
    Next i
    Set GetStatusList = Col
End Function

Public Function GetModeList(ChooseType As String, SelectionType As String) As Collection
    
    Dim i As Integer
    Dim Col As New Collection
    Dim temp As Collection
    Dim lMode() As tMode

    Select Case UCase(ChooseType)
    Case UCase("Phase")
        lMode = lPhMode
    Case UCase("Equipment")
        lMode = lEqMode
    Case Else
        lMode = lPhMode
    End Select
    
    For i = 1 To UBound(lMode)
        If SelectionType = lMode(i).phase Then
            Set temp = New Collection
            temp.Add lMode(i).id, "PLC_ID"
            temp.Add lMode(i).Description, "Name"
            Col.Add temp
        End If
    Next i
    Set GetModeList = Col
End Function


Public Sub GetCMListBox(ListBox() As Variant, CMType As String, Filter As String)
    'Fill the listbox based on filter
    Dim i As Integer
    Dim x, y As Integer
    Dim c As Collection
    Dim Transpose() As Variant
    i = 0
    ReDim ListBox(4, 0)
    For Each c In lCMIS
        If (UCase(c("Display")) = UCase(CMType)) Or (CMType = "%") Then
            ReDim Preserve ListBox(4, UBound(ListBox, 2) + 1)
            If UCase(c("DisplayName")) Like UCase(Filter) Then
                ListBox(0, i) = c("Name")
                ListBox(1, i) = c("Display")
                ListBox(2, i) = c("Topic")
                ListBox(3, i) = c("DisplayName")
                i = i + 1
            End If
        End If
    Next

   
    'transpose the array. have to do this because only first dimension can be changed with redim
    If i = 0 Then
        ReDim ListBox(0, 0)
    Else
        ReDim Transpose(i, 4)
        For x = 0 To i
            For y = 0 To 3
                Transpose(x, y) = ListBox(y, x)
            Next y
        Next x
        ListBox = Transpose
    End If
    
End Sub


Public Sub GetMaterialCombo(MaterialComboBox As Object, MaterialID As Long, Materialclass As String)
    Dim i As Integer
    Dim j As Integer
    'see if the requested material class is all or selected class
    If Materialclass = "%" Then
        For j = 1 To UBound(lMaterial)
            MaterialComboBox.AddItem (lMaterial(j).Display)
            If lMaterial(j).PLC_ID = MaterialID Then
                MaterialComboBox.ListIndex = j - 1
            End If
        Next j
    Else
        'search through class till match found
        For i = 1 To UBound(lMaterialClass)
            If UCase(lMaterialClass(i).Name) = UCase(Materialclass) Then
                'go through list of materials and find a match and add to combo box
                For j = 1 To UBound(lMaterialClass(i).MaterialList)
                    MaterialComboBox.AddItem (lMaterialClass(i).MaterialList(j).Display)
                    If lMaterialClass(i).MaterialList(j).PLC_ID = MaterialID Then
                        MaterialComboBox.ListIndex = j - 1
                    End If
                Next j
                GoTo found
             
            End If
        Next i
    End If
found:
    If MaterialComboBox.ListCount < 1 Then
        MaterialComboBox.AddItem "ID: " & MaterialID & " missing"
    End If
End Sub

Public Function GetMaterialList(Materialclass As String) As Collection
'On Error Resume Next
    Dim i As Integer, j As Integer, k As Integer
    Dim x As Integer, y As Integer
    Dim Materialclasses() As String
    Dim temp As Collection
    Dim Col As New Collection
    Dim checkcol As New Collection
    'see if the requested material class is all or selected class
    If Materialclass = "%" Then
        For j = 1 To UBound(lMaterial)
            Set temp = New Collection
            temp.Add lMaterial(j).PLC_ID, "PLC_ID"
            temp.Add lMaterial(j).Name, "Name"
            temp.Add lMaterial(j).SiteAlias, "SiteAlias" 'ERP Change
            Col.Add temp
        Next j
    Else
        'Check if more than one material class is required
        ReDim Materialclasses(0)
        If Materialclass Like "*|*" Then
            x = 1: y = 1
            Do 'Until y <= 0
                y = InStr(x, Materialclass, "|")
                If y > 0 Then
                    Materialclasses(UBound(Materialclasses)) = Mid(Materialclass, x, y - x)
                Else 'select the last material class in the string
                    Materialclasses(UBound(Materialclasses)) = Mid(Materialclass, x, Len(Materialclass) - x + 1)
                    Exit Do
                End If
                ReDim Preserve Materialclasses(UBound(Materialclasses) + 1)
                x = y + 1
            Loop
        Else
            Materialclasses(0) = Materialclass
        End If
        
        For k = 0 To UBound(Materialclasses)
            For i = 1 To UBound(lMaterialClass)
                If UCase(lMaterialClass(i).Name) = UCase(Materialclasses(k)) Then 'UCase(Materialclass) Then
                    'go through list of materials and find a match and add to combo box
                    For j = 1 To UBound(lMaterialClass(i).MaterialList)
                        Set temp = New Collection
                        temp.Add lMaterialClass(i).MaterialList(j).PLC_ID, "PLC_ID"
                        temp.Add lMaterialClass(i).MaterialList(j).Name, "name"
                        temp.Add lMaterialClass(i).MaterialList(j).SiteAlias, "SiteAlias" 'ERP Change
                        On Error GoTo errinsert
                        'make sure isn't already in list
                        For Each checkcol In Col
                            If checkcol("PLC_ID") = lMaterialClass(i).MaterialList(j).PLC_ID Then
                                GoTo dontadd
                            End If
                        Next
                        Col.Add temp
dontadd:
                        
                        
errinsert:
                    Next j
                End If
            Next i
        Next k
    
    End If
    Set GetMaterialList = Col
    
End Function

Public Sub GetMaterialClassCombo(MaterialComboBox As Object)
    Dim i As Integer
    Dim j As Integer
    'return a combobox list for material class

    For j = 1 To UBound(lMaterialClass)
        MaterialComboBox.AddItem (lMaterialClass(j).Name)
    Next j

End Sub

Public Sub GetMaterialClassCollection(MaterialComboBox As Collection)

    Dim j As Integer
    Dim t As Collection
    Dim r As Collection
    Set t = New Collection
    For j = 1 To UBound(lMaterialClass)
        Set r = New Collection
        r.Add lMaterialClass(j).id, "ID"
        r.Add lMaterialClass(j).Name, "Name"
        t.Add r
    Next j

    Set MaterialComboBox = t

End Sub

Public Sub GetMaterialGroupCombo(MaterialComboBox As Object)
    Dim i As Integer
    Dim j As Integer
    'return a combobox list for material class

    For j = 1 To UBound(lMaterialGroup)
        MaterialComboBox.AddItem (CStr(j - 1) & " - " & lMaterialGroup(j).Name)
    Next j

End Sub

Public Sub GetMaterialGroupCollection(MaterialComboBox As Collection)
    Dim i As Integer
    Dim j As Integer
    Dim t As Collection
    Set t = New Collection
    'return a combobox list for material class
    For j = 1 To UBound(lMaterialGroup)
        t.Add (CStr(j - 1) & " - " & lMaterialGroup(j).Name)
    Next j
    Set MaterialComboBox = t

End Sub

Public Sub SavetoDatabase(parameters As Collection)
    'save the snapshot of the object to the database
    Dim t As Collection
    
    Dim par1 As adodb.Parameter
    Set par1 = New adodb.Parameter
    par1.Type = adVarChar                      'Parameter type
    par1.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
   ' par1.value = group                 'Parameter value (MaterialClass)
    par1.Size = 150                             'Max. size for Parameter value
    
    Dim par3 As adodb.Parameter
    Set par3 = New adodb.Parameter
    par3.Type = adDouble                      'Parameter type
    par3.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
  '  par3.value = unit                 'Parameter value (MaterialClass)
    'par3.Size = 50                             'Max. size for Parameter value


    Tagdb.ConnectDSN
    If Tagdb.getConnection.State = adStateOpen Then

        For Each t In parameters
            par1.value = t("tagname")
            par3.value = t("value")
            Tagdb.executeCommand2Par "[ui_UpdateParameterValue]", par1, par3
            
        Next
        
        Tagdb.Disconnect
    End If

End Sub

Public Function UnitNotes(object As String) As Collection
    'read notes from database
    Dim temp As Collection
    Dim Col As Collection
    Dim rsa As adodb.Recordset

    Set Col = New Collection
    Dim par1 As adodb.Parameter
    Set par1 = New adodb.Parameter
    par1.Type = adVarChar                      'Parameter type
    par1.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par1.value = object                 'Parameter value (MaterialClass)
    par1.Size = 50                             'Max. size for Parameter value

    Dim par3 As adodb.Parameter
    Set par3 = New adodb.Parameter
    par3.Type = adVarChar                      'Parameter type
    par3.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par3.value = ""                 'Parameter value (MaterialClass)
    par3.Size = 10                             'Max. size for Parameter value


    Tagdb.ConnectDSN
    If Tagdb.getConnection.State = adStateOpen Then
        'Execute Command
        Set rsa = Tagdb.executeCommand2Par("[uiNotesLog]", par1, par3)
        'object type list returned from stored procedure
        While Not rsa.EOF
            Set temp = New Collection
            temp.Add rsa.Fields("DateTimeStamp").value, "DateTimeStamp"
            temp.Add rsa.Fields("OperatedBy").value, "OperatedBy"
            temp.Add rsa.Fields("Comment").value, "Comment"
            Col.Add temp
            rsa.MoveNext
        Wend
        rsa.Close
        Tagdb.Disconnect
    End If
Set UnitNotes = Col
End Function

Public Function GetReceptionParameters(DeliverID As Integer) As Collection
    Dim temp As Collection
    Dim Col As Collection
    Dim rsa As adodb.Recordset
    
    Set Col = New Collection
    Dim par1 As adodb.Parameter
    Set par1 = New adodb.Parameter
    par1.Type = adVarChar                      'Parameter type
    par1.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par1.value = DeliverID                     'Parameter value (DeliverID)
    par1.Size = 30                             'Max. size for Parameter value
    
    Dim par3 As adodb.Parameter
    Set par3 = New adodb.Parameter
    par3.Type = adVarChar                      'Parameter type
    par3.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par3.value = ""                 'Parameter value (MaterialClass)
    par3.Size = 10                             'Max. size for Parameter value
   
    Tagdb.ConnectDSN
    If Tagdb.getConnection.State = adStateOpen Then
        'Execute Command
        Set rsa = Tagdb.executeCommand2Par("[uiReceptionParamGet]", par1, par3)            'Execute command
    
        'parameters returned from stored procedure
        While Not rsa.EOF
            Set temp = New Collection
            temp.Add rsa.Fields("ParameterName").value, "ParameterName"
            temp.Add rsa.Fields("Value").value, "Value"
            Col.Add temp
            rsa.MoveNext
        Wend
        rsa.Close
        Tagdb.Disconnect
    End If
    Set GetReceptionParameters = Col
End Function

Public Sub LoadData()

    'establish connection to TPMDB database
    RaiseEvent StatusChanged("Connecting to TPMDB", "0")
    If debugmode Then LogDiagnosticsMessage "Connecting to TPMDB", ftDiagSeverityInfo
On Error Resume Next
    db.ConnectDSN
    If db.getConnection.State <> adStateOpen Then
        If db.DSN = lDSNTPMDB Then
            Set db = db2
        Else
            Set db = db1
        End If
        db.ConnectDSN
    End If
On Error GoTo ErrHandler
    If db.getConnection.State = adStateOpen Then
        If debugmode Then LogDiagnosticsMessage "Connected to TPMDB", ftDiagSeverityInfo
'        LogDiagnosticsMessage "Connection to TPMDB database active", ftDiagSeverityInfo
        'load local tables with data
        RaiseEvent StatusChanged("Loading User Data", "5")
        If debugmode Then LogDiagnosticsMessage "Loading User Data", ftDiagSeverityInfo
        LoadUserData
        RaiseEvent StatusChanged("Loading Conditions", "10")
        If debugmode Then LogDiagnosticsMessage "Loading Conditions", ftDiagSeverityInfo
        LoadConditionTextsData
        RaiseEvent StatusChanged("Loading Cstat", "20")
        If debugmode Then LogDiagnosticsMessage "Loading Cstat Data", ftDiagSeverityInfo
        LoadCstatData
        RaiseEvent StatusChanged("Loading Estat", "30")
        If debugmode Then LogDiagnosticsMessage "Loading Estat Data", ftDiagSeverityInfo
        LoadEstatData
        RaiseEvent StatusChanged("Loading Phase Modes", "35")
        If debugmode Then LogDiagnosticsMessage "Loading Phase Modes", ftDiagSeverityInfo
        If lAlternateModeTexts = 0 Then
            LoadPhModeData
        End If
        RaiseEvent StatusChanged("Loading Equipment Modes", "36")
        If debugmode Then LogDiagnosticsMessage "Loading Equipment Modes", ftDiagSeverityInfo
        If lAlternateModeTexts = 0 Then
            LoadEqModeData
        End If
        RaiseEvent StatusChanged("Loading QA Result", "37")
        If debugmode Then LogDiagnosticsMessage "Loading QA Result", ftDiagSeverityInfo
        LoadQAResultData
        RaiseEvent StatusChanged("Loading Materials", "40")
        If debugmode Then LogDiagnosticsMessage "Loading Materials Data", ftDiagSeverityInfo
        LoadMaterialData
        RaiseEvent StatusChanged("Loading Material Class", "50")
        If debugmode Then LogDiagnosticsMessage "Loading Material Class", ftDiagSeverityInfo
        LoadMaterialClassData
        RaiseEvent StatusChanged("Loading Yes\No", "55")
        If debugmode Then LogDiagnosticsMessage "Loading Yes\No Data", ftDiagSeverityInfo
        LoadYesNo
        RaiseEvent StatusChanged("TPMDB Complete", "60")
        If debugmode Then LogDiagnosticsMessage "TPMDB Complete", ftDiagSeverityInfo
        'close database connection
        db.Disconnect
    Else
         LogDiagnosticsMessage "Connection to TPMDB database not active", ftDiagSeverityWarning
    End If
    
    RaiseEvent StatusChanged("Connecting to TagDB", "70")
    If debugmode Then LogDiagnosticsMessage "Connecting to TagDB", ftDiagSeverityInfo
    'establish connection to TagDB database
On Error Resume Next
    Tagdb.ConnectDSN
    If Tagdb.getConnection.State <> adStateOpen Then
        If Tagdb.DSN = lDSNTagDB Then
            Set Tagdb = Tagdb2
        Else
            Set Tagdb = Tagdb1
        End If
        Tagdb.ConnectDSN
    End If
    
On Error GoTo ErrHandler
    If Tagdb.getConnection.State = adStateOpen Then
        If debugmode Then LogDiagnosticsMessage "Connected to TagDB", ftDiagSeverityInfo
'        LogDiagnosticsMessage "Connection to TagDB database active", ftDiagSeverityInfo
        'load local tables with data
        RaiseEvent StatusChanged("Loading CIP Recipe", "75")
        If debugmode Then LogDiagnosticsMessage "Loading CIP Recipe", ftDiagSeverityInfo
        LoadCIPRecipeData
        LoadCIPStandaloneRecipeData
        RaiseEvent StatusChanged("Loading CIP Activities", "76")
        If debugmode Then LogDiagnosticsMessage "Loading CIP Activities", ftDiagSeverityInfo
        LoadCIPActivity
        RaiseEvent StatusChanged("Loading CIP Phases", "77")
        If debugmode Then LogDiagnosticsMessage "Loading CIP Phases", ftDiagSeverityInfo
        LoadCCPhaseList
        RaiseEvent StatusChanged("Loading Parameters", "80")
        If debugmode Then LogDiagnosticsMessage "Loading Parameters", ftDiagSeverityInfo
        LoadAreas
        RaiseEvent StatusChanged("Loading Phase Modes", "81")
        If debugmode Then LogDiagnosticsMessage "Loading Phase Modes", ftDiagSeverityInfo
        If lAlternateModeTexts = 1 Then
            LoadPhModeData_Alternate
        End If
        RaiseEvent StatusChanged("Loading Equipment Modes", "82")
        If debugmode Then LogDiagnosticsMessage "Loading Equipment Modes", ftDiagSeverityInfo
        If lAlternateModeTexts = 1 Then
            LoadEqModeData_Alternate
        End If
        RaiseEvent StatusChanged("Loading Phase Steps", "83")
        If debugmode Then LogDiagnosticsMessage "Loading Phase Steps", ftDiagSeverityInfo
        LoadPhaseSteps
        RaiseEvent StatusChanged("Loading CMIS", "90")
        If debugmode Then LogDiagnosticsMessage "Loading CMIS Data", ftDiagSeverityInfo
        LoadCMIS
        RaiseEvent StatusChanged("Loading Material Groups", "95")
        If debugmode Then LogDiagnosticsMessage "Loading Material Groups", ftDiagSeverityInfo
        LoadMaterialGroupData
        'close database connection
        Tagdb.Disconnect
    Else
         LogDiagnosticsMessage "Connection to TagDB database not active", ftDiagSeverityWarning
    End If
    RaiseEvent StatusChanged("Complete", "100")
    If debugmode Then LogDiagnosticsMessage "Load Data complete", ftDiagSeverityInfo
Exit Sub
    'Error message
ErrHandler:
    
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on LocalDatabase LoadData", ftDiagSeverityError
End Sub


Public Property Get DSNTagDB() As String
    DSNTagDB = lDSNTagDB
End Property

Public Property Let DSNTagDB(ByVal value As String)
    lDSNTagDB = value
End Property

Public Property Get DSNTagDB2() As String
    DSNTagDB2 = lDSNTagDB2
End Property

Public Property Let DSNTagDB2(ByVal value As String)
    lDSNTagDB2 = value
End Property

Public Property Get DSNTPMDB() As String
    DSNTPMDB = lDSNTPMDB
End Property

Public Property Let DSNTPMDB(ByVal value As String)
    lDSNTPMDB = value
End Property

Public Property Get DSNTPMDB2() As String
    DSNTPMDB2 = lDSNTPMDB2
End Property

Public Property Let DSNTPMDB2(ByVal value As String)
    lDSNTPMDB2 = value
End Property

Public Property Get ConnectionType() As Integer
    ConnectionType = lconnectiontype
End Property

Public Property Let ConnectionType(ByVal value As Integer)
    lconnectiontype = value
End Property

Public Property Get AlternateModeTexts() As Integer
    AlternateModeTexts = lAlternateModeTexts
End Property

Public Property Let AlternateModeTexts(ByVal value As Integer)
    lAlternateModeTexts = value
End Property

Public Property Get EnableDebug() As Boolean
    EnableDebug = debugmode
End Property

Public Property Let EnableDebug(ByVal value As Boolean)
    debugmode = value
End Property


Public Sub CreateReportUser(Name As String)

On Error Resume Next
    Dim rs As adodb.Recordset
    Dim plcid As Integer
    Dim par1 As adodb.Parameter
    Set par1 = New adodb.Parameter
    par1.Type = adInteger                      'Parameter type
    par1.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par1.value = -1                 'Parameter value (MaterialClass)
    par1.Size = 50                             'Max. size for Parameter value
    
    Dim par1a As adodb.Parameter
    Set par1a = New adodb.Parameter
    par1a.Type = adInteger                      'Parameter type
    par1a.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par1a.value = -1                 'Parameter value (MaterialClass)
    par1a.Size = 50
    db.ConnectDSN

    If db.getConnection.State = adStateOpen Then

        'Execute Command
        'get new number for plcid
        Set rs = db.executeCommand2Par("[naUsrMgr_evaluatePLC_ID]", par1, par1a)
        If Not rs.EOF Then
            plcid = rs(0).value
        End If

        Dim par2 As adodb.Parameter
        Set par2 = New adodb.Parameter
        par2.Type = adVarChar                      'Parameter type
        par2.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
        par2.value = "A"                 'Parameter value (MaterialClass)
        par2.Size = 1                             'Max. size for Parameter value
        
        Dim par3 As adodb.Parameter
        Set par3 = New adodb.Parameter
        par3.Type = adVarChar                      'Parameter type
        par3.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
        par3.value = Name                 'Parameter value (MaterialClass)
        par3.Size = 30                             'Max. size for Parameter value
        
        Dim Par3a As adodb.Parameter
        Set Par3a = New adodb.Parameter
        Par3a.Type = adVarChar                      'Parameter type
        Par3a.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
        Par3a.value = Name                 'Parameter value (MaterialClass)
        Par3a.Size = 30
        
        Dim par4 As adodb.Parameter
        Set par4 = New adodb.Parameter
        par4.Type = adInteger                      'Parameter type
        par4.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
        par4.value = plcid                 'Parameter value (MaterialClass)
        par4.Size = 50                             'Max. size for Parameter value
        
        Dim par4a As adodb.Parameter
        Set par4a = New adodb.Parameter
        par4a.Type = adInteger                      'Parameter type
        par4a.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
        par4a.value = plcid                 'Parameter value (MaterialClass)
        par4a.Size = 50
        
        Dim par5 As adodb.Parameter
        Set par5 = New adodb.Parameter
        par5.Type = adInteger                      'Parameter type
        par5.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
        par5.value = 1                 'Parameter value (MaterialClass)
        par5.Size = 50                             'Max. size for Parameter value
        
        Dim par6 As adodb.Parameter
        Set par6 = New adodb.Parameter
        par6.Type = adVarChar                      'Parameter type
        par6.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
        par6.value = ""                 'Parameter value (MaterialClass)
        par6.Size = 20                             'Max. size for Parameter value
        
        Dim par6a As adodb.Parameter
        Set par6a = New adodb.Parameter
        par6a.Type = adVarChar                      'Parameter type
        par6a.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
        par6a.value = ""                 'Parameter value (MaterialClass)
        par6a.Size = 50
        
        Dim par6b As adodb.Parameter
        Set par6b = New adodb.Parameter
        par6b.Type = adVarChar                      'Parameter type
        par6b.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
        par6b.value = ""                 'Parameter value (MaterialClass)
        par6b.Size = 50
        
        Dim par7 As adodb.Parameter
        Set par7 = New adodb.Parameter
        par7.Type = adDBTimeStamp                      'Parameter type
        par7.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
        par7.value = Now                 'Parameter value (MaterialClass)
        par7.Size = 20
        
        Dim par8 As adodb.Parameter
        Set par8 = New adodb.Parameter
        par8.Type = adVarChar                      'Parameter type
        par8.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
        par8.value = ""                 'Parameter value (MaterialClass)
        par8.Size = 50
        
'        @ADU_Text                       AS nVarchar(1),
'        @ConfiguredBy                   AS nVarchar(30),
'        @Users_ID                       AS Int,
'        @Users_PLC_ID                   AS Int,
'        @Users_Name                     AS nVarchar(30),
'        @Users_Active                   AS Bit,
'        @Users_Password                 AS nVarchar(20),
'        @Users_Site                     AS nVarchar(50),
'        @Users_Email                    AS nVarchar(50),
'        @Users_LastSaved                AS DateTime,
'        @Users_CustomUrl                AS nVarchar(1000)
        
        'create new user
        db.executeCommand11Par "[naUsrMgr_aduUser]", par2, par3, par1, par4, Par3a, par4a, par6, par6a, par6b, par7, par8
        
        'reload user list
        LoadUserData
        db.Disconnect

    End If
Exit Sub
    'Error message
ErrHandler:
    
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on LocalDatabase CreateReportUser", ftDiagSeverityError
End Sub



Public Function RecipeXML(rid As String, version As String) As String
    'insert the note to db

    Dim par1 As adodb.Parameter
    Set par1 = New adodb.Parameter
    par1.Type = adVarChar                      'Parameter type
    par1.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par1.value = rid                 'Parameter value (MaterialClass)
    par1.Size = 25                             'Max. size for Parameter value
    
    Dim par2 As adodb.Parameter
    Set par2 = New adodb.Parameter
    par2.Type = adVarChar                      'Parameter type
    par2.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par2.value = version                 'Parameter value (MaterialClass)
    par2.Size = 10                             'Max. size for Parameter value
    
    Dim rsa As adodb.Recordset
    db.ConnectDSN
    If db.getConnection.State = adStateOpen Then
        'Execute Command
        Set rsa = db.executeCommand2Par("[TPIBK_getRecipeXML]", par1, par2)
        If Not rsa.EOF Then
          RecipeXML = rsa.Fields(0).value
        End If
        rsa.Close

        db.Disconnect
    End If

End Function

Public Function GetCMdescription(Tagname As String)
On Error Resume Next 'Error Check

    'Declare local variables
    Dim c As Collection

    For Each c In lCMIS
        If c("DisplayName") = Tagname Then
            If IsNull(c("Description")) Then
                GetCMdescription = ""
            Else
                GetCMdescription = c("Description")
            End If
            Exit Function
        End If
    Next c
    
    GetCMdescription = "Not found in TagDB"

End Function

Public Sub GetRecipeParameternames(rid As String, version As String, step As String, ret As Collection)
    
    Dim sqlstring As String
    Dim temp As Collection
    Dim Col As Collection
    Set Col = New Collection
    
    sqlstring = "SELECT Description, REPLACE(Name, ' ', '_') AS Name " + _
                    "FROM dbo.v_TPIBK_StepParameterNames " + _
                    "WHERE        (Recipe_RID = N'" & rid & "') AND (Recipe_Version = N'" & version & "') AND (Step = '" & step & "') " + _
                    "ORDER BY TPIBK_RecipeParameters_ID"
    
    Dim rsa As adodb.Recordset
    db.ConnectDSN
    If db.getConnection.State = adStateOpen Then
        'Execute Command
        Set rsa = db.getRecords(sqlstring)
       While Not rsa.EOF
            Set temp = New Collection
            temp.Add rsa.Fields("Name").value, "Name"
            temp.Add rsa.Fields("Description").value, "Description"
            Col.Add temp
            rsa.MoveNext
        Wend
        rsa.Close
        Set ret = Col
        db.Disconnect
    End If

End Sub

Public Sub EditMaterial(id As Integer, Name As String, ClassID As Integer, MaterialType As Integer, Alias As String, user As String, Action As Integer)
    
    db.ConnectDSN
    If db.getConnection.State = adStateOpen Then
        Dim par1 As adodb.Parameter
        Set par1 = New adodb.Parameter
        par1.Type = adInteger
        par1.Direction = adParamInput
        par1.value = id
        par1.Size = 30
        
        Dim par2 As adodb.Parameter
        Set par2 = New adodb.Parameter
        par2.Type = adVarChar
        par2.Direction = adParamInput
        par2.value = Name
        par2.Size = 30
        
        Dim par3 As adodb.Parameter
        Set par3 = New adodb.Parameter
        par3.Type = adInteger
        par3.Direction = adParamInput
        par3.value = ClassID
        par3.Size = 30

        Dim par4 As adodb.Parameter
        Set par4 = New adodb.Parameter
        par4.Type = adInteger
        par4.Direction = adParamInput
        par4.value = MaterialType
        par4.Size = 30
        
        Dim par5 As adodb.Parameter
        Set par5 = New adodb.Parameter
        par5.Type = adVarChar
        par5.Direction = adParamInput
        par5.value = Alias
        par5.Size = 50
        
        Dim par6 As adodb.Parameter
        Set par6 = New adodb.Parameter
        par6.Type = adVarChar
        par6.Direction = adParamInput
        par6.value = user
        par6.Size = 30
        
        Dim par7 As adodb.Parameter
        Set par7 = New adodb.Parameter
        par7.Type = adInteger
        par7.Direction = adParamInput
        par7.value = Action
        par7.Size = 30

        'Execute setup Command with Parameters
        db.executeCommand7Par "ui_ModifyMaterial", par1, par2, par3, par4, par5, par6, par7          'Execute command

        LoadMaterialClassData
        db.Disconnect
        
    End If
End Sub


Public Sub EditMatClass(id As Integer, Name As String, user As String, Action As Integer)

    db.ConnectDSN
    If db.getConnection.State = adStateOpen Then
        Dim par1 As adodb.Parameter
        Set par1 = New adodb.Parameter
        par1.Type = adInteger
        par1.Direction = adParamInput
        par1.value = id
        par1.Size = 30
        
        Dim par2 As adodb.Parameter
        Set par2 = New adodb.Parameter
        par2.Type = adVarChar
        par2.Direction = adParamInput
        par2.value = Name
        par2.Size = 30
        
        Dim par3 As adodb.Parameter
        Set par3 = New adodb.Parameter
        par3.Type = adInteger
        par3.Direction = adParamInput
        par3.value = Action
        par3.Size = 30
        
        Dim par6 As adodb.Parameter
        Set par6 = New adodb.Parameter
        par6.Type = adVarChar
        par6.Direction = adParamInput
        par6.value = user
        par6.Size = 30

        'Execute setup Command with Parameters
        db.executeCommand4Par "ui_ModifyMaterialClass", par1, par2, par6, par3             'Execute command

        LoadMaterialClassData
        db.Disconnect
        
    End If

End Sub

Public Function GetSiteMaterialAlias(plcid As Long) As String 'Added AKB 8/20/21
    Dim i As Integer
    Dim temp As String
    'initialize to not found
    temp = "ID: " & CStr(plcid) & " missing"
    For i = 1 To UBound(lMaterial)
        If lMaterial(i).PLC_ID = plcid Then
            temp = lMaterial(i).SiteAlias
            GoTo found
        End If
    Next i
found:
    GetSiteMaterialAlias = temp
End Function

Public Function GetMaterialEU(plcid As Long) As String 'ND_20230109
    Dim i As Integer
    Dim temp As String
    'initialize to not found
    temp = "ID: " & CStr(plcid) & " missing"
    For i = 1 To UBound(lMaterial)
        If lMaterial(i).PLC_ID = plcid Then
            temp = lMaterial(i).EU
            GoTo found
        End If
    Next i
found:
    GetMaterialEU = temp
End Function
