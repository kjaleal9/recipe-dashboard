Option Explicit
'              Tetra PlantMaster
'
'Function:     Footer.gfx
'v6.4.00_00
'Created by:   Carey Warren
'------------------------------------------------------------------------------
Private Tag_Group As TagGroup
Private TagsInError As StringList
Private MaterialTag_Group As TagGroup
Private WithEvents UserLogin As DisplayClient.Application
Private WithEvents ApplicationStatus As Application
Private UserID_TPMDB As Integer
Private CurrentDisplayname As String
Private LastDisplayname As String
Private WithEvents ld As LocalDatabase
Private lParameterTypetoCopy As String
Private lParameterUnittoCopy As String
Private lParameterTagtoCopy As String
Private lCommissioningMode As Boolean
Public LastTrendTemplatename As String
Public lUserCodes As String 'Added for TPIBK
Private Splash As Display
Private lDisplayLev2ParKeypad As Integer
Private lParameterGrouptoCopy As String

Public Computername As String
Private PIWebServername As String

Public HMIServername As String
Public AlarmServername As String
Public DataServername As String

Private lUsername As String
Const DB_SingleSQL = 1 'standard install, 1 sql server
Const DB_RedundantSQL = 2 '2 SQL servers for tagdb and TPMDB

Private AlternateModeTexts As Integer

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassname As String, ByVal lpWindowname As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim monitor As Integer
Private lhistorian As String
Private debugmode As Boolean
Private lcippath As String


Public Sub SetupDatabase()
 'Error check
    On Error GoTo ErrHandler
    If Me.TagParameters(1) = 1 Then
        RSView32SEIdleDetectControl.Enabled = True 'Disable if autologout not required
    Else
        RSView32SEIdleDetectControl.Enabled = False 'Disable if autologout not required
    End If
    RSView32SEIdleDetectControl.Interval = Me.TagParameters(2) 'seconds
    RSView32SEIdleDetectControl.Visible = RSView32SEIdleDetectControl.Enabled
'check which source to use for equipment and phase mode texts
'0=PI model in TPMDB
'1=Alternate tables in TagDB

    On Error GoTo AlternateModeTextserr
    AlternateModeTexts = Me.TagParameters(14)
    GoTo AlternateModeTextsok
AlternateModeTextserr:
    AlternateModeTexts = 0
AlternateModeTextsok:

    'Check to see if the localdatabase has been configured
    If ld Is Nothing Then
        Set ld = New LocalDatabase
    End If
    
    ld.ConnectionType = CInt(Me.TagParameters(3)) ' single or redundant database
    'Use 32 bit DSN Connections
    
    ld.DSNTagDB = CStr(Me.TagParameters(4)) 'Required
    ld.DSNTPMDB = CStr(Me.TagParameters(5)) 'Required
    ld.DSNTagDB2 = CStr(Me.TagParameters(6)) 'Optional
    ld.DSNTPMDB2 = CStr(Me.TagParameters(7)) 'Optional
    ld.AlternateModeTexts = AlternateModeTexts
    ld.EnableDebug = debugmode
    If debugmode Then LogDiagnosticsMessage "DSN TagDB= " & ld.DSNTagDB, ftDiagSeverityInfo
    If debugmode Then LogDiagnosticsMessage "DSN TPMDB= " & ld.DSNTPMDB, ftDiagSeverityInfo
          
    ld.Initialize
    ld.LoadData

    PIWebServername = Me.TagParameters(8) 'ip address or name of PI web server
    
    'enable keyboard
    lDisplayLev2ParKeypad = Me.TagParameters(9)
    
    'cip recipe path
    lcippath = Me.TagParameters(16)
    

    
Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " SetupDatabase", ftDiagSeverityError
End Sub

Public Property Get PIWebAddress() As String

    'Error check
    On Error GoTo ErrHandler
    
    PIWebAddress = PIWebServername

Exit Property

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " PIWebAddress", ftDiagSeverityError
End Property


Public Property Get UserID() As Integer

    'Error check
    On Error GoTo ErrHandler
    
    UserID = UserID_TPMDB

Exit Property

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " UserID", ftDiagSeverityError
End Property


Private Sub Button10_Released()
    'Abort Me;display splash; Display Footer /pfooter
    executeCommand "Abort Me;display Splash /CC;Display Footer /P" & Me.ParameterFileName & " /M" & CStr(monitor)
End Sub

Private Sub Button6_Released()
    executeCommand "display Splash /CC"
    ld.LoadData
End Sub

Private Sub CalcButton_Released()
 'Error check

    'Start Application
    FnSetForegroundWindow "Calculator", "AppStart calc"


End Sub

Private Sub Display_AnimationStart()
    
    'Error check
    On Error GoTo ErrHandler

    'Declare local variables
    Dim Username As String
    Dim strComputername As String
    debugmode = False
    If Me.TagParameters(10) = 1 Then
        debugmode = True
        LogDiagnosticsMessage "Debug Mode Active", ftDiagSeverityInfo
    End If
    
    SetupDatabase
    Set ApplicationStatus = Me.Application
    Call ModSetForegroundWindow.PublishClientHandle(ModSetForegroundWindow.GetClientName(Application.ConfigurationFileName), Application.WindowHandle)
        
      'Check security
    Set UserLogin = Me.Application
    Username = CurrentUserName
    Call UserLogin_Login(Username)
    
    ' Get the name of the client computer or if terminals services get the session name
    strComputername = Environ("SESSIONname")

    If strComputername = "Console" Then
        strComputername = Environ("COMPUTERname")
    End If
    
    Computername = strComputername
    HMIServername = Me.TagParameters(11)
    DataServername = Me.TagParameters(12)
    AlarmServername = Me.TagParameters(13)
    lParameterGrouptoCopy = ""
    lhistorian = Me.TagParameters(15)
    
    If Me.Left > 0 Then
        monitor = 2
    Else
        monitor = 1
    End If

Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " Display_AnimationStart", ftDiagSeverityError
End Sub

Private Sub LabButton_Released()
    WebButton "http://" & PIWebServername & "/Navigator/LoginForm.aspx?name=QCResult"
End Sub

Private Sub ld_StatusChanged(status As String, value As Integer)
    'Error check
    On Error GoTo ErrHandler
    Dim d As Display
    For Each d In Me.Application.LoadedDisplays
        If UCase(d.Name) = UCase("Splash") Then
            Set Splash = d
            Call Splash.ChangeStatus(status, value)
            Set Splash = Nothing
        End If
    Next d
Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " ld_StatusChanged", ftDiagSeverityError
End Sub

Private Sub LogoutButton_Released()
     'Error check
    On Error GoTo ErrHandler
    executeCommand "display Splash /CC"
        ld.LoadData
    executeCommand "Login Default password"
Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " LogoutButton_Released", ftDiagSeverityError
End Sub

Private Sub PIButton_Released()
    WebButton "http://" & PIWebServername & "/Navigator/"
End Sub

Private Sub RSView32SEIdleDetectControl_EnterIdleState()
    
    'Error check
On Error GoTo ErrHandler
    executeCommand "display Splash /CC"
    ld.LoadData
    executeCommand "Login Default password"

Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " RSView32SEIdleDetectControl_EnterIdleState", ftDiagSeverityError
End Sub

Private Sub UserLogin_Login(ByVal Username As String)
    
    'Error check
    On Error GoTo ErrHandler
    lUserCodes = CurrentUserCodes
    'Get the user id from TPMDB for the current user
    'Call GetUserID_FromTPMDB
    lUsername = Username
    If Not ld Is Nothing Then
        UserID_TPMDB = ld.GetUserID(Username)
    End If
    If Me.TagParameters(1) = 1 Then
        If Username = "Default" Then
            RSView32SEIdleDetectControl.Enabled = False
        Else
            RSView32SEIdleDetectControl.Enabled = True
        End If
        RSView32SEIdleDetectControl.Visible = RSView32SEIdleDetectControl.Enabled
    End If
     'user doesnt exist in reporting ask to create
    If (UserID_TPMDB = 999) And (Username <> "Default") Then
        Dim r As Integer
        r = MsgBox("User does not exist in reporting systems. Create user?", vbYesNo)
        If r = vbYes Then
            If Not ld Is Nothing Then
                ld.CreateReportUser (Username)
                UserID_TPMDB = ld.GetUserID(Username)
                
                If UserID_TPMDB = 999 Then
                    MsgBox "Error Creating User.  Please contact the system administrator.", vbExclamation
                Else
                    MsgBox "User Created", vbInformation
                End If
                
            End If
        End If
    End If
Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " UserLogin_Login", ftDiagSeverityError
    'UserIDSpinButton.value = UserID_TPMDB
End Sub

Private Sub UserLogin_Logout(ByVal Username As String)
 'Error check
    On Error GoTo ErrHandler
    
    'Get the user id from TPMDB for the current user
    If Username <> "Default" Then
        ld.LoadData
    End If
Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " UserLogin_Logout", ftDiagSeverityError
End Sub

Public Sub GetMaterialList_FromTPMDB(MaterialComboBox As Object, MaterialID As Long, Materialclass As String)
    
    'Error check
    On Error GoTo ErrHandler
    'load the combobox with values from the local database
    ld.GetMaterialCombo MaterialComboBox, MaterialID, Materialclass

Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetMaterialList_FromTPMDB", ftDiagSeverityError
End Sub

Public Sub GetMaterialClassCombo(MaterialComboBox As Object)

    'Error check
    On Error GoTo ErrHandler
    'load the combobox with values from the local database
    ld.GetMaterialClassCombo MaterialComboBox

Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetMaterialClassCombo", ftDiagSeverityError
End Sub

Public Function Materialname_FromTPMDB(MaterialID As Long) As String
    
    'Error check
    On Error GoTo ErrHandler
    'load the material name from the local database
    Materialname_FromTPMDB = ld.GetMaterialName(MaterialID)
    
Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetMaterialList_FromTPMDB", ftDiagSeverityError
End Function

Public Function SiteMaterialAlias2_FromTPMDB(SiteMaterialAlias As String) As Integer
    'Error check
    On Error GoTo ErrHandler
    'load the material name from the local database
       
    SiteMaterialAlias2_FromTPMDB = ld.GetSiteMaterialAlias2(SiteMaterialAlias)
    
Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetMaterialList_FromTPMDB", ftDiagSeverityError
End Function

Public Function Stepname_FromTPMDB(Tagname As String, step As String, Optional LikeSearch As Boolean = False) As String
    
    'Error check
    On Error GoTo ErrHandler
    'load the step name from the local database
    If LikeSearch Then
        Stepname_FromTPMDB = ld.GetPhaseStep_LikeSearch(Tagname, step)
    Else
        Stepname_FromTPMDB = ld.GetPhaseStep(Tagname, step)
    End If
    
Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " Stepname_FromTPMDB", ftDiagSeverityError
End Function


Public Function GetMaterialList(Materialclass As String) As Collection
    
    'Error check
    On Error GoTo ErrHandler
    'load the material name from the local database
    Set GetMaterialList = ld.GetMaterialList(Materialclass)
    
Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetMaterialList", ftDiagSeverityError
End Function

Public Function GetStatusList(StatusType As String) As Collection
    
    'Error check
    On Error GoTo ErrHandler
    'load the material name from the local database
    Set GetStatusList = ld.GetStatusList(StatusType)
    
Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetStatusList", ftDiagSeverityError
End Function

Public Function GetModeList(StatusType As String, SelectionValue As String) As Collection
    
    'Error check
    On Error GoTo ErrHandler
    'load the material name from the local database
    Set GetModeList = ld.GetModeList(StatusType, SelectionValue)
    
Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetModeList", ftDiagSeverityError
End Function

Public Function GetEquipmentMode(plcid As Long, unit As String) As String
    
    'Error check
    On Error GoTo ErrHandler
    'load the CIP Recipe name from the local database
    GetEquipmentMode = ld.GetEquipmentMode(plcid, unit)
    
Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetEquipmentMode", ftDiagSeverityError
End Function

Public Function GetPhaseMode(plcid As Long, phase As String) As String
    
    'Error check
    On Error GoTo ErrHandler
    'load the CIP Recipe name from the local database
    GetPhaseMode = ld.GetPhaseMode(plcid, phase)
    
Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetPhaseMode", ftDiagSeverityError
End Function

Public Function CIPRecipename_FromTPMDB(RecipeID As Long) As String
    
    'Error check
    On Error GoTo ErrHandler
    'load the CIP Recipe name from the local database
    CIPRecipename_FromTPMDB = ld.GetCIPRecipeName(RecipeID)
    
Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " CIPRecipename_FromTPMDB", ftDiagSeverityError
End Function

Public Function CIPStandaloneRecipename_FromTPMDB(RecipeID As Long) As String
    
    'Error check
    On Error GoTo ErrHandler
    'load the CIP Recipe name from the local database
    CIPStandaloneRecipename_FromTPMDB = ld.GetCIPStandaloneRecipeName(RecipeID)
    
Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " CIPStandaloneRecipename_FromTPMDB", ftDiagSeverityError
End Function

Public Sub FillComboBox(ComboBox As Object, status As Long, StatusType As String)
    
    'Error check
    On Error GoTo ErrHandler
    'load the status combo from the local database
    ld.GetStatusCombo ComboBox, status, StatusType
    
Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " FillComboBox", ftDiagSeverityError
End Sub



Public Sub FillCMListBox(ListBox() As Variant, CMType As String, Filter As String)
    
    'Error check
    On Error GoTo ErrHandler
    
    'load the type combo from the local database
    ld.GetCMListBox ListBox, CMType, Filter
    
Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " FillCMListBox", ftDiagSeverityError
End Sub


Public Sub SavetoDatabase(parameters As Collection)

    'Error check
    On Error GoTo ErrHandler

    'load the snapshot from the local database
    ld.SavetoDatabase parameters

Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " SavetoDatabase", ftDiagSeverityError
End Sub


Public Function ConditionList() As String()
    
    'Error check
    On Error GoTo ErrHandler
    'find the status name in the local database
    ConditionList = ld.GetConditionList
    
Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " ConditionList", ftDiagSeverityError
End Function

Public Function GetAreasList() As Collection
    
    'Error check
    On Error GoTo ErrHandler
    'find the status name in the local database
    Set GetAreasList = ld.GetAreasList
    
Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetAreasList", ftDiagSeverityError
End Function


Public Function Statusname_FromFile(status As Long, StatusType As String) As String
    
    'Error check
    On Error GoTo ErrHandler
    'find the status name in the local database
    Statusname_FromFile = ld.GetStatusName(status, StatusType)
    
Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " Statusname_FromFile", ftDiagSeverityError
End Function

Public Function GetParameters(area As String, unit As String, Optional phase As String = "none") As Collection

    'Error check
    On Error GoTo ErrHandler
    'find the parameter list in the local database
    Set GetParameters = ld.GetParameters(area, unit, phase)

Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetParameters", ftDiagSeverityError
End Function

Public Function GetReceptionParameters(DeliveryID As Integer) As Collection
    
    'Error check
    On Error GoTo ErrHandler
    'load the material name from the local database
    Set GetReceptionParameters = ld.GetReceptionParameters(DeliveryID)
    
Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetReceptionParameters", ftDiagSeverityError
End Function


Public Sub ReloadSelectedParameters(area As String, unit As String, group As String, Optional phase As String = "none")

    'Error check
    On Error GoTo ErrHandler
    'find the parameter list in the local database
    ld.ReloadSelectedParameters area, unit, group, phase

Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " ReloadSelectedParameters", ftDiagSeverityError
End Sub

Public Function GetCIPRecipes() As String()

    'Error check
    On Error GoTo ErrHandler
    'return list of cip recipes
    GetCIPRecipes = ld.GetCIPRecipeList

Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetCIPRecipes", ftDiagSeverityError
End Function

Public Function GetCIPStandaloneRecipes() As Collection

    'Error check
    On Error GoTo ErrHandler
    'return list of cip recipes
    Set GetCIPStandaloneRecipes = ld.GetCIPStandaloneRecipeList

Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetCIPStandaloneRecipes", ftDiagSeverityError
End Function

Public Function getEnabledCIPRecipeList(EnabledRecipes As Long) As Collection

    'Error check
    On Error GoTo ErrHandler
    'return list of cip recipes
    Set getEnabledCIPRecipeList = ld.getEnabledCIPRecipeList(EnabledRecipes)

Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " getEnabledCIPRecipeList", ftDiagSeverityError
End Function

Public Sub WriteOperatedByID(PLCTag As String)
    
    'Error check
    On Error GoTo ErrHandler
    
    'Define taggroup
    Set Tag_Group = Application.CreateTagGroup(Me.AreaName, 250)
    Tag_Group.Add PLCTag
    Tag_Group.Active = True
    If Not Tag_Group.RefreshFromSource(TagsInError) Then
        LogDiagnosticsMessage "Taggroup refresh failed on display " & Name
    End If

    executeCommand (Tag_Group.Item(PLCTag).Name & " = " & UserID_TPMDB)
    
    'Close tag group
    Tag_Group.Active = False
    Tag_Group.RemoveAll
    Set Tag_Group = Nothing

Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " WriteOperatedByID", ftDiagSeverityError
End Sub

Private Sub Display_AfterAnimationStop()

    'Error check
    On Error GoTo ErrHandler
    
    'Clear
    Set UserLogin = Nothing
    
Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " Display_AfterAnimationStop", ftDiagSeverityError
End Sub


Private Sub WebButton(url As String)

    'Error check
    On Error GoTo ErrHandler

    Dim logintitle As String
    Dim navigatortitle As String
    Dim tpmpetitle As String
    Dim runbatchstandalonetitle As String
    Dim batchconsoletitle As String
    Dim runbatchtitle As String
    Dim plcpumptitle As String
    Dim qatitle As String

    Dim loginfound As Long
    Dim navigatorfound As Long
    Dim tpmpefound As Long
    Dim runbatchstandalonefound As Long
    Dim batchconsolefound As Long
    Dim runbatchfound As Long
    Dim plcpumpfound As Long
    Dim qafound As Long

    'name of PI windows
    logintitle = "Tetra PlantMaster Login - Internet Explorer"
    qatitle = "Tetra PlantMaster Login Quality Control Result - Internet Explorer"
    navigatortitle = "Tetra PlantMaster Navigator - Internet Explorer"
    tpmpetitle = "TPM PE - Internet Explorer"
    runbatchstandalonetitle = "RunBatchStandalone - Internet Explorer"
    batchconsoletitle = "BatchConsole - Internet Explorer"
    runbatchtitle = "RunBatch - Internet Explorer"
    plcpumptitle = "PlcPump Console - Internet Explorer"

    'window id if found
    loginfound = FindWindow(vbNullString, logintitle)
    navigatorfound = FindWindow(vbNullString, navigatortitle)
    tpmpefound = FindWindow(vbNullString, tpmpetitle)
    runbatchstandalonefound = FindWindow(vbNullString, runbatchstandalonetitle)
    batchconsolefound = FindWindow(vbNullString, batchconsoletitle)
    runbatchfound = FindWindow(vbNullString, runbatchtitle)
    plcpumpfound = FindWindow(vbNullString, plcpumptitle)
    qafound = FindWindow(vbNullString, qatitle)

    'no windows open
    If (loginfound = 0) And (navigatorfound = 0) And (tpmpefound = 0) And (qafound = 0) Then
        'open login form if base forms not open
        'executeCommand "AppStart " & url
        FnSetForegroundWindow logintitle, "AppStart " & url
       

    End If

    If (loginfound > 0) Then
       FnSetForegroundWindow logintitle, vbNullString
    End If

    If (navigatorfound > 0) Then
        FnSetForegroundWindow navigatortitle, vbNullString
    End If

    If (tpmpefound > 0) Then
        FnSetForegroundWindow tpmpetitle, vbNullString
    End If
    'misc windows
    If (runbatchstandalonefound > 0) Then
        FnSetForegroundWindow runbatchstandalonetitle, vbNullString
    End If

    If (batchconsolefound > 0) Then
        FnSetForegroundWindow batchconsoletitle, vbNullString
    End If

    If (runbatchfound > 0) Then
        FnSetForegroundWindow runbatchtitle, vbNullString
    End If

    If (plcpumpfound > 0) Then
        FnSetForegroundWindow plcpumptitle, vbNullString
    End If
    
    If (qafound > 0) Then
        FnSetForegroundWindow qatitle, vbNullString
    End If

Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " NavigatorButton_Released", ftDiagSeverityError
End Sub

Public Function GetComputername()

    'Error check
    On Error GoTo ErrHandler
    
    GetComputername = Computername

Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetComputername", ftDiagSeverityError
End Function

Public Function GetHistorianname()

    'Error check
    On Error GoTo ErrHandler
    
    GetHistorianname = lhistorian

Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetHistorianname", ftDiagSeverityError
End Function

Public Sub Navigation(Displayname As String)
    
    'Error check
    On Error GoTo ErrHandler
    
    'Update variable for Previous button and close last diplay
    If CurrentDisplayname <> Displayname Then
        LastDisplayname = CurrentDisplayname
        CurrentDisplayname = Displayname
    End If

Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " Navigation", ftDiagSeverityError
End Sub


Public Property Get CommissioningMode() As Boolean
    CommissioningMode = lCommissioningMode
End Property

Public Property Let CommissioningMode(ByVal vNewValue As Boolean)
    lCommissioningMode = vNewValue
End Property

Public Sub FillBox(sql As String, Box As Object, fieldlist As String, valuetofind)
'On Error Resume Next 'Error Check
    'Declare local variables
    Dim rs As adodb.Recordset
    Dim db As dbConnect
    
    Set db = ld.db
    db.ConnectDSN
    ListRefreshSQL rs, ld.db, sql, Box, fieldlist, valuetofind
    Set db = Nothing
End Sub

Public Sub GetMaterialnamesCombo(MaterialComboBox As Object)

    'Error check
    On Error GoTo ErrHandler
    'load the combobox with values from the local database
    ld.GetMaterialGroupCombo MaterialComboBox

Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetMaterialnamesCombo", ftDiagSeverityError
End Sub

Public Sub GetMaterialGroupCollection(MaterialComboBox As Collection)

    'Error check
    On Error GoTo ErrHandler
    'load the combobox with values from the local database
    ld.GetMaterialGroupCollection MaterialComboBox

Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetMaterialGroupCollection", ftDiagSeverityError
End Sub


Public Sub GetMaterialClassCollection(MaterialComboBox As Collection)

    'Error check
    On Error GoTo ErrHandler
    'load the combobox with values from the local database
    ld.GetMaterialClassCollection MaterialComboBox

Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetMaterialClassCollection", ftDiagSeverityError
End Sub


Public Sub MaterialGroupUpdate(index As Long, newname As String)

    'Error check
    On Error GoTo ErrHandler
    'find the parameter list in the local database
    ld.MaterialGroupUpdate index, newname

Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " MaterialGroupUpdate", ftDiagSeverityError
End Sub

Public Function GetDisplayLev2ParKeypad() As Boolean
    If lDisplayLev2ParKeypad = 1 Then
        GetDisplayLev2ParKeypad = True
    Else
        GetDisplayLev2ParKeypad = False
    End If
End Function

'**************************************************
'***** THE FOLLOWING WAS ADDED FOR THE TPI BK *****
'**************************************************


Public Sub OpenRecipeSearch()
On Error Resume Next 'Error Check
    Dim frm As New FrmRecipeSearch
    frm.SetLD = ld
    frm.UserCode = "M"
    frm.Init
    RSView32SEIdleDetectControl.Enabled = False
    frm.Show
    If Me.TagParameters(1) = 1 Then
        RSView32SEIdleDetectControl.Enabled = True
    End If
End Sub

Public Sub OpenTrains()
On Error Resume Next 'Error Check
    Dim frm As New FrmTrains
    frm.SetLD = ld
    frm.UserCode = "M"
    frm.Init
    RSView32SEIdleDetectControl.Enabled = False
    frm.Show
    If Me.TagParameters(1) = 1 Then
        RSView32SEIdleDetectControl.Enabled = True
    End If
End Sub

Public Function OpenDownload(PLCTag As String, Controller As String, BatchTank As String)
On Error Resume Next 'Error Check
    Dim frm As New FrmDownload
    frm.SetLD = ld
    frm.PLCTag = PLCTag
    frm.Controller = Controller
    frm.BatchTank = BatchTank   'Kris Modification
    frm.Init
    RSView32SEIdleDetectControl.Enabled = False
    frm.Show
    If Me.TagParameters(1) = 1 Then
        RSView32SEIdleDetectControl.Enabled = True
    End If
End Function

'Public Function OpenDownload(PLCTag As String, Controller As String)   'Kris Modification Backup
'On Error Resume Next 'Error Check
'    Dim frm As New FrmDownload
'    frm.SetLD = ld
'    frm.PLCTag = PLCTag
'    frm.Controller = Controller
'    frm.Init
'    RSView32SEIdleDetectControl.Enabled = False
'    frm.Show
'    If Me.TagParameters(1) = 1 Then
'        RSView32SEIdleDetectControl.Enabled = True
'    End If
'End Function

Public Function GetCMdescription(Tagname As String) As String
On Error Resume Next 'Error Check

    GetCMdescription = ld.GetCMdescription(Tagname)

End Function

'**************************************************
'***** THE FOLLOWING WAS ADDED FOR Property Copy *****
'**************************************************


Public Property Get ParameterTypetoCopy() As String
    ParameterTypetoCopy = lParameterTypetoCopy
End Property

Public Property Let ParameterTypetoCopy(ByVal value As String)
    lParameterTypetoCopy = value
End Property

Public Property Get ParameterUnittoCopy() As String
    ParameterUnittoCopy = lParameterUnittoCopy
End Property

Public Property Let ParameterUnittoCopy(ByVal value As String)
    lParameterUnittoCopy = value
End Property

Public Property Get ParameterTagtoCopy() As String
    ParameterTagtoCopy = lParameterTagtoCopy
End Property

Public Property Let ParameterTagtoCopy(ByVal value As String)
    lParameterTagtoCopy = value
End Property

Public Property Get ParameterGrouptoCopy() As String
    ParameterGrouptoCopy = lParameterGrouptoCopy
End Property

Public Property Let ParameterGrouptoCopy(ByVal value As String)
    lParameterGrouptoCopy = value
End Property


Private Function Round_Up(ByVal d As Double) As Integer
    Dim result As Integer
    result = Math.Round(d)
    If result >= d Then
        Round_Up = result
    Else
        Round_Up = result + 1
    End If
End Function

Public Sub GetRecipeParameternames(rid As String, version As String, step As String, ret As Collection)

    'Error check
    On Error GoTo ErrHandler
    'find the parameter list in the local database
    ld.GetRecipeParameternames rid, version, step, ret

Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetRecipeParameternames", ftDiagSeverityError
End Sub

Public Sub EditMaterial(id As Integer, Materialname As String, ClassID As Integer, MaterialType As Integer, Alias As String, Action As Integer)
    
    'Error check
    On Error GoTo ErrHandler
    
    'load the class list from the local database
    ld.EditMaterial id, Materialname, ClassID, MaterialType, Alias, lUsername, Action
    
Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " EditMaterial", ftDiagSeverityError
End Sub


Public Sub EditMaterialClass(id As Integer, Classname As String, Action As Integer)
    
    'Error check
    On Error GoTo ErrHandler
    
    'load the class list from the local database
    ld.EditMatClass id, Classname, lUsername, Action
    
Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " EditMaterialClass", ftDiagSeverityError
End Sub


Public Function GetCIPRecipePath() As String

    'Error check
    On Error GoTo ErrHandler
    'return plc path of cip recipes
    GetCIPRecipePath = lcippath

Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " GetCIPRecipePath", ftDiagSeverityError
End Function

Public Function SiteMaterialAlias_FromTPMDB(MaterialID As Long) As String 'Added AKB 8/20/21
    
    'Error check
    On Error GoTo ErrHandler
    'load the material name from the local database
    SiteMaterialAlias_FromTPMDB = ld.GetSiteMaterialAlias(MaterialID)
        
Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " SiteMaterialAlias_FromTPMDB", ftDiagSeverityError
End Function

Public Function MaterialEU_FromTPMDB(MaterialID As Long) As String 'ND_20230109
    
    'Error check
    On Error GoTo ErrHandler
    'load the material name from the local database
    MaterialEU_FromTPMDB = ld.GetMaterialEU(MaterialID)
    
Exit Function

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " MaterialEU_FromTPMDB", ftDiagSeverityError
End Function
