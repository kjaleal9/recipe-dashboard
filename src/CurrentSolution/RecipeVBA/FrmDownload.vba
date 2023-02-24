

Option Explicit
'Option Compare Text

Private rs As adodb.Recordset
Private rs1 As adodb.Recordset
Private db As dbConnect
Private valuetofind As Long
Private sqlstring As String
Private sqltrain As String
Private basesize As Long
'Private conn As OPCclient           '!!!!! Comment out in Factory Talk !!!!!
Private Tag_Group As TagGroup      '!!!!! Comment out in Excel !!!!!
Private TagsInError As StringList  '!!!!! Comment out in Excel !!!!!
Private PLCTagValue As String
Private ControllerName As String

Private rcpBSAct As Double 'Actual batch size
Private rcpBSNom As Double 'Nominal batch size
Private rcpBSMin As Double 'Minimum batch size
Private rcpBSMax As Double 'Maximum batch size
Private Declare Sub Sleep Lib "Kernel32.dll" (ByVal ms As Long)
Private ld As LocalDatabase
Private bcancel As Boolean

Private BatchTankName As String 'Kris Leal mod

Public Sub Init()
'TPMDB Connection
       If Application.Name = "Microsoft Excel" Then
        Set db = New dbConnect
'        db.ServerName = ActiveWorkbook.Sheets("Setup").Range("b1").value
'        db.DatabaseName = ActiveWorkbook.Sheets("Setup").Range("b2").value
'        db.UserName = ActiveWorkbook.Sheets("Setup").Range("b10").value
''        db.Password = ActiveWorkbook.Sheets("Setup").Range("b11").value
'        db.Connect
    Else
        Set db = ld.db
        db.ConnectDSN
    End If
    
    CmbTrain.Style = fmStyleDropDownList
    CmbTrain.BoundColumn = 0
    CmbTrain.ListWidth = 140
    CmbTrain.ColumnCount = 1
    CmbTrain.ColumnWidths = "1"
    
    CmbMaterial.Style = fmStyleDropDownList
    CmbMaterial.BoundColumn = 0
    CmbMaterial.ListWidth = 140
    CmbMaterial.ColumnCount = 3
    CmbMaterial.ColumnWidths = "0,40,40" 'plc_ID,SiteMaterialAlias,Name
    
    CmbRecipe.Style = fmStyleDropDownList
    CmbRecipe.BoundColumn = 0
    CmbRecipe.ListWidth = 140
    CmbRecipe.ColumnCount = 6
    CmbRecipe.ColumnWidths = "1,0,0,0,0,0" 'rid,version,description,batchsizenominal,batchsizemin,batchsizemax
    'Load material combobox
    sqlstring = "SELECT DISTINCT Material.plc_ID,(Material.SiteMaterialAlias), Material.Name " & _
                "FROM Material, Recipe " & _
                "WHERE Material.SiteMaterialAlias in " & _
                    "(SELECT ProductID from Recipe where Status='Approved') " & _
                        "AND Recipe.IsPackagingRecipeType <> 1 " & _
                "ORDER BY Material.SiteMaterialAlias"
    ListRefreshSQL rs, db, sqlstring, CmbMaterial, "plc_ID,SiteMaterialAlias,Name", 0
    If CmbMaterial.ListCount > 0 Then CmbMaterial.RemoveItem (CmbMaterial.ListCount - 1) 'Remove the default blank line from the combobox
Exit Sub

'Error message
ErrHandler:
    MsgBox _
        Title:="VBA Error", _
        Buttons:=vbCritical, _
        Prompt:=Now() & " VBA error #" & Err.Number & " - " & Err.Description & " - in Private Sub UserForm_Initialize(). Contact Tetra Pak for assistance"
        Err.Clear
End Sub

Private Sub CmdCancel_Click()
    bcancel = True
End Sub

Private Sub UserForm_Activate()
'On Error GoTo ErrHandler 'Error Check
    Me.Caption = Controller & " Recipe Download"
End Sub

Private Sub Download()
'On Error GoTo ErrHandler 'Error Check
    Dim tag As String
    Dim fieldlist() As String
    Dim fieldlistDB() As String
    Dim i, j As Integer
    CmdDownload.Enabled = False
    CmdClose.Enabled = False
    CmdCancel.Enabled = True
    bcancel = False
    fieldlist = Split(FieldStringArray, ",")
    fieldlistDB = Split(FieldStringArrayDB, ",")
        
    'clear current recipe
    LstDownload.Clear
    LstDownload.AddItem ("Reset PLC Data")
    If Application.Name = "Microsoft Excel" Then
'        tag = ActiveWorkbook.Sheets("Setup").Range("b4").value
        'Set conn = New OPCclient
'        conn.Init ActiveWorkbook.Sheets("Setup").Range("b6").value, ActiveWorkbook.Sheets("Setup").Range("b3").value
'        conn.WriteTag ActiveWorkbook.Sheets("Setup").Range("b7").value, 1
'        conn.WriteTag ActiveWorkbook.Sheets("Setup").Range("b8").value, 0
'        conn.pause
    Else

        'Define taggroup
'''''        Set Tag_Group = Application.CreateTagGroup(ThisDisplay.AreaName, 250)
        Set Tag_Group = Application.CreateTagGroup(Application.LoadedDisplays.Item("footer").AreaName, 250) 'Replace previous AB 2017-07-29
        Tag_Group.Add PLCTag & ".Clear_Recipe"
        Tag_Group.Add PLCTag & ".Recipe_Download_Complete"

        Tag_Group.Active = True
        If Not Tag_Group.RefreshFromSource(TagsInError) Then
            LogDiagnosticsMessage "Taggroup refresh failed on display " & Name & " for Database connection tags ", ftDiagSeverityError
        End If
        Tag_Group.Item(PLCTag & ".Clear_Recipe").value = 1
        Tag_Group.Item(PLCTag & ".Recipe_Download_Complete").value = 0
    End If

    Sleep 3000 '!!!!! Do not delete. This allows the PLC to reset the previous recipe before downloading the Header info.
    LstDownload.AddItem ("PLC Data Reset")
    
    'download header data
    LstDownload.AddItem ("Download Header")
    sqlstring = "select top 1 * from recipe where rid='" & CmbRecipe.Column(0, CmbRecipe.ListIndex) & "' and version='" & CmbRecipe.Column(1, CmbRecipe.ListIndex) & "'"
    Set rs = db.getRecords(sqlstring)
    
    If Not rs.EOF Then
        If Application.Name = "Microsoft Excel" Then
'            conn.WriteTag tag & "._TPI_BK_Recipe.Header.RID", rs.Fields("RID").value
'            conn.WriteTag tag & "._TPI_BK_Recipe.Header.Version", rs.Fields("Version").value
'            conn.WriteTag tag & "._TPI_BK_Recipe.Header.Description", rs.Fields("Description").value
'            conn.WriteTag tag & "._TPI_BK_Recipe.Header.ProductID", rs.Fields("ProductID").value
'            conn.WriteTag tag & "._TPI_BK_Recipe.Header.MaterialID", GetValue(CmbMaterial, 0, 1)
'            conn.WriteTag tag & "._TPI_BK_Recipe.Header.BatchSize", TxtBS.Text
'            conn.WriteTag tag & "._TPI_BK_Recipe.Header.BatchSizeMin", rs.Fields("BatchSizeMin").value
'            conn.WriteTag tag & "._TPI_BK_Recipe.Header.BatchSizeMax", rs.Fields("BatchSizeMax").value
'            conn.WriteTag tag & "._TPI_BK_Recipe.Header.Train", CmbTrain.Column(0, CmbTrain.ListIndex)
        Else
            'Define taggroup
            Tag_Group.Add PLCTag & "._TPI_BK_Recipe.Header.RID"
            Tag_Group.Add PLCTag & "._TPI_BK_Recipe.Header.Version"
            Tag_Group.Add PLCTag & "._TPI_BK_Recipe.Header.Description"
            Tag_Group.Add PLCTag & "._TPI_BK_Recipe.Header.ProductID"
            Tag_Group.Add PLCTag & "._TPI_BK_Recipe.Header.MaterialID"
            Tag_Group.Add PLCTag & "._TPI_BK_Recipe.Header.BatchSize"
            Tag_Group.Add PLCTag & "._TPI_BK_Recipe.Header.BatchSizeMin"
            Tag_Group.Add PLCTag & "._TPI_BK_Recipe.Header.BatchSizeMax"
            Tag_Group.Add PLCTag & "._TPI_BK_Recipe.Header.Train"
            If Not Tag_Group.RefreshFromSource(TagsInError) Then
                LogDiagnosticsMessage "Taggroup refresh failed on display " & Name & " for Database connection tags ", ftDiagSeverityError
            End If
            
            Tag_Group.Item(PLCTag & "._TPI_BK_Recipe.Header.RID").value = rs.Fields("RID").value
            Tag_Group.Item(PLCTag & "._TPI_BK_Recipe.Header.Version").value = rs.Fields("Version").value
            Tag_Group.Item(PLCTag & "._TPI_BK_Recipe.Header.Description").value = rs.Fields("Description").value
            Tag_Group.Item(PLCTag & "._TPI_BK_Recipe.Header.ProductID").value = rs.Fields("ProductID").value
            Tag_Group.Item(PLCTag & "._TPI_BK_Recipe.Header.MaterialID").value = GetValue(CmbMaterial, 0, 1)
            Tag_Group.Item(PLCTag & "._TPI_BK_Recipe.Header.BatchSize").value = CDbl(TxtBS.Text)
            Tag_Group.Item(PLCTag & "._TPI_BK_Recipe.Header.BatchSizeMin").value = rs.Fields("BatchSizeMin").value
            Tag_Group.Item(PLCTag & "._TPI_BK_Recipe.Header.BatchSizeMax").value = rs.Fields("BatchSizeMax").value
            Tag_Group.Item(PLCTag & "._TPI_BK_Recipe.Header.Train").value = CmbTrain.Column(0, CmbTrain.ListIndex)
        End If
        LstDownload.AddItem ("Header Downloaded")
    Else
        MsgBox "Error Retreiving Data"
        Exit Sub
    End If

    'download step data
    '////////////////////////////////////////////////////////////dont forget to do the scaling logic
    LstDownload.AddItem ("Downloading Recipe Steps")
    Sleep 1000

    j = 0
'    sqlstring = "select * from v_recipestepdownload where rid='" & CmbRecipe.Column(0, CmbRecipe.ListIndex) & "' and version='" & CmbRecipe.Column(1, CmbRecipe.ListIndex) & "' and ((RecipeTrainName= '" & CmbTrain.Column(0, CmbTrain.ListIndex) & "') or (RecipeTrainName is null) )"
'    Set rs = db.executeCommand("z_TPIBK_getRecipestepdownload " & CmbRecipe.Column(0, CmbRecipe.ListIndex) & "'" & CmbRecipe.Column(1, CmbRecipe.ListIndex) & "'" & CmbTrain.Column(0, CmbTrain.ListIndex))   'db.getRecords(sqlstring)
    
    Dim par1 As adodb.Parameter
    Dim par2 As adodb.Parameter
    Dim par3 As adodb.Parameter
    Dim parvalue

    Set par1 = New adodb.Parameter
    par1.Type = adVarChar                      'Parameter type
    par1.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par1.Size = 50                             'Max. size for Parameter value
    par1.value = CmbRecipe.Column(0, CmbRecipe.ListIndex)

    Set par2 = New adodb.Parameter
    par2.Type = adVarChar                      'Parameter type
    par2.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par2.Size = 50                             'Max. size for Parameter value
    par2.value = CmbRecipe.Column(1, CmbRecipe.ListIndex)

    Set par3 = New adodb.Parameter
    par3.Type = adVarChar                      'Parameter type
    par3.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par3.Size = 50                             'Max. size for Parameter value
    par3.value = CmbTrain.Column(0, CmbTrain.ListIndex)

    Set rs = db.executeCommand3Par("TPIBK_getRecipestepdownload", par1, par2, par3)
    
    'MsgBox "EXEC TPIBK_getRecipestepdownload, " & Par1 & ", " & Par2 & ", " & Par3
    
    
    If rs Is Nothing Then
        LstDownload.AddItem ("Download Failed")
        LstDownload.ListIndex = LstDownload.ListCount - 1
    Else
        While Not rs.EOF
            DoEvents
            If bcancel Then
                GoTo Cancel
            End If
            LstDownload.AddItem ("Downloading Step " & rs.Fields("Step").value)
            LstDownload.ListIndex = LstDownload.ListCount - 1
            For i = 0 To UBound(fieldlist)
                If Not IsNull(rs.Fields(fieldlistDB(i)).value) And rs.Fields(fieldlistDB(i)).value <> "" And rs.Fields(fieldlistDB(i)).value <> 0 Then
                   If fieldlistDB(i) = "Material_Amount" And rs.Fields("MaterialScaled").value = 1 Then 'Or fieldlistDB(i) Like "*DINT*#" Or fieldlistDB(i) Like "*Real*#") Then

                        '!!!!! Scale materiual amounts or user DINTS and Reals only !!!!!
                        parvalue = rs.Fields(fieldlistDB(i)).value * (rcpBSAct / rcpBSNom)
                    Else
                        parvalue = rs.Fields(fieldlistDB(i)).value
                    End If
                    
                    'Download parameter value
                    If Application.Name = "Microsoft Excel" Then
'                        conn.WriteTag tag & ".Steps[" & j & "]." & fieldlist(i), parvalue
                    Else
                        Tag_Group.Add PLCTag & "._TPI_BK_Recipe.Steps[" & j & "]." & fieldlist(i)
                        Tag_Group.Item(PLCTag & "._TPI_BK_Recipe.Steps[" & j & "]." & fieldlist(i)).value = parvalue
                    End If
                End If
            Next i
            j = j + 1
            rs.MoveNext
        Wend
Cancel:
        If bcancel Then
            LstDownload.AddItem ("Download Canceled!")
        End If
        rs.Close
        Set rs = Nothing
        
        If Application.Name = "Microsoft Excel" Then
'            conn.WriteTag ActiveWorkbook.Sheets("Setup").Range("b8").value, 1
'            Set conn = Nothing
        Else
            If bcancel Then
                Tag_Group.Item(PLCTag & ".Recipe_Download_Complete").value = 0
            Else
                Tag_Group.Item(PLCTag & ".Recipe_Download_Complete").value = 1
            End If
            Tag_Group.Item(PLCTag & ".Recipe_Download_Complete").value = 1
            'Close tag group
            Tag_Group.Active = False
            Tag_Group.RemoveAll
            Set Tag_Group = Nothing
        End If
        LstDownload.AddItem ("Download Complete")
        LstDownload.ListIndex = LstDownload.ListCount - 1
    End If
    CmdClose.Enabled = True
    CmdDownload.Enabled = True
    CmdCancel.Enabled = False
Exit Sub

'Error message
ErrHandler:
    CmdClose.Enabled = True
    CmdDownload.Enabled = True
    CmdCancel.Enabled = False
    MsgBox _
        Title:="VBA Error", _
        Buttons:=vbCritical, _
        Prompt:=Now() & " VBA error #" & Err.Number & " - " & Err.Description & " - in Private Sub Download(). Contact Tetra Pak for assistance"
        Err.Clear
End Sub

Public Property Get PLCTag() As String
'On Error Resume Next 'Error Check
    PLCTag = PLCTagValue
End Property

Public Property Let PLCTag(ByVal vNewValue As String)
'On Error Resume Next 'Error Check
    PLCTagValue = vNewValue
End Property

Public Property Get Controller() As String
'On Error Resume Next 'Error Check
    Controller = ControllerName
End Property

Public Property Let Controller(ByVal vNewValue As String)
'On Error Resume Next 'Error Check
    ControllerName = vNewValue
End Property

Public Property Get BatchTank() As String
'On Error Resume Next 'Error Check
    BatchTank = BatchTankName
End Property

Public Property Let BatchTank(ByVal vNewValue As String)
'On Error Resume Next 'Error Check
    BatchTankName = vNewValue
End Property

Private Sub CmbMaterial_Change()
'On Error GoTo ErrHandler 'Error Check
    If CmbMaterial.ListIndex = -1 Then
        Exit Sub 'No material selected
    ElseIf IsNull(CmbMaterial.Column(0, CmbMaterial.ListIndex)) Or CmbMaterial.Column(0, CmbMaterial.ListIndex) = "" Then
        Exit Sub 'Invalid material selected
    Else
    
    sqlstring = "SELECT rid,version,description,batchsizenominal,batchsizemin,batchsizemax " & _
            "FROM recipe where productid='" & CmbMaterial.Column(1, CmbMaterial.ListIndex) & "' " & _
            "AND status='Approved' and usebatchkernel='True'"
        ListRefreshSQL rs, db, sqlstring, CmbRecipe, "rid,version,description,batchsizenominal,batchsizemin,batchsizemax", 0
        LblDescription.Caption = CmbMaterial.Column(2, CmbMaterial.ListIndex)
        If CmbRecipe.ListCount > 0 Then CmbRecipe.RemoveItem (CmbRecipe.ListCount - 1) 'Remove the default blank line from the combobox
        CmdDownload.Enabled = CmbMaterial.ListIndex >= 0 And CmbRecipe.ListIndex >= 0 And CmbTrain.ListIndex >= 0
    End If
Exit Sub

'Error message
ErrHandler:
    MsgBox _
        Title:="VBA Error", _
        Buttons:=vbCritical, _
        Prompt:=Now() & " VBA error #" & Err.Number & " - " & Err.Description & " - in Private Sub CmbMaterial_Change(). Contact Tetra Pak for assistance"
        Err.Clear
End Sub

Private Sub CmbRecipe_Change()
'On Error GoTo ErrHandler 'Error Check

    If CmbRecipe.ListIndex = -1 Then Exit Sub
    TxtBS.Text = CmbRecipe.Column(3, CmbRecipe.ListIndex)
    rcpBSNom = CmbRecipe.Column(3, CmbRecipe.ListIndex)
    TxtBSMin.Text = CmbRecipe.Column(4, CmbRecipe.ListIndex)
    TxtBSMax.Text = CmbRecipe.Column(5, CmbRecipe.ListIndex)
    If IsNumeric(CmbRecipe.Column(3, CmbRecipe.ListIndex)) Then basesize = CmbRecipe.Column(3, CmbRecipe.ListIndex) Else basesize = 0
    LblRecipe = CmbRecipe.Column(2, CmbRecipe.ListIndex)
    
    'load trains
    sqltrain = "select distinct(Name) from RecipeTrain order by Name ASC"
    Set rs = db.getRecords(sqltrain)
    CmbTrain.Clear
    
    LogDiagnosticsMessage BatchTankName

'    While Not rs.EOF
    
        sqlstring = "EXEC [dbo].[TPIBK_getValidTrains] @RID = '" & CmbRecipe.Column(0, CmbRecipe.ListIndex) & "', @Ver = " & CmbRecipe.Column(1, CmbRecipe.ListIndex) & ", @BatchTank = '" & BatchTank & "'"
        LogDiagnosticsMessage sqlstring
        
'        sqlstring = "SELECT ProcessClass_Name " + _
'                    "FROM RecipeEquipmentRequirement, ProcessClass " + _
'                    "WHERE RecipeEquipmentRequirement.Recipe_RID= '" & CmbRecipe.Column(0, CmbRecipe.ListIndex) & "' " + _
'                        "AND RecipeEquipmentRequirement.Recipe_Version= '" & CmbRecipe.Column(1, CmbRecipe.ListIndex) & "' " + _
'                        "AND ProcessClass.Name=RecipeEquipmentRequirement.ProcessClass_Name " + _
'                        "AND ProcessClass.ID in " + _
'                            "(SELECT Equipment.ProcessClass_ID " + _
'                            "FROM Equipment, RecipeTrainEquipment, RecipeTrain " + _
'                            "WHERE Equipment.id = RecipeTrainEquipment.Equipment_ID " + _
'                                "AND RecipeTrain.Name= '" & rs.Fields("name").value & "' " + _
'                                "AND RecipeTrain.ID=RecipeTrainEquipment.RecipeTrain_ID)" 'Removed Not AKB 2021-01-26

         Set rs1 = db.getRecords(sqlstring)
         While Not rs1.EOF 'Then 'Added Not AKB 2021-01-26
'MsgBox sqlstring
            CmbTrain.AddItem (rs1.Fields("name").value)
         'End If
            rs1.MoveNext
         Wend
         
         rs1.Close
         'rs.MoveNext
'    Wend
    rs.Close
    Set rs = Nothing
    CmbTrain.Enabled = True
    If CmbTrain.ListCount = 0 Then 'Added AB 2017-08-14
        CmbTrain.Enabled = False
        MsgBox "No Valid Train Available For Selected Recipe"
    ElseIf CmbTrain.ListCount = 1 Then
        CmbTrain.ListIndex = 0
        CmbTrain.Enabled = False
    End If
    CmdDownload.Enabled = CmbMaterial.ListIndex >= 0 And CmbRecipe.ListIndex >= 0 And CmbTrain.ListIndex >= 0
Exit Sub

'Error message
ErrHandler:
    MsgBox _
        Title:="VBA Error", _
        Buttons:=vbCritical, _
        Prompt:=Now() & " VBA error #" & Err.Number & " - " & Err.Description & " - in Private Sub CmbRecipe_Change(). Contact Tetra Pak for assistance"
        Err.Clear
End Sub

Private Sub CmbTrain_Change()
'On Error Resume Next 'Error Check
    CmdDownload.Enabled = CmbMaterial.ListIndex >= 0 And CmbRecipe.ListIndex >= 0 And CmbTrain.ListIndex >= 0
End Sub

Private Sub CmdClose_Click()
'On Error Resume Next 'Error Check
    Unload Me
End Sub

Private Sub CmdDownload_Click()
'On Error Resume Next 'Error Check
    If rcpBSAct >= TxtBSMin And rcpBSAct <= TxtBSMax Then
        'perform download
        Download
    ElseIf rcpBSAct < CLng(TxtBSMin.Text) Then
        MsgBox "The requested batch size is less than the minimum amount, " & rcpBSMin
    Else
        MsgBox "The requested batch size is greater than the maximum amount, " & rcpBSMax
    End If
End Sub

Private Sub TxtBS_Change()
'On Error Resume Next 'Error Check
    If IsNull(TxtBS) Or TxtBS = "" Then rcpBSAct = 0 Else rcpBSAct = TxtBS
End Sub

Private Sub TxtBS_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'On Error Resume Next 'Error Check
    CheckNumericInput KeyAscii, TxtBS
End Sub

Private Sub TxtBSMax_Change()
'On Error Resume Next 'Error Check
    If IsNull(TxtBSMax) Or TxtBSMax = "" Then rcpBSMax = 0 Else rcpBSMax = TxtBSMax
End Sub

Private Sub TxtBSMin_Change()
'On Error Resume Next 'Error Check
    If IsNull(TxtBSMin) Or TxtBSMin = "" Then rcpBSMin = 0 Else rcpBSMin = TxtBSMin
End Sub

Public Property Let SetLD(ByVal vNewValue As LocalDatabase)
    Set ld = vNewValue
End Property
