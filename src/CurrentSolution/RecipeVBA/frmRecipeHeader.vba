Option Explicit
'Option Compare Text

Private rs As adodb.Recordset
Private rs1 As adodb.Recordset
Private rstype As adodb.Recordset
Private db As dbConnect
Private projectid As Long
Private valuetofind As Long
Private sqlunits As String
Private sqlselected As String
Private BID As String
Private Bver As String
Private NewRecipe As Boolean
Private EditRecipe As Boolean
Private SaveAsRecipe As Boolean
Private sqlproduct As String
Private sqlclass As String
Private sqlavailable As String
Private sqlstring As String
Private UserCodeValue As String

'Variables used to detect a change and enable the cmdSave button
Private mRID As String
Private mDesc As String
Private mMaterial As Integer
Private mProduct As String
Private mBSNom As String
Private mBSMin As String
Private mBSMax As String
Private ld As LocalDatabase

Public Sub Init()
On Error GoTo ErrHandler 'Error Check

       If Application.Name = "Microsoft Excel" Then
        Set db = New dbConnect
    Else
        Set db = ld.db
        db.ConnectDSN
    End If
    
    CmbMaterial.Style = fmStyleDropDownList
    CmbMaterial.BoundColumn = 0
    CmbMaterial.ListWidth = 140
    CmbMaterial.ColumnCount = 2
    CmbMaterial.ColumnWidths = "0,1"
    
    CmbProduct.Style = fmStyleDropDownList
    CmbProduct.BoundColumn = 0
    CmbProduct.ListWidth = 140
    CmbProduct.ColumnCount = 4
    CmbProduct.ColumnWidths = "0,40,40,0" 'plc_ID,SiteMaterialAlias,Name,materialclass_id
    
    LstSelected.BoundColumn = 0
    LstSelected.ListWidth = 140
    LstSelected.ColumnCount = 5
    LstSelected.ColumnWidths = "0,0,1,0,0" 'id,ProcessClass_Name,Equipment_Name,message,ProcessClass_Description
    
    LstUnits.BoundColumn = 0
    LstUnits.ListWidth = 140
'''''    LstUnits.ColumnCount = 2
    LstUnits.ColumnCount = 3
'''''    LstUnits.ColumnWidths = "0,1" 'id,ProcessClass_Name
    LstUnits.ColumnWidths = "0,0,1" 'id,Name,Description
Exit Sub

'Error message
ErrHandler:
   ' MsgBox "VBA error Animation Start " & Err.Number & " " & Err.Description & " on display " & name
End Sub

Public Sub LoadBatch(oFor As String)
'On Error Resume Next 'Error Check
    Dim sqlstring As String
    Dim Mat, MatID, plcid As String
    Dim typecount
    
    Me.Caption = "Recipe (" & oFor & ")"

    If oFor = "View" Or oFor = "Edit" Or oFor = "SaveAs" Then
        sqlstring = "select top 1 * from recipe where rid='" & BatchID & "' and version='" & BatchVersion & "'"

        typecount = 0
        'get the base info for the recipe
        Set rs = db.getRecords(sqlstring)
        
        If Not rs Is Nothing And Not rs.EOF Then
            TxtDesc.Text = rs.Fields("description").value
            If IsNull(rs.Fields("ProductID").value) Or rs.Fields("ProductID").value = "" Then Mat = "%" Else Mat = rs.Fields("ProductID").value
            
            'Populate product combobox (must be before material combobox)
            sqlproduct = "SELECT plc_ID, SiteMaterialAlias,[Name],materialclass_id " & _
                            "FROM Material " & _
                            "WHERE SiteMaterialAlias = '" & Mat & "' " & _
                            "ORDER by SiteMaterialAlias"
            Set rs1 = db.getRecords(sqlproduct)
            If IsNull(rs1.Fields("materialclass_id").value) Or rs1.Fields("materialclass_id").value = "" Then MatID = "" Else MatID = rs1.Fields("materialclass_id").value
            If IsNull(rs1.Fields("plc_ID").value) Or rs1.Fields("plc_ID").value = "" Then plcid = "" Else plcid = rs1.Fields("plc_ID").value

            'Populate material combobox
            sqlclass = "SELECT DISTINCT MaterialClass.ID, MaterialClass.[Name] " & _
                            "FROM MaterialClass " & _
                            "ORDER BY MaterialClass.[Name]"
            ListRefreshSQL rs1, db, sqlclass, CmbMaterial, "id,Name", rs1.Fields("materialclass_id").value 'CmbProduct.Column(3, CmbProduct.ListIndex)
            If CmbMaterial.ListCount > 0 Then CmbMaterial.RemoveItem (CmbMaterial.ListCount - 1)
            mMaterial = CmbMaterial.ListIndex

            'Populate product combobox (must be before material combobox)
            sqlproduct = "SELECT plc_ID, SiteMaterialAlias,[Name],materialclass_id " & _
                            "FROM Material " & _
                            "WHERE materialclass_id = '" & MatID & "' " & _
                            "ORDER by SiteMaterialAlias"
            ListRefreshSQL rs1, db, sqlproduct, CmbProduct, "plc_ID,SiteMaterialAlias,Name,materialclass_id", plcid
            If CmbProduct.ListCount > 0 Then CmbProduct.RemoveItem (CmbProduct.ListCount - 1)
            mProduct = CmbProduct.ListIndex
            
            LblStatus.Caption = rs.Fields("status").value
            If rs.Fields("status").value = "Registered" Then oFor = "Edit" 'Allow editing of Registered recipes to minimize the version number
            LblType.Caption = rs.Fields("recipetype").value
            LblVersion.Caption = BatchVersion
            LblVerDate.Caption = rs.Fields("versiondate").value

            TxtBSNom.Text = rs.Fields("batchsizenominal").value
            TxtBSMin.Text = rs.Fields("batchsizemin").value
            TxtBSMax.Text = rs.Fields("batchsizemax").value
            
            If oFor <> "View" Then
                TxtDesc.Enabled = True
                CmbMaterial.Enabled = True
                CmbProduct.Enabled = True
                TxtBSNom.Enabled = True
                TxtBSMin.Enabled = True
                TxtBSMax.Enabled = True
                CmdProcedure.Enabled = True
'                Frame1.Enabled = True
            End If
            
            Select Case oFor
            Case "Edit"
                EditRecipe = True
                CmdEditRecipe.Enabled = False
                
                TxtName.Text = BID
               
                Select Case rs.Fields("status").value
                Case "Registered"
                    CmdValidate.Enabled = True
                    CmdApprove.Enabled = False
                Case "Valid"
                    CmdValidate.Enabled = False
                    CmdApprove.Enabled = True
                    CmdEditRecipe.Enabled = True 'Added AB 2017-08-11
                Case "Approved"
                    CmdValidate.Enabled = False
                    CmdApprove.Enabled = False
                End Select
            
            Case "SaveAs"
                TxtName.Text = ""
                TxtName.Enabled = True
                BatchVersion = 1
                LblStatus.Caption = "Registered"
                LblType.Caption = rs.Fields("recipetype").value
                LblVerDate.Caption = Format(Now, "yyyy-mm-dd hh:mm:ss")
                SaveAsRecipe = True

            Case "View"
                TxtName.Text = BID
                If LblStatus.Caption = "Registered" Then CmdValidate.Enabled = True
                If LblStatus.Caption = "Valid" Then CmdApprove.Enabled = True
                If LblStatus.Caption = "Approved" Then CmdApprove.Enabled = False
            End Select
        End If
        rs.Close
        Set rs = Nothing
            
        RefreshLists
        RefreshMainUnit
    
    ElseIf oFor = "New" Then
        NewRecipe = True
        TxtName.Text = ""
        TxtName.Enabled = True
        TxtDesc.Enabled = True
        CmbMaterial.Enabled = True
        CmbProduct.Enabled = True
        TxtBSNom.Enabled = True
        TxtBSMin.Enabled = True
        TxtBSMax.Enabled = True
        CmdProcedure.Enabled = True
                        
        LblType.Caption = ""
        LblVersion.Caption = 1
        TxtDesc.Text = ""
        TxtBSMax = 0
        TxtBSMin = 0
        TxtBSNom = 0
        LblStatus.Caption = ""
        LblVerDate.Caption = ""
        LblMainBatch.Caption = ""
'        LstSelected.Clear

        sqlproduct = "select plc_ID, SiteMaterialAlias,[Name],materialclass_id " & _
                        "from Material " & _
                        "where SiteMaterialAlias Is Not Null and materialclass_id is not null " & _
                        "order by SiteMaterialAlias"
        ListRefreshSQL rs1, db, sqlproduct, CmbProduct, "plc_ID,SiteMaterialAlias,Name,materialclass_id", 0

        sqlclass = "SELECT DISTINCT MaterialClass.ID, MaterialClass.[Name] " & _
                    "FROM MaterialClass " & _
                        "left join Material " & _
                        "on MaterialClass.ID = Material.MaterialClass_ID "
        ListRefreshSQL rs1, db, sqlclass, CmbMaterial, "id,Name", 0
    End If

    mRID = TxtName
    mDesc = TxtDesc
    mBSNom = TxtBSNom.Text
    mBSMin = TxtBSMin.Text
    mBSMax = TxtBSMax.Text
    Call EnableSave
End Sub

Private Sub RefreshLists()
On Error Resume Next 'Error Check
    
    'Define local variables
    Dim i As Integer
    
    'Clear lists
    LstUnits.Clear
    LstSelected.Clear

    'Add selected Process Classes (Units) to LstSelected
    sqlselected = "SELECT RER.ID, ProcessClass_Name, 
                    CASE coalesce(Equipment_Name,PC.Description) 
                        WHEN '' THEN PC.Description 
                        ELSE Equipment_Name 
                    END as Equipment_Name, 
                    
                    CASE 
                        WHEN ROW_NUMBER() over(partition by ProcessClass_Name order by ProcessClass_Name)<2 THEN ProcessClass_Name 
                        ProcessClass_Name+' #'+ltrim(ROW_NUMBER() over(partition by ProcessClass_Name order by ProcessClass_Name)) 
                    END as message, PC.Description As Description " & _
                    
                    "FROM RecipeEquipmentRequirement RER " & _
                    "JOIN ProcessClass PC ON RER.ProcessClass_Name = PC.Name " & _
                    "WHERE Recipe_RID= '" & BatchID & "' AND Recipe_Version=" & BatchVersion & _
                    "ORDER BY Description"

    ListRefreshSQL rs1, db, sqlselected, LstSelected, "id,ProcessClass_Name,Equipment_Name,message,Description", 0

    'Add available Process Classes (Units) to LstUnits
    sqlunits = "EXEC [dbo].[TPIBK_getAvailUnits] @RID = '" & BatchID & "', @Ver = " & BatchVersion
    ListRefreshSQL rs1, db, sqlunits, LstUnits, "id,Name,Description", 0
    
    If LstSelected.ListCount > 0 Then LstSelected.RemoveItem (LstSelected.ListCount - 1)
    If LstUnits.ListCount > 0 Then LstUnits.RemoveItem (LstUnits.ListCount - 1)
    CmdDelete.Enabled = EditRecipe And LstSelected.ListCount > 0
    CmdAdd.Enabled = EditRecipe And LstUnits.ListCount > 0
End Sub

Private Sub RefreshMainUnit()
    CmdMainAdd.Enabled = False
    CmdMainDelete.Enabled = False
    sqlselected = "SELECT RER.ID, ProcessClass_Name, Equipment_Name,ismainbatchunit, PC.Description As Description " & _
                                "FROM RecipeEquipmentRequirement RER " & _
                                "JOIN ProcessClass PC ON RER.ProcessClass_Name = PC.Name " & _
                                "WHERE Recipe_RID= '" & BatchID & "' " & _
                                "and ismainbatchunit= 1 " & _
                                "and   Recipe_Version=" & BatchVersion
    Set rs1 = db.getRecords(sqlselected)
    If Not rs1 Is Nothing And Not rs1.EOF Then
'''''        LblMainBatch.Caption = rs1.Fields("ProcessClass_Name").value
        LblMainBatch.Caption = rs1.Fields("Equipment_Name").value
        If NewRecipe Or EditRecipe Or SaveAsRecipe Then
            CmdMainAdd.Enabled = False
            CmdMainDelete.Enabled = True
        End If
    Else
        LblMainBatch.Caption = ""
        If NewRecipe Or EditRecipe Or SaveAsRecipe Then
            LblMainBatch.Caption = ""
            CmdMainAdd.Enabled = True
            CmdMainDelete.Enabled = False
        End If
    End If
    rs1.Close
    Set rs1 = Nothing
End Sub

Private Sub CmbMaterial_Change()
    If (NewRecipe Or EditRecipe Or SaveAsRecipe) And CmbMaterial.ListIndex >= 0 Then
        sqlproduct = "select plc_ID, SiteMaterialAlias,[Name],materialclass_id " & _
                        "from Material " & _
                        "where SiteMaterialAlias Is Not Null and materialclass_id is not null " & _
                        "and materialclass_id=" & CmbMaterial.Column(0, CmbMaterial.ListIndex) & _
                        "order by SiteMaterialAlias"
        ListRefreshSQL rs1, db, sqlproduct, CmbProduct, "plc_ID,SiteMaterialAlias,Name,materialclass_id", 0
        If CmbProduct.ListCount > 0 Then CmbProduct.RemoveItem (CmbProduct.ListCount - 1)
    End If
    Call EnableSave
End Sub

' Form event handlers
Private Sub CmbProduct_Change()
On Error Resume Next 'Error Check
    LblProductDesc.Caption = CmbProduct.Column(2, CmbProduct.ListIndex)
    Call EnableSave
End Sub

Private Sub TxtName_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'On Error Resume Next 'Error Check
    CheckSpecialInput KeyAscii, TxtName
    Call EnableSave
End Sub

Private Sub TxtName_Change()
'On Error Resume Next 'Error Check
    Call EnableSave
End Sub

Private Sub TxtDesc_Change()
'On Error Resume Next 'Error Check
    Call EnableSave
End Sub

Private Sub TxtBSNom_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'On Error Resume Next 'Error Check
    CheckNumericInput KeyAscii, TxtBSNom
End Sub

Private Sub TxtBSNom_Change()
'On Error Resume Next 'Error Check
    Call EnableSave
End Sub

Private Sub TxtBSMin_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'On Error Resume Next 'Error Check
    CheckNumericInput KeyAscii, TxtBSMin
End Sub

Private Sub TxtBSMin_Change()
'On Error Resume Next 'Error Check
    Call EnableSave
End Sub

Private Sub TxtBSMax_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'On Error Resume Next 'Error Check
    CheckNumericInput KeyAscii, TxtBSMax
End Sub

Private Sub TxtBSMax_Change()
'On Error Resume Next 'Error Check
    Call EnableSave
End Sub

Private Sub EnableSave()
    CmdSave.Enabled = TxtName <> "" And Not (TxtName = mRID And TxtDesc = mDesc And CmbMaterial.ListIndex = mMaterial And _
        CmbProduct.ListIndex = mProduct And TxtBSNom.Text = mBSNom And TxtBSMin.Text = mBSMin And TxtBSMax.Text = mBSMax)
End Sub
'
Private Sub CmdValidate_Click()
'On Error Resume Next 'Error Check
    
    'Declare local variables
    Dim cmd As adodb.Command
    Dim par1 As adodb.Parameter
    Dim par2 As adodb.Parameter
    Dim par3 As adodb.Parameter
    
    Set cmd = New adodb.Command
    Set par1 = cmd.CreateParameter("@RID", adVarChar, adParamInput, 30, TxtName.Text)
    Set par2 = cmd.CreateParameter("@Version", adVarChar, adParamInput, 10, LblVersion.Caption)
    Set par3 = cmd.CreateParameter("@ReturnMsg", adVarChar, adParamOutput, 100)
    
    If CDec(TxtBSMin.Text) > CDec(TxtBSMax.Text) Then
        MsgBox "The minimum volume must be less than or equal to the maximum volume"
        GoTo Invalid
    ElseIf CDec(TxtBSNom.Text) < CDec(TxtBSMin.Text) Then
        MsgBox "The nominal volume must be greater than or equal to the minimum volume"
        GoTo Invalid
    ElseIf CDec(TxtBSNom.Text) > CDec(TxtBSMax.Text) Then
        MsgBox "The nominal volume must be less than or equal to the maximum volume"
        GoTo Invalid
    ElseIf (CDec(TxtBSNom.Text) + CDec(TxtBSMax.Text) + CDec(TxtBSMin.Text)) = 0 Then
        MsgBox "Enter values for Nominal, Min, and Max sizes"
        GoTo Invalid
    End If
    
    If LstSelected.ListCount = 0 Then
        MsgBox "Must select at least one process class"
        GoTo Invalid
    End If
    
    If LblMainBatch = "" Then
        MsgBox "Must select a main batch unit"
        GoTo Invalid
    End If
    
    db.executeCommand3Par "TPIBK_ValidateRecipe", par1, par2, par3
    If par3.value <> "Valid" Then
        MsgBox par3.value
        GoTo Invalid 'Load
    End If
    
    LblStatus.Caption = "Valid"
    LblVerDate.Caption = Format(Now, "yyyy-mm-dd hh:mm:ss")
    'If the recipe is valid update the status
    sqlstring = "UPDATE [dbo].[Recipe] " + _
                "SET [Status] = 'Valid', [VersionDate] = '" & LblVerDate.Caption & "'" + _
                "WHERE RID = '" & BatchID & "' and Version  = '" & BatchVersion & "'"
    db.execute sqlstring
    
''''''    LoadBatch "Valid" 'added AKB
''''''Exit Sub 'Added AB 2017-08-11

Invalid:
    LoadBatch "Edit"
End Sub

Private Sub CmdSave_Click()
'On Error GoTo ErrHandler 'Error Check
    'Define local variables
    Dim rsSaveAs As adodb.Recordset
    Dim Mat As String
    
    If TxtName.Text = "" Then
        MsgBox "Please enter a valid recipe ID"
        Exit Sub
    End If

    If NewRecipe Or SaveAsRecipe Then
        'Check the new recipe name is unique
        sqlstring = "SELECT [RID]FROM RECIPE WHERE [RID] = '" & TxtName.Text & "'"
        Set rsSaveAs = db.getRecords(sqlstring)
        If rsSaveAs.RecordCount > 0 Then
            MsgBox "Must enter a unique recipe ID"
            rsSaveAs.Close
            Set rsSaveAs = Nothing
            Exit Sub
        End If
                    
        LblVersion.Caption = 1
        If IsNull(CmbProduct.Column(1, CmbProduct.ListIndex)) Or (CmbProduct.Column(1, CmbProduct.ListIndex)) = "" Then
            Mat = 1
        Else
            Mat = CmbProduct.Column(1, CmbProduct.ListIndex)
        End If

        sqlstring = "INSERT INTO [Recipe] ([RID],[Version],[RecipeType],[NbrOfExecutions],[VersionDate],[Description],[EffectiveDate],[ExpirationDate],[ProductID],[BatchSizeNominal],[BatchSizeMin],[BatchSizeMax],[Status],[UseBatchKernel],[CurrentElementID],[RecipeData],[RunMode],[IsPackagingRecipeType]) " + _
                    "Values " + _
               "('" & TxtName.Text & "' " + _
               ", '" & LblVersion.Caption & "' " + _
               ",'Master', NULL " + _
               ",'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " + _
               ",'" & TxtDesc.Text & "', NULL, NULL " + _
               ",'" & Mat & "' " + _
               ",'" & TxtBSNom.Text & "' " + _
               ",'" & TxtBSMin.Text & "' " + _
               ",'" & TxtBSMax.Text & "' " + _
               ",'Registered', 1, NULL, NULL, 0, 0)"
    Else
        'Update existing record
        sqlstring = "UPDATE [dbo].[Recipe] " + _
                      "SET [Description] = '" & TxtDesc.Text & "' " + _
                         ",[ProductID] = '" + CmbProduct.Column(1, CmbProduct.ListIndex) + "' " + _
                         ",[BatchSizeNominal] = '" & TxtBSNom.Text & "' " + _
                         ",[BatchSizeMin] = '" & TxtBSMin.Text & "' " + _
                         ",[BatchSizeMax] = '" & TxtBSMax.Text & "' " + _
                    "WHERE RID = '" & BatchID & "' and Version  = '" + LblVersion.Caption + "'"
    End If
    db.execute sqlstring
    
    'Copy the selected units to the new reicpe
    If SaveAsRecipe Then
        sqlstring = "TPIBK_CopyRecipe '" & BatchID & "', '" & BatchVersion & "', '" & TxtName.Text & "', '" & BatchVersion & "'"
        db.execute sqlstring
    End If
    
    EditRecipe = False
    NewRecipe = False
    SaveAsRecipe = False
    
    BatchID = TxtName.Text
    BatchVersion = LblVersion.Caption
    LoadBatch "Edit"
Exit Sub

'Error message
ErrHandler:
    MsgBox "VBA error Save " & Err.Number & " " & Err.Description & " on display " & Name

End Sub

Private Sub CmdAdd_Click()
'    sqlstring = "insert into RecipeEquipmentRequirement( EquipmentType , Recipe_RID , Recipe_Version , Equipment_Name , ProcessClass_Name , LateBinding , IsMainBatchUnit) " & _
'                "values( 'Class', '" & BatchID & "', '" & LblVersion.Caption & "', '', '" & GetValue(LstUnits, 1, "NULL") & "', 0, 0)"
    
    sqlstring = "EXEC [dbo].[TPIBK_insProcessClass] @RID = '" & BatchID & "', @VER = " & LblVersion.Caption & ", @EN = '" & GetValue(LstUnits, 2, "NULL") & "', @PC = '" & GetValue(LstUnits, 1, "NULL") & "'"

    db.execute sqlstring
    RefreshLists
    RefreshMainUnit
End Sub

Private Sub CmdApprove_Click()
'On Error Resume Next 'Error Check
    'Change previous revision to obsolete
    If BatchVersion > 1 Then
        sqlstring = "UPDATE [Recipe] " + _
                        "SET [Status] = 'Obsolete' " + _
                   "WHERE RID = '" & BatchID & "' and Version  < '" & BatchVersion & "'"
        db.execute sqlstring
    End If
    
    'update existing record
    sqlstring = "UPDATE [Recipe] " + _
                    "SET [Status] = 'Approved' " + _
               "WHERE RID = '" & BatchID & "' and Version  = '" & BatchVersion & "'"
    db.execute sqlstring
    LoadBatch "Edit"
End Sub

Private Sub CmdEditRecipe_Click()
'On Error Resume Next 'Error Check
    Dim rsEdit As adodb.Recordset

    'Retrieve the highest version number
    sqlstring = "SELECT Max(CONVERT(INT,[Version])) AS Version FROM [dbo].[Recipe] WHERE [RID] = '" & BatchID & "'"
    Set rsEdit = db.getRecords(sqlstring)

    If LblStatus <> "Valid" Then 'Added AB 2017-08-11
        LblStatus.Caption = "Registered"
        LblVersion.Caption = rsEdit.Fields("Version").value + 1
        LblVerDate.Caption = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
        'Make all previous versions of the current recipe Obsolete
        If LblVersion.Caption > 1 Then
            sqlstring = "UPDATE [dbo].[Recipe] " + _
                        "SET [Status] = 'Obsolete' " + _
                        "WHERE RID = '" & BatchID & "' and Version  < " & LblVersion.Caption
            db.execute sqlstring
        End If
        
        'Create a new revision of the current recipe
        sqlstring = "INSERT INTO [Recipe]([RID],[Version],[RecipeType],[NbrOfExecutions],[VersionDate],[Description],[EffectiveDate],[ExpirationDate],[ProductID],[BatchSizeNominal],[BatchSizeMin],[BatchSizeMax],[Status],[UseBatchKernel],[CurrentElementID],[RecipeData],[RunMode],[IsPackagingRecipeType]) " + _
                    "Values " + _
                        "('" & TxtName.Text & "' " + _
                        ", '" & LblVersion.Caption & "' " + _
                        ",'Master', NULL " + _
                        ",'" & LblVerDate.Caption & "' " + _
                        ",'" & TxtDesc.Text & "', NULL, NULL " + _
                        ",'" & CmbProduct.Column(1, CmbProduct.ListIndex) & "' " + _
                        ",'" & TxtBSNom.Text & "' " + _
                        ",'" & TxtBSMin.Text & "' " + _
                        ",'" & TxtBSMax.Text & "' " + _
                        ",'Registered', 1, NULL, NULL, 0, 0)"
        db.execute sqlstring
        
        'Copy the selected units to the new reicpe revision
        db.execute "TPIBK_CopyRecipe '" & BatchID & "', '" & BatchVersion & "', '" & TxtName.Text & "', '" & LblVersion.Caption & "'"
        BatchVersion = LblVersion.Caption
        
    Else
        LblStatus.Caption = "Registered"
        LblVerDate.Caption = Format(Now, "yyyy-mm-dd hh:mm:ss")
        'Update
        sqlstring = "UPDATE [dbo].[Recipe] " + _
                    "SET [Status] = 'Registered', [VersionDate] = '" & LblVerDate.Caption & "'" + _
                    "WHERE RID = '" & BatchID & "' and Version  = " & LblVersion.Caption
        db.execute sqlstring
    End If
    
    LoadBatch "Edit"
End Sub

Private Sub CmdClose_Click()
On Error Resume Next 'Error Check
    Unload Me
End Sub

Private Sub CmdDelete_Click()
On Error Resume Next 'Error Check
    db.execute "delete from RecipeEquipmentRequirement where id = " & LstSelected.Column(0, LstSelected.ListIndex)
    
    RefreshLists
    RefreshMainUnit
End Sub

Private Sub CmdMainAdd_Click()
On Error Resume Next 'Error Check
    db.execute "update RecipeEquipmentRequirement set IsMainBatchUnit =1 where id =" & LstSelected.Column(0, LstSelected.ListIndex)
    RefreshMainUnit
End Sub

Private Sub CmdMainDelete_Click()
On Error Resume Next 'Error Check
    db.execute "update RecipeEquipmentRequirement set IsMainBatchUnit =0 where IsMainBatchUnit =1 and  Recipe_RID = '" & BatchID & "' and Recipe_Version  = '" & BatchVersion & "'"
    RefreshMainUnit
End Sub

Private Sub CmdProcedure_Click()
On Error Resume Next 'Error Check
    Dim frm As New FrmRecipeProcedure

    frm.BatchID = BatchID
    frm.BatchVersion = BatchVersion
    frm.BatchStatus = LblStatus.Caption
    frm.UserCode = UserCode
    frm.SetLD = ld
    frm.Init
    frm.LoadClass
    frm.LoadProcedure
    frm.Show
End Sub

Public Property Get BatchID() As String
    BatchID = BID
End Property

Public Property Let BatchID(ByVal vNewValue As String)
    BID = vNewValue
End Property

Public Property Get BatchVersion() As String
    BatchVersion = Bver
End Property

Public Property Let BatchVersion(ByVal vNewValue As String)
    Bver = vNewValue
End Property


Public Property Get UserCode() As String
'On Error Resume Next 'Error Check
    UserCode = UserCodeValue
End Property

Public Property Let UserCode(ByVal vNewValue As String)
'On Error Resume Next 'Error Check
    UserCodeValue = vNewValue
End Property

Public Property Let SetLD(ByVal vNewValue As LocalDatabase)
    Set ld = vNewValue
End Property

