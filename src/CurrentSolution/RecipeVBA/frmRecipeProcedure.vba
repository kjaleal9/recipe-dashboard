

Option Explicit
'Option Compare Text

Private rs As adodb.Recordset
Private rs1 As adodb.Recordset
Private rstype As adodb.Recordset
Private db As dbConnect
Private valuetofind As Long
Private BID As String
Private Bver As String
Private Bstat As String
Private NewRecipe As Boolean
Private sqlstring As String
Private newstep As Boolean
Private highstep As Integer
Private UserCodeValue As String
Private mCmbType As String
Private mCmbPClass As String
Private mCmbPtype As String
Private mCmbMaterial As String
Private mTxtStep As String
Private mTxtString As String
Private mCmbTransition As String
Private mCmbStep As String
Private ld As LocalDatabase
Private selrecipeid As Long

Public Sub Init()
'On Error Resume Next

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
    
    newstep = False
    selrecipeid = 0
    
'''''    CmbType.Enabled = True 'Deleted AB 2017-08-11
    CmbType.Style = fmStyleDropDownList
    CmbType.BoundColumn = 0
    CmbType.ListWidth = 140
    CmbType.ColumnCount = 2
    CmbType.ColumnWidths = "0,1"
    
    CmbPClass.Style = fmStyleDropDownList
    CmbPClass.BoundColumn = 0
    CmbPClass.ListWidth = 140
    CmbPClass.ColumnCount = 4
    CmbPClass.ColumnWidths = "0,0,0,1" 'RER.id,message,pclass_Id,description
    
    CmbPtype.Style = fmStyleDropDownList
    CmbPtype.BoundColumn = 0
    CmbPtype.ListWidth = 140
    CmbPtype.ColumnCount = 4
    CmbPtype.ColumnWidths = "0,1,0,0" 'id,name,phase_type_id,phasecatagory_id , only show bk phases , use phase catagory to filter material enable
    
    CmbMaterial.Style = fmStyleDropDownList
    CmbMaterial.BoundColumn = 0
    CmbMaterial.ListWidth = 140
    CmbMaterial.ColumnCount = 2
    CmbMaterial.ColumnWidths = "0,1"
    
    CmbTransition.Style = fmStyleDropDownList
    CmbTransition.BoundColumn = 0
    CmbTransition.ListWidth = 140
    CmbTransition.ColumnCount = 2
    CmbTransition.ColumnWidths = "0,1"
    
'''''    TxtStep.Enabled = True 'Deleted AB 2017-08-11
    
    CmbStep.Style = fmStyleDropDownList
    CmbStep.BoundColumn = 0
    CmbStep.ListWidth = 140
    CmbStep.ColumnCount = 2
    CmbStep.ColumnWidths = "0,1"
    
    LstRecipes.BoundColumn = 0
    LstRecipes.ListWidth = LstRecipes.Width - 5
    LstRecipes.ColumnCount = 13
    LstRecipes.ColumnWidths = "0,30,50,0,0,0,0,0,0,0,0,0,0" 'id,step,message,TPIBK_Steptype_ID,processclassphase_id,step1,userstring,recipeequipmenttransition_data_id,nextstep,allocation_type_id,latebinding,material_id,ProcessClass_ID
    
    Lstparameters.BoundColumn = 0
    Lstparameters.ListWidth = Lstparameters.Width - 5
    Lstparameters.ColumnCount = 16
    Lstparameters.ColumnWidths = "0,0,150,0,0,0,0,0,0,0,0,0,50,0,0,30" 'ID,Name,Description,TPIBK_RecipeParameters_ID,ProcessClassPhase_ID,ValueType,Scaled,MinValue,MaxValue,DefValue,IsMaterial,TPIBK_RecipeParameterData_ID,Value,TPIBK_RecipeStepData_ID,defEU,EU
    
    CmdDelete.Enabled = False
    FrAllocation.Enabled = False
    
    Label27.Enabled = False
    Label28.Enabled = False
    Label29.Enabled = False
    Label30.Enabled = False
    Label31.Enabled = False
    Label1.Enabled = False
        
'    ChkLevel.Enabled = FrAllocation.Enabled
    ChkMatcheck.Enabled = FrAllocation.Enabled
    ChkQA.Enabled = FrAllocation.Enabled
    
    sqlstring = "select id,name from TPIBK_StepType order by id"
    ListRefreshSQL rs, db, sqlstring, CmbType, "id,name", 0
    If CmbType.ListCount > 0 Then CmbType.RemoveItem (CmbType.ListCount - 1)
        
    'sqlstring = "select id,ltrim(sitematerialalias) + ' ' + name as name from material where not sitematerialalias is null order by id"
    sqlstring = "select id,ltrim(sitematerialalias) + ' ' + name as name from material where not sitematerialalias is null order by name"
    ListRefreshSQL rs, db, sqlstring, CmbMaterial, "id,name", 0
    If CmbMaterial.ListCount > 0 Then CmbMaterial.RemoveItem (CmbMaterial.ListCount - 1)
    '''''''
Exit Sub
        
'Error message
ErrHandler:
    MsgBox Now() & " VBA error Insert " & Err.Number & " " & Err.Description & " on display " & Name & " for Sub Init()"
    '''
End Sub

Private Sub UserForm_Activate()
'On Error Resume Next

    CmbType.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08Test-11
    TxtStep.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11

End Sub

Private Sub CommandButton2_Click()
   Dim rs As adodb.Recordset
    
    Set rs = db.getRecords("insert into users(plc_id,name,active) output inserted.id values(199,'test',1)")
    If Not rs.EOF Then
        MsgBox rs(0).value
    End If
    rs.Close
    Set rs = Nothing
   
End Sub

Private Sub EnableCmbMaterial() 'Added AB 2017-8-11
'On Error Resume Next
    
    If CmbPtype.ListIndex < 0 Or CmbPtype.Column(0, CmbPtype.ListIndex) = "" Then
        Exit Sub
    ElseIf CmbType.Column(0, CmbType.ListIndex) = 8 Then
        CmbMaterial.Enabled = ChkMatcheck.value And (BatchStatus = "Registered")
    Else
        Select Case CmbPtype.Column(3, CmbPtype.ListIndex)
            Case 1
                CmbMaterial.Enabled = CmbPtype.Enabled And (BatchStatus = "Registered")
            Case 5
                CmbMaterial.Enabled = CmbPtype.Enabled And (BatchStatus = "Registered")
            Case Else
                CmbMaterial.Enabled = False
        End Select
    End If
    Label29.Enabled = CmbMaterial.Enabled
End Sub

Private Sub EnableSave()
'On Error Resume Next 'Error Check

    CmdSave.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
Exit Sub

    If IsNull(CmbPClass.Column(0, CmbPClass.ListIndex)) Or IsNull(CmbPtype.Column(0, CmbPtype.ListIndex)) Or IsNull(CmbMaterial.Column(0, CmbMaterial.ListIndex)) Or _
        IsNull(CmbTransition.Column(0, CmbTransition.ListIndex)) Or IsNull(CmbStep.Column(0, CmbStep.ListIndex)) Then
        CmdSave.Enabled = False
        Exit Sub
    End If


'''''    CmdSave.Enabled = (ThisDisplay.lUserCodes Like "*" & UserCode & "*" Or ThisDisplay.lUserCodes = "*") And
    CmdSave.Enabled = (Application.LoadedDisplays.Item("footer").lUserCodes Like "*" & UserCode & "*" Or Application.LoadedDisplays.Item("footer").lUserCodes = "*") And _
        (mCmbPClass <> CmbPClass.Column(0, CmbPClass.ListIndex) Or _
        mCmbPtype <> CmbPtype.Column(0, CmbPtype.ListIndex) Or _
        mCmbMaterial <> CmbMaterial.Column(0, CmbMaterial.ListIndex) Or _
        mTxtString <> TxtString Or _
        mCmbTransition <> CmbTransition.Column(0, CmbTransition.ListIndex) Or _
        mCmbStep <> CmbStep.Column(0, CmbStep.ListIndex) Or _
        newstep)
End Sub

Private Sub LoadParameters()
    Select Case CmbType.Column(0, CmbType.ListIndex)
        Case 1 'run phase
        Case 2 'start phase
        Case 3 'end phase
        Case 4 'check phase
        Case 5 'download parameters
        Case 6 'transition
            GoTo noparameters
        Case 7 'operator confirmation
            GoTo noparameters
        Case 8 'allocate unit
            GoTo noparameters
        Case 9 'deallocate unit
            GoTo noparameters
        Case 10 'jump
            GoTo noparameters
    End Select
        sqlstring = "SELECT ID, Name, Description, TPIBK_RecipeParameters_ID, ProcessClassPhase_ID, ValueType, Scaled, MinValue, MaxValue, DefValue, " + _
                                          "IsMaterial, MAX(TPIBK_RecipeParameterData_ID) AS TPIBK_RecipeParameterData_ID, SUM(Value) AS Value, MAX(TPIBK_RecipeStepData_ID) " + _
                                          "AS TPIBK_RecipeStepData_ID, defEU, Max(EU) As EU " + _
                    "FROM v_TPIBK_RecipeParameters " + _
                    "WHERE     (TPIBK_RecipeBatchData_ID IN (0, " & LstRecipes.Column(0, LstRecipes.ListIndex) & ")) " + _
                    "GROUP BY ID, Name, Description, TPIBK_RecipeParameters_ID, ProcessClassPhase_ID, ValueType, MinValue, MaxValue, DefValue, Scaled, " + _
                                          "IsMaterial, DefEU " + _
                    "HAVING (ProcessClassPhase_ID = " & CmbPtype.Column(0, CmbPtype.ListIndex) & ") " & _
                    "ORDER BY Description"
       ' MsgBox sqlstring
        ListRefreshSQL rs, db, sqlstring, Lstparameters, "ID,Name,Description,TPIBK_RecipeParameters_ID,ProcessClassPhase_ID,ValueType,Scaled,MinValue,MaxValue,DefValue,IsMaterial,TPIBK_RecipeParameterData_ID,Value,TPIBK_RecipeStepData_ID,defEU,EU", -15
        If Lstparameters.ListCount > 0 Then Lstparameters.RemoveItem (Lstparameters.ListCount - 1)
Exit Sub
noparameters:
        Lstparameters.Clear
End Sub

Public Sub LoadProcedure()
    'fill the main list of steps
    Dim sqlstringstep As String

    sqlstring = "select id,step,message,TPIBK_Steptype_ID,processclassphase_id,step as step1,userstring,recipeequipmenttransition_data_id,nextstep,allocation_type_id,latebinding,material_id,ProcessClass_ID from v_TPIBK_REcipeBatchData where Recipe_RID = '" & BatchID & "' and Recipe_Version  = '" & BatchVersion & "' order by step"
    sqlstringstep = sqlstring
    ListRefreshSQL rs, db, sqlstring, LstRecipes, "id,step,message,TPIBK_Steptype_ID,processclassphase_id,step1,userstring,recipeequipmenttransition_data_id,nextstep,allocation_type_id,latebinding,material_id,ProcessClass_ID", selrecipeid

    If LstRecipes.ListCount > 1 Then
        If IsNull(LstRecipes.Column(1, LstRecipes.ListCount - 2)) Or LstRecipes.Column(1, LstRecipes.ListCount - 2) = "" Then
            highstep = 0
        Else
            highstep = CInt(LstRecipes.Column(1, LstRecipes.ListCount - 2))
        End If
        LstRecipes.RemoveItem (LstRecipes.ListCount - 1)
    End If
    ListRefreshSQL rs, db, sqlstringstep, CmbStep, "id,step", 0
    If CmbStep.ListCount > 0 Then CmbStep.RemoveItem (CmbStep.ListCount - 1)
    If BatchStatus = "Approved" Then
        CmdNew.Enabled = False
        CmdDelete.Enabled = False
        CommandButton1.Enabled = False
    End If
End Sub

Private Sub CmbPClass_Change()
'On Error Resume Next 'Error Check
    If CmbPClass.ListIndex < 0 Then Exit Sub
    If CmbPClass.Column(0, CmbPClass.ListIndex) = "" Then Exit Sub
    
    sqlstring = "SELECT ProcessClassPhase.ID, ProcessClassPhase.Name, ProcessClassPhase.PhaseType_ID, PhaseType.PhaseCategory_ID FROM PhaseType INNER JOIN ProcessClassPhase ON PhaseType.ID = ProcessClassPhase.PhaseType_ID where typebatchkernel=1 and processclass_id=" & CmbPClass.Column(2, CmbPClass.ListIndex) & " order by name"
    ListRefreshSQL rs, db, sqlstring, CmbPtype, "id,name,phasetype_id,PhaseCategory_ID", 0
    If CmbPtype.ListCount > 0 Then CmbPtype.RemoveItem (CmbPtype.ListCount - 1)
    EnableSave
    LoadTransitions
End Sub

Private Sub CmbPtype_Change()
'On Error GoTo errorout 'Error Check

    Call EnableCmbMaterial 'Added AB 2017-08-11
    EnableSave 'Enable save button
Exit Sub
        
'Error message
errorout:
End Sub

Private Sub CmbType_Change()
'On Error Resume Next 'Error Check
    
    'enable the different details options
    FrAllocation.Enabled = False
    TxtString.Enabled = False
    CmbStep.Enabled = False
    CmbTransition.Enabled = False
    Lstparameters.Enabled = False
    TxtString.Text = ""

    Select Case CmbType.Column(0, CmbType.ListIndex)
        Case 1 'run phase
            CmbPClass.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
            CmbPtype.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
            Lstparameters.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
        Case 2 'start phase
            CmbPClass.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
            CmbPtype.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
            Lstparameters.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11ue
        Case 3 'end phase
            CmbPClass.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
            CmbPtype.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
            CmbMaterial.Enabled = False
        Case 4 'check phase
            CmbPClass.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
            CmbPtype.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
            CmbMaterial.Enabled = False
        Case 5 'download parameters
            CmbPClass.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
            CmbPtype.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
            Lstparameters.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
        Case 6 'transition
            CmbPClass.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
            CmbPtype.Enabled = False
            CmbMaterial.Enabled = False
            CmbStep.Enabled = False
            CmbTransition.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
        Case 7 'operator confirmation
            CmbPClass.Enabled = False
            CmbPtype.Enabled = False
            CmbMaterial.Enabled = False
            TxtString.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
        Case 8 'allocate unit
            CmbPClass.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
            CmbPtype.Enabled = False
            If CmbType.Column(0, CmbType.ListIndex) = "8" And ChkMatcheck.value = True Then
                CmbMaterial.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
            Else
                CmbMaterial.Enabled = False
            End If
            
            FrAllocation.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
        Case 9 'deallocate unit
            CmbPClass.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
            CmbPtype.Enabled = False
            CmbMaterial.Enabled = False
        Case 10 'jump
            CmbPClass.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
            CmbPtype.Enabled = False
            CmbMaterial.Enabled = False
            CmbStep.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
            CmbTransition.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
    End Select
    Label27.Enabled = CmbPClass.Enabled
    Label28.Enabled = CmbPtype.Enabled
    Label29.Enabled = CmbMaterial.Enabled
    Label30.Enabled = CmbTransition.Enabled
    Label31.Enabled = CmbStep.Enabled
    Label1.Enabled = TxtString.Enabled
        
'    ChkLevel.Enabled = FrAllocation.Enabled
    ChkMatcheck.Enabled = FrAllocation.Enabled
    ChkQA.Enabled = FrAllocation.Enabled
    
    CmdTransitions.Enabled = CmbTransition.Enabled
    
    Label25.Enabled = Lstparameters.Enabled
    
    Call EnableCmbMaterial 'Added AB 2017-8-11
    
    EnableSave 'Enable save button
End Sub

Private Sub ChkMatcheck_Click()
    Call EnableCmbMaterial
End Sub

Private Sub CmdClose_Click()
'On Error Resume Next 'Error Check
    Unload Me
End Sub

Private Sub CmdDelete_Click()
'On Error Resume Next 'Error Check
    Dim r As Integer
    r = MsgBox("Delete selected step?", vbYesNo, "Delete Step")
    If r = vbYes Then
        sqlstring = "delete from TPIBK_REcipeBatchData where id = " & LstRecipes.Column(0, LstRecipes.ListIndex)
        db.execute sqlstring
        If LstRecipes.ListIndex > 0 Then LstRecipes.ListIndex = LstRecipes.ListIndex - 1
        LoadProcedure
    End If
End Sub

Private Sub CmdNew_Click()
'On Error Resume Next 'Error Check
    TxtStep.Text = highstep + 5
    Lstparameters.Clear
    newstep = True
    CmdNew.Enabled = Not newstep And (BatchStatus = "Registered") 'Modified AB 2017-08-11
    CmdDelete.Enabled = Not newstep And (BatchStatus = "Registered") 'Modified AB 2017-08-11
End Sub

Private Sub CmdSave_Click()
'On Error Resume Next 'Error Check
    Dim alloc As Integer
    Dim latebinding As Integer
    Dim Stypestring As String
    Dim Pclassstring As String
    Dim Ptypestring As String
    Dim stepstring As String
    Dim transstring As String
    Dim Matstring As String
    
    Stypestring = GetValue(CmbType, 0, "")
    Pclassstring = GetValue(CmbPClass, 0, "NULL")
    If Stypestring = "" Or ((Pclassstring = "NULL") And (CmbPClass.Enabled)) Then
        MsgBox "Inavlid step type or process", vbInformation, "Error"
        Exit Sub
    Else
        If ChkLevel.value Then alloc = alloc + CInt(ChkLevel.tag)
        If ChkMatcheck.value Then alloc = alloc + CInt(ChkMatcheck.tag)
        If ChkQA.value Then alloc = alloc + CInt(ChkQA.tag)
        
        transstring = GetValue(CmbTransition, 0, "NULL")
        stepstring = GetValue(CmbStep, 0, "NULL")
        Matstring = GetValue(CmbMaterial, 0, "NULL")
        Ptypestring = GetValue(CmbPtype, 0, "NULL")
        
'    If IsNull(CmbTransition.Column(0, CmbTransition.ListIndex)) Or CmbTransition.Column(0, CmbTransition.ListIndex) = "" Then transstring = "NULL" Else transstring = CmbTransition.Column(0, CmbTransition.ListIndex)
'    If IsNull(CmbStep.Column(0, CmbStep.ListIndex)) Or CmbStep.Column(0, CmbStep.ListIndex) = "" Then stepstring = "NULL" Else stepstring = CmbStep.Column(0, CmbStep.ListIndex)
'    If IsNull(CmbMaterial.Column(0, CmbMaterial.ListIndex)) Or CmbMaterial.Column(0, CmbMaterial.ListIndex) = "" Then Matstring = "Null" Else Matstring = CmbMaterial.Column(0, CmbMaterial.ListIndex)
'    If IsNull(CmbPtype.Column(0, CmbPtype.ListIndex)) Or CmbPtype.Column(0, CmbPtype.ListIndex) = "" Then typestring = "NULL" Else typestring = CmbPtype.Column(0, CmbPtype.ListIndex)
    End If

    If LstRecipes.ListIndex = -1 Or newstep Then
        'insert new record step
        sqlstring = "insert into tpibk_recipebatchdata(recipe_rid,recipe_version,TPIBK_Steptype_ID,processclassphase_id,step,userstring,recipeequipmenttransition_data_id,nextstep,allocation_type_id,latebinding,material_id,ProcessClass_ID) " + _
                    " output inserted.id values(" + _
                    "'" & BatchID & "'," + _
                    "'" & BatchVersion & "'," + _
                    Stypestring & "," + _
                    "" & Ptypestring & "," + _
                    "'" & TxtStep.Text & "'," + _
                    "'" & TxtString.Text & "'," + _
                    "" & transstring & "," + _
                    "" & stepstring & "," + _
                    CStr(alloc) & "," + _
                    CStr(latebinding) & "," + _
                    "" & Matstring & "," + _
                    "" & Pclassstring & "" + _
                    ")"
        'insert new record
        Dim rs As adodb.Recordset
        Set rs = db.getRecords(sqlstring)
        If Not rs.EOF Then
            selrecipeid = rs(0).value
        End If
        rs.Close
        
        sqlstring = "SELECT TPIBK_RecipeParameterData.ID, ProcessClassPhaseParameter.DefValue As Value, ProcessClassPhaseParameter.EU " + _
                        "FROM ProcessClassPhase INNER JOIN " + _
                            "ProcessClassPhaseParameter INNER JOIN " + _
                            "TPIBK_RecipeParameterData ON ProcessClassPhaseParameter.ID = TPIBK_RecipeParameterData.ProcessClassPhaseParameter_ID ON " + _
                            "ProcessClassPhase.id = ProcessClassPhaseParameter.ProcessClassPhase_ID " + _
                        "WHERE (ProcessClassPhase.id = " & Ptypestring & ")"
        Set rs = db.getRecords(sqlstring)
        While Not rs.EOF
            sqlstring = "INSERT INTO [TPIBK_RecipeStepData]([TPIBK_RecipeParameterData_ID],[TPIBK_RecipeBatchData_ID],[Value],[CustomEU]) VALUES (" & rs("ID").value & "," & selrecipeid & ",'" & rs("Value").value & "','" & rs("EU").value & "')"
            db.execute sqlstring
            rs.MoveNext
        Wend
        rs.Close
        
        
        Set rs = Nothing
    Else
        'update existing step record
        sqlstring = "update tpibk_recipebatchdata set " + _
                    "TPIBK_Steptype_ID=" & Stypestring & ", " + _
                    "processclassphase_id=" & Ptypestring & ", " + _
                    "step='" & TxtStep.Text & "', " + _
                    "userstring='" & TxtString.Text & "', " + _
                    "recipeequipmenttransition_data_id=" & transstring & ", " + _
                    "nextstep=" & stepstring & ", " + _
                    "allocation_type_id=" & CStr(alloc) & ", " + _
                    "latebinding=" & CStr(latebinding) & ", " + _
                    "material_id=" & Matstring & ", " + _
                    "ProcessClass_ID =" & Pclassstring & " " + _
                    "where id = " & LstRecipes.Column(0, LstRecipes.ListIndex)
        selrecipeid = LstRecipes.Column(0, LstRecipes.ListIndex)
        db.execute (sqlstring)

    End If
'    On Error GoTo ErrHandler

    
    newstep = False
    LoadProcedure
Exit Sub

'Error message
ErrHandler:
    MsgBox Now() & " VBA error Insert " & Err.Number & " " & Err.Description & " on display " & Name & " for " & Application.ActiveDisplay.ActiveElement.Name
End Sub

Private Sub LoadTransitions()
    sqlstring = "SELECT TOP (100) PERCENT dbo.ProcessClassTransition.ID, dbo.ProcessClassTransition.Name " + _
                "FROM         dbo.ProcessClassTransition INNER JOIN " + _
                                      "dbo.RecipeEquipmentTransition ON dbo.ProcessClassTransition.ID = dbo.RecipeEquipmentTransition.ProcessClassTransition_ID INNER JOIN " + _
                                      "dbo.RecipeEquipmentTransition_Data ON " + _
                                      "dbo.RecipeEquipmentTransition.ID = dbo.RecipeEquipmentTransition_Data.RecipeEquipmentTransition_ID INNER JOIN " + _
                                      "dbo.Transition_Index ON dbo.RecipeEquipmentTransition_Data.Transition_Index_ID = dbo.Transition_Index.ID " + _
                "WHERE (dbo.ProcessClassTransition.ProcessClass_ID = " & CmbPClass.Column(2, CmbPClass.ListIndex) & ") " + _
                "GROUP BY dbo.ProcessClassTransition.Name, dbo.ProcessClassTransition.ID " + _
                "ORDER BY dbo.ProcessClassTransition.Name"
    ListRefreshSQL rs, db, sqlstring, CmbTransition, "id,name", 0
    If CmbTransition.ListCount > 1 Then CmbTransition.RemoveItem (CmbTransition.ListCount - 1)
End Sub

Private Sub CmdTransitions_Click()
    Dim frm As New FrmTransition
    frm.SetLD = ld
    frm.Init
    frm.Show
    LoadTransitions
End Sub

Private Sub CommandButton1_Click()
''ID,Name,Description,TPIBK_RecipeParameters_ID,ProcessClassPhase_ID,ValueType,Scaled,MinValue,MaxValue,DefValue,IsMaterial,TPIBK_RecipeParameterData_ID,Value,TPIBK_RecipeStepData_ID,defEU,EU

    Dim f As New FrmParameterEdit
    Dim v As Long
    Dim propvalue As String
    Dim EU As String
    
    f.SetLD = ld
    
   
    f.Caption = Lstparameters.Column(1, Lstparameters.ListIndex)
    
    Select Case Lstparameters.Column(3, Lstparameters.ListIndex)
    Case 3
        'phase mode
        If Lstparameters.Column(11, Lstparameters.ListIndex) = "" Then
            v = Lstparameters.Column(9, Lstparameters.ListIndex)
        Else
            v = Lstparameters.Column(12, Lstparameters.ListIndex)
        End If
               
        f.Init Lstparameters.Column(3, Lstparameters.ListIndex), Lstparameters.Column(4, Lstparameters.ListIndex), v
        
'    Case 4
'        'equipment mode
'        If Lstparameters.Column(11, Lstparameters.ListIndex) = "" Then
'            v = Lstparameters.Column(9, Lstparameters.ListIndex)
'        Else
'            v = Lstparameters.Column(12, Lstparameters.ListIndex)
'        End If
'
'        f.Init Lstparameters.Column(3, Lstparameters.ListIndex), Lstparameters.Column(4, Lstparameters.ListIndex), v
    Case Else
        '    f.LblEU = LstParameters.Column(14, LstParameters.ListIndex)
        f.Lblmin = Lstparameters.Column(7, Lstparameters.ListIndex)
        f.ChkScaled.value = Lstparameters.Column(6, Lstparameters.ListIndex)
        f.LblMax = Lstparameters.Column(8, Lstparameters.ListIndex)
        f.LblType = Lstparameters.Column(5, Lstparameters.ListIndex)
                       
        If Lstparameters.Column(11, Lstparameters.ListIndex) = "" Then
            f.TxtValue = Lstparameters.Column(9, Lstparameters.ListIndex)
        Else
            f.TxtValue = Lstparameters.Column(12, Lstparameters.ListIndex)
        End If
    
        'load EU
        If Lstparameters.Column(15, Lstparameters.ListIndex) = "" Then
            f.TxtEU = Lstparameters.Column(14, Lstparameters.ListIndex)
        Else
            f.TxtEU = Lstparameters.Column(15, Lstparameters.ListIndex)
        End If
        f.Init 0, 0, 0
    End Select

    f.Show vbModal
    If f.DialogResult = vbOK And BatchStatus <> "Approved" Then
        Select Case Lstparameters.Column(3, Lstparameters.ListIndex)
        Case 3
        'phase mode
            propvalue = CStr(f.CmbMode.Column(0, f.CmbMode.ListIndex))
            EU = ""
'        Case 4
'        'eqipment mdoe
'            propvalue = CStr(f.CmbMode.Column(0, f.CmbMode.ListIndex))
'            eu = ""
        Case Else
            propvalue = f.TxtValue.Text
            EU = f.TxtEU.Text
        End Select

        If Lstparameters.Column(13, Lstparameters.ListIndex) = 0 Then 'TPIBK_RecipeStepData_ID
            'insert parameter value
            sqlstring = "insert into TPIBK_RecipeStepData(TPIBK_RecipeParameterData_ID,TPIBK_RecipeBatchData_ID,value,CustomEU) " & _
             "values(" & _
                     Lstparameters.Column(0, Lstparameters.ListIndex) & "," & _
                     LstRecipes.Column(0, LstRecipes.ListIndex) & "," & _
                    "'" & propvalue & "','" & EU & "')"
        Else
            'update parameter value
             sqlstring = "UPDATE TPIBK_RecipeStepData " & _
                            "SET value='" & propvalue & "', CustomEU='" & EU & "' " & _
                            "WHERE id=" & Lstparameters.Column(13, Lstparameters.ListIndex)
        End If
        
        
        
        db.execute (sqlstring)
        LoadParameters
    End If
       
    Unload f
End Sub

Private Sub Lstparameters_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    CommandButton1_Click
End Sub

Private Sub LstRecipes_Change()
'On Error Resume Next 'Error Check
    If BatchStatus <> "Approved" Then
        CmdNew.Enabled = Not newstep And (BatchStatus = "Registered") 'Modified AB 2017-08-11
        CmdDelete.Enabled = Not newstep And (BatchStatus = "Registered") 'Modified AB 2017-08-11
    End If

    If LstRecipes.ListIndex < 0 Then
        TxtStep.Text = 1
        Exit Sub
    End If
    If LstRecipes.Column(3, LstRecipes.ListIndex) = "" Then Exit Sub
    '0 id,
    '1 step,
    '2 Message,
    '3 TPIBK_Steptype_ID,
    '4 processclassphase_id,
    '5 step,
    '6 userstring,
    '7 recipeequipmenttransition_data_id,
    '8 nextstep,
    '9 allocation_type_id,
    '10 latebinding,
    '11 material_id
    '12 ProcessClass_ID
    newstep = False

    SelectRow CmbType, 0, LstRecipes.Column(3, LstRecipes.ListIndex), CmbType
    
    TxtStep.Text = LstRecipes.Column(5, LstRecipes.ListIndex)
    If CmbType.Column(0, CmbType.ListIndex) = "7" Then
        TxtString.Text = LstRecipes.Column(6, LstRecipes.ListIndex)
    Else
        TxtString.Text = ""
    End If
    mTxtString = TxtString

    SelectRow CmbPClass, 0, LstRecipes.Column(12, LstRecipes.ListIndex), mCmbPClass
    SelectRow CmbPtype, 0, LstRecipes.Column(4, LstRecipes.ListIndex), mCmbPtype
    SelectRow CmbMaterial, 0, LstRecipes.Column(11, LstRecipes.ListIndex), mCmbMaterial
    SelectRow CmbTransition, 0, LstRecipes.Column(7, LstRecipes.ListIndex), mCmbTransition
    SelectRow CmbStep, 0, LstRecipes.Column(8, LstRecipes.ListIndex), mCmbStep

    If (LstRecipes.Column(9, LstRecipes.ListIndex) And CInt(ChkLevel.tag)) > 0 Then ChkLevel.value = 1 Else ChkLevel.value = 0
    If (LstRecipes.Column(9, LstRecipes.ListIndex) And CInt(ChkMatcheck.tag)) > 0 Then ChkMatcheck.value = 1 Else ChkMatcheck.value = 0
    If (LstRecipes.Column(9, LstRecipes.ListIndex) And CInt(ChkQA.tag)) > 0 Then ChkQA.value = 1 Else ChkQA.value = 0
    
    If CmbType.Column(0, CmbType.ListIndex) = "8" And ChkMatcheck.value = True Then
        CmbMaterial.Enabled = (BatchStatus = "Registered") 'Modified AB 2017-08-11
    End If
    
    If LstRecipes.ListIndex >= 0 Then LoadParameters
'    Select Case CmbType.Column(0, CmbType.ListIndex)
'        Case 1 'run phase
'        Case 2 'start phase
'        Case 3 'end phase
'        Case 4 'check phase
'        Case 5 'download parameters
'        Case 6 'transition
'        Case 7 'operator confirmation
'        Case 8 'allocate unit
'        Case 9 'deallocate unit
'        Case 10 'jump
'    End Select
End Sub

Private Sub TxtStep_Change()
    If TxtStep.value <> "" Then
        If Application.Name = "Display Client" Then
            If CInt(TxtStep.value) > 999 Then
                MsgBox "Out of Range"
                TxtStep.value = 0
            End If
        Else
        End If
    End If
    EnableSave 'Enable save button
End Sub

Public Sub LoadClass()
    sqlstring = "SELECT RER.ID, RER.ProcessClass_Name, case when ROW_NUMBER() over(partition by RER.ProcessClass_Name order by RER.ProcessClass_Name)<2 then RER.ProcessClass_Name else RER.ProcessClass_Name+' #'+ltrim(ROW_NUMBER() over(partition by RER.ProcessClass_Name order by RER.ProcessClass_Name)) end as message,PC.id as pclass_Id, " & _
                "case coalesce(Equipment_Name,PC.Description) When '' Then PC.Description Else Equipment_Name End As Equipment_Name " & _
                "FROM RecipeEquipmentRequirement RER INNER JOIN ProcessClass PC ON RER.ProcessClass_Name = PC.Name " & _
                "WHERE Recipe_RID= '" & BatchID & "' " & _
                "and   Recipe_Version=" & BatchVersion

    ListRefreshSQL rs, db, sqlstring, CmbPClass, "id,message,pclass_Id,Equipment_Name", 0
    If CmbPClass.ListCount > 0 Then CmbPClass.RemoveItem (CmbPClass.ListCount - 1)
End Sub

Private Sub CmbMaterial_Change()
'On Error Resume Next 'Error Check
    EnableSave 'Enable save button
End Sub

Private Sub CmbStep_Change()
'On Error Resume Next 'Error Check
    EnableSave 'Enable save button
End Sub

Private Sub CmbTransition_Change()
'On Error Resume Next 'Error Check
    EnableSave 'Enable save button
End Sub

Private Sub TxtStep_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    CheckNumericInput KeyAscii, TxtStep
'    If TxtStep.value <> "" Then
'        If CInt(TxtStep.value) > 128 Then KeyAscii = 0
'    End If
End Sub

Private Sub TxtString_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    CheckSpecialInput KeyAscii, TxtString
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

Public Property Get BatchStatus() As String
    BatchStatus = Bstat
End Property

Public Property Let BatchStatus(ByVal vNewValue As String)
    Bstat = vNewValue
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
