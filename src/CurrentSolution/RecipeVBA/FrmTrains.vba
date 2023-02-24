
Option Explicit
'Option Compare Text

Private rs As adodb.Recordset
Private db As dbConnect
Private ld As LocalDatabase

Private valuetofind As Long
Private sqlunits As String
Private sqlselected As String
Private UserCodeValue As String

Public Sub Init()
'On Error GoTo ErrHandler 'Error Check
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
    CmbTrain.ColumnCount = 3
    CmbTrain.ColumnWidths = "0,1,0"

    CmbPClass.Style = fmStyleDropDownList
    CmbPClass.BoundColumn = 0
    CmbPClass.ListWidth = 140
    CmbPClass.ColumnCount = 3
    CmbPClass.ColumnWidths = "0,0,1"

    LstUnits.BoundColumn = 0
    LstUnits.ListWidth = LstUnits.Width - 5
    LstUnits.ColumnCount = 3
    LstUnits.ColumnWidths = "0," & CStr(LstUnits.ListWidth) & ",0"

    LstSelected.BoundColumn = 0
    LstSelected.ListWidth = LstSelected.Width - 5
    LstSelected.ColumnCount = 5
    LstSelected.ColumnWidths = "0," & CStr(LstSelected.ListWidth) & ",0,0,0"

    doRefresh
Exit Sub
   
'Error message
ErrHandler:
    'MsgBox "VBA error Animation Start " & Err.Number & " " & Err.Description & " on display " & Name
End Sub

Private Sub CmbTrain_Change()
'On Error Resume Next
   If CmbTrain.ListIndex > -1 Then RefreshLists
'   if Not IsNull(CmbTrain.Column(0, CmbTrain.ListIndex)) Or CmbTrain.Column(0, CmbTrain.ListIndex) <> "" Then RefreshLists
End Sub

Private Sub CmbPClass_Change()
'On Error Resume Next
   If CmbPClass.ListIndex > -1 Then RefreshLists
End Sub
    
Private Sub RefreshLists()
    Dim i As Integer
    LstSelected.Clear
    LstUnits.Clear
    CmdAdd.Enabled = False
    
    If CmbTrain.ListIndex >= 0 And CmbPClass.ListIndex >= 0 Then
       
        sqlunits = "SELECT     TOP (100) PERCENT dbo.Equipment.ID AS Equipment_ID, dbo.Equipment.Name AS Equipment_Name, dbo.ProcessClass.ID as PID " & _
                    "FROM         dbo.ProcessClass INNER JOIN " & _
                                          "dbo.Equipment ON dbo.ProcessClass.ID = dbo.Equipment.ProcessClass_ID " & _
                    "WHERE     (dbo.ProcessClass.TypeBatchKernel = 1) AND (dbo.ProcessClass.ID = " & CmbPClass.Column(0, CmbPClass.ListIndex) & ") AND (NOT (dbo.Equipment.ID IN " & _
                                              "(SELECT     Equipment_ID " & _
                                                "FROM          dbo.RecipeTrainEquipment AS RecipeTrainEquipment_1 " & _
                                                "WHERE      (RecipeTrain_ID = " & CmbTrain.Column(0, CmbTrain.ListIndex) & ")))) OR " & _
                                          "(dbo.ProcessClass.TypeRecipeHandler = 1) AND (dbo.ProcessClass.ID = " & CmbPClass.Column(0, CmbPClass.ListIndex) & ") AND (NOT (dbo.Equipment.ID IN " & _
                                              "(SELECT     Equipment_ID " & _
                                                "FROM          dbo.RecipeTrainEquipment AS RecipeTrainEquipment_2 " & _
                                                "WHERE      (RecipeTrain_ID = " & CmbTrain.Column(0, CmbTrain.ListIndex) & ")))) " & _
                        "ORDER BY Equipment_Name"
                    
        sqlselected = "SELECT     TOP (100) PERCENT dbo.Equipment.ID AS Equipment_ID, + CAST(EqIdx1 as nvarchar) + ' - ' + dbo.Equipment.Name AS Equipment_Name, dbo.ProcessClass.ID as pclass_id, " & _
                                          "dbo.RecipeTrainEquipment.RecipeTrain_ID,dbo.RecipeTrainEquipment.ID " & _
                    "FROM         dbo.ProcessClass INNER JOIN " & _
                                          "dbo.Equipment ON dbo.ProcessClass.ID = dbo.Equipment.ProcessClass_ID INNER JOIN " & _
                                          "dbo.RecipeTrainEquipment ON dbo.Equipment.ID = dbo.RecipeTrainEquipment.Equipment_ID " & _
                    "WHERE     (dbo.ProcessClass.TypeBatchKernel = 1) AND (dbo.RecipeTrainEquipment.RecipeTrain_ID = " & CmbTrain.Column(0, CmbTrain.ListIndex) & ") OR " & _
                                          "(dbo.ProcessClass.TypeRecipeHandler = 1) AND (dbo.ProcessClass.ID = " & CmbPClass.Column(0, CmbPClass.ListIndex) & ") AND (dbo.RecipeTrainEquipment.RecipeTrain_ID = " & CmbTrain.Column(0, CmbTrain.ListIndex) & ") " & _
                    "ORDER BY Equipment_Name"

        ListRefreshSQL rs, db, sqlunits, LstUnits, "Equipment_ID,Equipment_Name,PID", 0
        ListRefreshSQL rs, db, sqlselected, LstSelected, "Equipment_ID,Equipment_Name,pclass_id,RecipeTrain_ID,ID", 0
        On Error Resume Next
        LstUnits.RemoveItem (LstUnits.ListCount - 1)
        LstSelected.RemoveItem (LstSelected.ListCount - 1) 'On Error GoTo ErrHandler
        For i = 0 To LstSelected.ListCount - 1
            If LstSelected.Column(2, i) = CmbPClass.Column(0, CmbPClass.ListIndex) Then
            
            'GoTo found
            End If
        Next i
        CmdAdd.Enabled = True
found:
    End If
Exit Sub
ErrHandler:
    MsgBox "Failed to load data", vbCritical, "Error"
End Sub


Private Sub CmdUpdate_Click()
    db.execute "update plc set name ='" & TxtName.Text & "', plc_type_id=" & CStr(CmbPClass.Column(0, CmbPClass.ListIndex)) & " where id =" & CmbTrain.Column(0, CmbTrain.ListIndex)
    valuetofind = CmbTrain.Column(0, CmbTrain.ListIndex)
    doRefresh
End Sub

Private Sub CmdDeleteRecipe_Click()
    db.execute "delete from RecipeTrain where id = " & CmbTrain.Column(0, CmbTrain.ListIndex)
    doRefresh
    RefreshLists
End Sub

Private Sub CmdSave_Click()

    On Error GoTo ErrHandler
    
    If TxtName.Text = "" Then
        MsgBox "Must enter a name"
        Exit Sub
    End If
        
    Dim newid As Long
    Dim sqlstring As String
    'create new record
    db.execute "insert into RecipeTrain (name,description) values('" & TxtName.Text & "','" & TxtDesc.Text & "')"
    'find new record
    sqlstring = "SELECT ID from Recipetrain where name='" & TxtName.Text & "'"
    
    Set rs = db.getRecords(sqlstring)
    If Not rs Is Nothing And Not rs.EOF Then
       newid = rs.Fields(0).value
    End If

    rs.Close
    Set rs = Nothing
    'if the new train is a copy get the existing records and insert them
    If CmbTrain.ListIndex >= 0 Then
        If Not IsNull(CmbTrain.Column(0, CmbTrain.ListIndex)) And CmbTrain.ListIndex <> "" Then
        sqlstring = "insert into recipetrainequipment (recipetrain_id,equipment_id,eqidx1,eqidx2) select " & CStr(newid) & " as recipetrain_id, equipment_id,eqidx1,eqidx2 from recipetrainequipment where recipetrain_id=" & CStr(CmbTrain.Column(0, CmbTrain.ListIndex))
        db.execute (sqlstring)
        End If
    End If
    doRefresh
    RefreshLists
    Frame2.Visible = False
    Exit Sub
    'Error message
ErrHandler:
    MsgBox "VBA error SaveAs " & Err.Number & " " & Err.Description & " on display " & Name

End Sub

Private Sub CmdSaveAs_Click()
    Frame2.Visible = Not Frame2.Visible
End Sub

Private Sub TxtName_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    CheckSpecialInput KeyAscii, TxtName
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub
Private Sub CmdAdd_Click()
    Dim typecount As Integer
    Dim sqlstring As String
    
    If LstUnits.ListIndex = -1 Then Exit Sub
    
    sqlstring = "SELECT   COUNT(dbo.RecipeTrainEquipment.ID) AS Num " & _
                "FROM         dbo.Equipment INNER JOIN " & _
                                      "dbo.RecipeTrainEquipment ON dbo.Equipment.ID = dbo.RecipeTrainEquipment.Equipment_ID " & _
                "WHERE (dbo.RecipeTrainEquipment.RecipeTrain_ID = " & CmbTrain.Column(0, CmbTrain.ListIndex) & ") And (dbo.Equipment.ProcessClass_ID = " & LstUnits.Column(2, LstUnits.ListIndex) & ")"
    
    typecount = 0
    'check to see if an equipment of the same type already exists for the train
    Set rs = db.getRecords(sqlstring)
    If Not rs.EOF Then
       typecount = rs.Fields(0).value + 1
    End If
   
    rs.Close
    Set rs = Nothing

    db.execute "INSERT INTO [RecipeTrainEquipment] ([RecipeTrain_ID], [Equipment_ID], [EqIdx1], [EqIdx2])  values(" & _
    CmbTrain.Column(0, CmbTrain.ListIndex) & "," & _
    LstUnits.Column(0, LstUnits.ListIndex) & "," & _
    CStr(typecount) & _
    ",0)"
    RefreshLists
End Sub

Private Sub CmdDelete_Click()
    If LstSelected.ListIndex = -1 Then Exit Sub
    db.execute "delete from RecipeTrainEquipment where id = " & LstSelected.Column(4, LstSelected.ListIndex)
    RefreshLists
End Sub

Public Sub doRefresh()

    Dim sqltrain As String
    Dim sqlclass As String
    
    CmbTrain.Clear
    CmbPClass.Clear
    
    sqlclass = "SELECT dbo.ProcessClass.ID as ID, dbo.ProcessClass.Name AS ptClass, Max(dbo.ProcessClass.Description) Description " & _
                "FROM         dbo.ProcessClass INNER JOIN " & _
                                      "dbo.Equipment ON dbo.ProcessClass.ID = dbo.Equipment.ProcessClass_ID " & _
                "WHERE     (dbo.ProcessClass.TypeBatchKernel = 1) OR " & _
                                      "(dbo.ProcessClass.TypeRecipeHandler = 1) " & _
                "GROUP BY dbo.ProcessClass.ID,dbo.ProcessClass.Name " & _
                "ORDER BY Description" 'ptClass"
    
    sqltrain = "select ID, Name, Description from RecipeTrain order by Name ASC"

    ListRefreshSQL rs, db, sqlclass, CmbPClass, "id,ptClass,Description", 0
    ListRefreshSQL rs, db, sqltrain, CmbTrain, "ID,Name,Description", 0
    On Error Resume Next
    
    CmbPClass.RemoveItem (CmbPClass.ListCount - 1)
    CmbTrain.RemoveItem (CmbTrain.ListCount - 1)
    
    RefreshLists
End Sub

Public Property Let SetLD(ByVal vNewValue As LocalDatabase)
    Set ld = vNewValue
End Property

Public Property Get UserCode() As String
'On Error Resume Next 'Error Check
    UserCode = UserCodeValue
End Property

Public Property Let UserCode(ByVal vNewValue As String)
'On Error Resume Next 'Error Check
    UserCodeValue = vNewValue
End Property

Private Sub UserForm_Activate()
    CmdSave.Enabled = Application.LoadedDisplays.Item("footer").lUserCodes Like "*" & UserCode & "*" Or Application.LoadedDisplays.Item("footer").lUserCodes = "*"
    CmdSaveAs.Enabled = CmdSave.Enabled
    CmdDelete.Enabled = CmdSave.Enabled
    CmdDeleteRecipe.Enabled = CmdSave.Enabled
    CmdAdd.Enabled = CmdSave.Enabled
End Sub

