
Option Explicit
'Option Compare Text

Private rs As adodb.Recordset
Private db As dbConnect
Private valuetofind As Long
Private UserCodeValue As String
Private ld As LocalDatabase

Public Sub Init()
On Error GoTo ErrHandler 'Error Check
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

    Dim status(5) As String
    status(0) = "*"
    status(1) = "Registered"
    status(2) = "Valid"
    status(3) = "Approved"
    status(4) = "Obsolete"
    CmbStatus.List = status
    CmbStatus.ListIndex = 3
    CmbStatus.RemoveItem (CmbStatus.ListCount - 1)
    
    LstRecipes.BoundColumn = 0
    LstRecipes.ListWidth = LstRecipes.Width - 5
    LstRecipes.ColumnCount = 8
    LstRecipes.ColumnWidths = "75,20,100,40,125,50,40,10" 'RID,Version,VersionDate,RecipeType,Description,Status,ProductID,MaterialName
    TxtName.Text = "*"
Exit Sub

'Error message
ErrHandler:
'    MsgBox "VBA error Animation Start " & Err.Number & " " & Err.Description & " on display " & name
End Sub


Private Sub CmdExport_Click()
    Dim frm As New FrmExport
    If LstRecipes.ListIndex = -1 Then Exit Sub
    frm.BatchID = LstRecipes.Column(0, LstRecipes.ListIndex)
    frm.BatchVersion = LstRecipes.Column(1, LstRecipes.ListIndex)
    frm.UserCode = UserCode
    frm.SetLD = ld
    frm.Init
    
    frm.Show

End Sub

Private Sub CmdImport_Click()
 Dim frm As New FrmImport
     
    frm.UserCode = UserCode
    frm.SetLD = ld
    frm.Init
    
    frm.Show
End Sub

Private Sub LstRecipes_AfterUpdate()
    If LstRecipes.ListIndex >= 0 Then
        CmdSaveAs.Enabled = CmdNew.Enabled
        CmdDelete.Enabled = CmdNew.Enabled
        CmdExport.Enabled = CmdNew.Enabled
    Else
        CmdSaveAs.Enabled = False
        CmdDelete.Enabled = False
        CmdExport.Enabled = False
    End If
End Sub

Private Sub UserForm_Activate()
'On Error Resume Next 'Error Check
'    If Application.name = "Microsoft Excel" Then
'        CmdNew.Enabled = True
'        CmdSaveAs.Enabled = True
'        CmdDelete.Enabled = True
'        LstRecipes.Locked = False
'    Else
'''''        CmdNew.Enabled = ThisDisplay.lUserCodes Like "*" & UserCode & "*" Or ThisDisplay.lUserCodes = "*"
'''''        CmdSaveAs.Enabled = ThisDisplay.lUserCodes Like "*" & UserCode & "*" Or ThisDisplay.lUserCodes = "*"
'''''        CmdDelete.Enabled = ThisDisplay.lUserCodes Like "*" & UserCode & "*" Or ThisDisplay.lUserCodes = "*"
'''''        LstRecipes.Locked = Not (ThisDisplay.lUserCodes Like "*" & UserCode & "*" Or ThisDisplay.lUserCodes = "*")
        CmdNew.Enabled = Application.LoadedDisplays.Item("footer").lUserCodes Like "*" & UserCode & "*" Or Application.LoadedDisplays.Item("footer").lUserCodes = "*"
'        CmdSaveAs.Enabled = CmdNew.Enabled
'        CmdDelete.Enabled = CmdNew.Enabled
'        CmdExport.Enabled = CmdNew.Enabled
        CmdImport.Enabled = CmdNew.Enabled
        LstRecipes.Locked = Not (Application.LoadedDisplays.Item("footer").lUserCodes Like "*" & UserCode & "*" Or Application.LoadedDisplays.Item("footer").lUserCodes = "*")
'    End If
End Sub

Public Property Get UserCode() As String
'On Error Resume Next 'Error Check
    UserCode = UserCodeValue
End Property

Public Property Let UserCode(ByVal vNewValue As String)
'On Error Resume Next 'Error Check
    UserCodeValue = vNewValue
End Property

Private Sub CmdSearch_Click()
On Error Resume Next 'Error Check
    Call RefreshRecipes
End Sub

Public Sub RefreshRecipes()
    Dim sqlstring As String
    Dim replacements(10) As String
    Dim tempstring As String
    
    If cbShowAll = False Then
'''''        sqlstring = "select Recipe.RID, max(convert(int,Recipe.Version)) as Version, Recipe.RecipeType, Recipe.Description, Recipe.Status, Recipe.ProductID, Material.Name as MaterialName, max(VersionDate) as VersionDate, 'Obsolete' as 'Obsolete' from Recipe "
        sqlstring = "select Recipe.RID, max(convert(int,Recipe.Version)) as Version, Recipe.RecipeType, Recipe.Description, Recipe.Status, Recipe.ProductID, Material.Name as MaterialName, max(VersionDate) as VersionDate, 'Obsolete' as 'Obsolete' from Recipe " & _
        "left outer join Material on Recipe.ProductID=Material.SiteMaterialAlias " & _
        "where RecipeType='Master'  %%1  %%2 group by RID, RecipeType, Recipe.Description, Recipe.Status, Recipe.ProductID, Material.Name  order by RID, max(VersionDate) DESC"
'''''        "left outer join BatchListEntry on Recipe.RID=BatchListEntry.Recipe_RID and Recipe.Version=BatchListEntry.Recipe_Version  "
    Else
'''''        sqlstring = "select Recipe.RID,convert(int,Recipe.Version) as Version, Recipe.RecipeType, Recipe.Description, Recipe.Status, Recipe.ProductID, Material.Name as MaterialName, VersionDate, BatchID, 'Obsolete' as 'Obsolete' from Recipe "
        sqlstring = "select Recipe.RID,convert(int,Recipe.Version) as Version, Recipe.RecipeType, Recipe.Description, Recipe.Status, Recipe.ProductID, Material.Name as MaterialName, VersionDate, 'Obsolete' as 'Obsolete' from Recipe " & _
        "left outer join Material on Recipe.ProductID=Material.SiteMaterialAlias " & _
        "where RecipeType='Master'  %%1  %%2 order by RID, VersionDate DESC"
'''''        "left outer join BatchListEntry on Recipe.RID=BatchListEntry.Recipe_RID and Recipe.Version=BatchListEntry.Recipe_Version  "
    End If

    If TxtName.TextLength = 0 Then
        replacements(1) = " "
    Else
        If InStr(1, TxtName.Text, "*") Then
            tempstring = Replace(TxtName.Text, "*", "%")
            replacements(1) = " and RID like '" & tempstring & "' "
        Else
            replacements(1) = " and RID= '" & TxtName.Text & "' "
        End If
    End If
    If CmbStatus.Text = "*" Then
        replacements(2) = " "
    Else
        replacements(2) = " and Recipe.Status= '" & CmbStatus.Text & "' "
    End If

    sqlstring = Placeholder(sqlstring, replacements)
    ListRefreshSQL rs, db, sqlstring, LstRecipes, "RID,Version,VersionDate,RecipeType,Description,Status,ProductID,MaterialName", 0
    
End Sub

Private Sub CmdNew_Click()
    Dim frm As New FrmRecipeHeader
'    frm.NewBatch
    frm.UserCode = UserCode
    frm.SetLD = ld
    frm.Init
    frm.LoadBatch "New"

    frm.Show
    RefreshRecipes
End Sub

Private Sub CmdSaveAs_Click()

'Create a new revision of the current recipe
    Dim sqlstring As String
    
    Dim frm As New FrmNewName
    frm.Show vbModal
    If frm.DialogResult = vbOK Then
    sqlstring = "INSERT INTO [Recipe] " + _
             "select " + _
                 "'" & frm.TxtNewName.Text & "' " + _
                 ", '1' " + _
                 ",[RecipeType],[NbrOfExecutions],[VersionDate],[Description],[EffectiveDate],[ExpirationDate],[ProductID],[BatchSizeNominal],[BatchSizeMin],[BatchSizeMax]" + _
                 ",'Registered', 1, NULL, NULL, 0, 0 from Recipe " + _
                 "where RID ='" & LstRecipes.Column(0, LstRecipes.ListIndex) & "' and Version = '" & LstRecipes.Column(1, LstRecipes.ListIndex) & "'"
    db.execute sqlstring
    'Copy the selected units to the new reicpe revision
    db.execute "TPIBK_CopyRecipe '" & LstRecipes.Column(0, LstRecipes.ListIndex) & "', '" & LstRecipes.Column(1, LstRecipes.ListIndex) & "', '" & frm.TxtNewName.Text & "', '1'"
    RefreshRecipes
    End If
       
    Unload frm
    
    
End Sub

Private Sub CmdDelete_Click()
'On Error Resume Next 'Error Check
    'Define local variables
    Dim BatchID As String
    Dim BatchVersion, sqlstring
    Dim version As String
    Dim r As Integer
    
    BatchID = LstRecipes.Column(0, LstRecipes.ListIndex)
    BatchVersion = LstRecipes.Column(1, LstRecipes.ListIndex)
        
    r = MsgBox("Are you sure you want to delete recipe " & BatchID & "?", vbYesNo, "Delete Recipe")
    If r = vbYes Then
        sqlstring = "DELETE [dbo].[Recipe] " + _
                    "WHERE RID = '" & BatchID & "' and Version  = '" & BatchVersion & "'"
        db.execute sqlstring
        RefreshRecipes
    End If
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub LstRecipes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim frm As New FrmRecipeHeader
    
    frm.BatchID = LstRecipes.Column(0, LstRecipes.ListIndex)
    frm.BatchVersion = LstRecipes.Column(1, LstRecipes.ListIndex)
    frm.UserCode = UserCode
    frm.SetLD = ld
    frm.Init
    frm.LoadBatch "View"

    frm.Show
    CmdSearch_Click
End Sub

Public Property Let SetLD(ByVal vNewValue As LocalDatabase)
    Set ld = vNewValue
End Property
