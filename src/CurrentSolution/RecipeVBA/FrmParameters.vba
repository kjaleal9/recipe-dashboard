
Option Explicit

Private rs As adodb.Recordset
Private rs1 As adodb.Recordset
Private db As dbConnect
Private valuetofind As Long
Private sqlstring As String
Private ld As LocalDatabase

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
        
    LstDefinition.BoundColumn = 0
    LstDefinition.ListWidth = 140
    LstDefinition.ColumnCount = 3
    LstDefinition.ColumnWidths = "0,1,0"
    
    LstParameters.BoundColumn = 0
    LstParameters.ListWidth = 140
    LstParameters.ColumnCount = 2
    LstParameters.ColumnWidths = "0,1"
    LstParameters.MultiSelect = fmMultiSelectExtended
    
    LstSelected.BoundColumn = 0
    LstSelected.ListWidth = 140
    LstSelected.ColumnCount = 2
    LstSelected.ColumnWidths = "0,1"
    LstSelected.MultiSelect = fmMultiSelectExtended
    
    TxtName.Text = "*"

    sqlstring = "SELECT TOP 1000 [ID],[Name],[DataType] FROM [dbo].[TPIBK_RecipeParameters]"
                                    
    ListRefreshSQL rs, db, sqlstring, LstDefinition, "id,name,datatype", 0
    LstDefinition.RemoveItem (LstDefinition.ListCount - 1)

'Error message
ErrHandler:
   ' MsgBox "VBA error Animation Start " & Err.Number & " " & Err.Description & " on display " & name
End Sub

Private Sub CmdAdd_Click()
    Dim i As Integer
    For i = 0 To LstParameters.ListCount - 1
        If LstParameters.Selected(i) Then
            sqlstring = "insert into TPIBK_RecipeParameterData(TPIBK_RecipeParameters_ID,ProcessClassPhaseParameter_ID) values(" & LstDefinition.Column(0, LstDefinition.ListIndex) & "," & LstParameters.Column(0, i) & ")"
            db.execute (sqlstring)
        End If
    Next i
    
    LoadLists
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub CmdDelete_Click()
    Dim i As Integer
    For i = 0 To LstSelected.ListCount - 1
        If LstSelected.Selected(i) Then
            sqlstring = "delete from TPIBK_RecipeParameterData where id = " & LstSelected.Column(0, i)
            db.execute (sqlstring)
        End If
    Next i

    LoadLists
End Sub

Private Sub LstDefinition_Change()
    LoadLists
End Sub

Private Sub TxtName_Change()
    LoadLists
End Sub

Private Sub LoadLists()
'    CmdAdd.Enabled = False
'    CmdDelete.Enabled = False
    
    If LstDefinition.ListIndex < 0 Then Exit Sub
    
    sqlstring = "SELECT     ProcessClassPhaseParameter.ID, ProcessClass.Name + ProcessClassPhase.Name + ProcessClassPhaseParameter.Name AS Name " + _
                            "FROM         ProcessClassPhaseParameter INNER JOIN " + _
                                                  "ProcessClassPhase ON ProcessClassPhaseParameter.ProcessClassPhase_ID = ProcessClassPhase.ID INNER JOIN " + _
                                                  "ProcessClass ON ProcessClassPhase.ProcessClass_ID = ProcessClass.ID " + _
                            "WHERE     (NOT (ProcessClassPhaseParameter.ID IN " + _
                                                      "(SELECT     ProcessClassPhaseParameter_ID " + _
                                                        "FROM          TPIBK_RecipeParameterData AS TPIBK_RecipeParameterData_1))) AND (ProcessClassPhaseParameter.TypeRecipePhaseParameter = 1) " + _
                                                  "AND (ProcessClass.Name + ProcessClassPhase.Name + ProcessClassPhaseParameter.Name like '" & Replace(TxtName.Text, "*", "%") & "') AND (ProcessClassPhaseParameter.ValueType = N'" & LstDefinition.Column(2, LstDefinition.ListIndex) & "') order by name"
    ListRefreshSQL rs, db, sqlstring, LstParameters, "id,name", 0
    
    sqlstring = "SELECT     TPIBK_RecipeParameterData.ID, ProcessClass.Name + ProcessClassPhase.Name + ProcessClassPhaseParameter.Name AS Name, " + _
                                                  "ProcessClassPhaseParameter.ValueType " + _
                            "FROM         ProcessClassPhaseParameter INNER JOIN " + _
                                                  "ProcessClassPhase ON ProcessClassPhaseParameter.ProcessClassPhase_ID = ProcessClassPhase.ID INNER JOIN " + _
                                                  "ProcessClass ON ProcessClassPhase.ProcessClass_ID = ProcessClass.ID INNER JOIN " + _
                                                  "TPIBK_RecipeParameterData ON ProcessClassPhaseParameter.ID = TPIBK_RecipeParameterData.ProcessClassPhaseParameter_ID " + _
                            "WHERE        (TPIBK_RecipeParameterData.TPIBK_RecipeParameters_ID = " & LstDefinition.Column(0, LstDefinition.ListIndex) & ") AND (ProcessClassPhaseParameter.ValueType = N'" & LstDefinition.Column(2, LstDefinition.ListIndex) & "') order by name"
    ListRefreshSQL rs, db, sqlstring, LstSelected, "id,name", 0
    
    LstParameters.RemoveItem (LstParameters.ListCount - 1)
    LstSelected.RemoveItem (LstSelected.ListCount - 1)
End Sub

Public Property Let SetLD(ByVal vNewValue As LocalDatabase)
    Set ld = vNewValue
End Property

