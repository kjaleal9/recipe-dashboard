
Option Explicit
'Option Compare Text

Private rs As adodb.Recordset
Private rs1 As adodb.Recordset
Private rstype As adodb.Recordset
Private db As dbConnect
Private sqlstring As String
Private ld As LocalDatabase

Private Sub CmdAdd_Click()
    If LstSelected.Column(0, LstSelected.ListIndex) <> 0 Then
        MsgBox "Invalid selection"
        Exit Sub
    End If
    sqlstring = "insert into RecipeEquipmentTransition_Data(RecipeEquipmentTransition_ID,Transition_Index_ID) values(" & LstAvailable.Column(0, LstAvailable.ListIndex) & "," & LstSelected.Column(2, LstSelected.ListIndex) & ")"
    db.execute (sqlstring)
    
    LoadLists
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub CmdDelete_Click()
  If LstSelected.Column(0, LstSelected.ListIndex) = 0 Then
        MsgBox "Invalid selection"
        Exit Sub
    End If
    sqlstring = "delete from RecipeEquipmentTransition_Data where id = " & LstSelected.Column(0, LstSelected.ListIndex)
    db.execute (sqlstring)
    
    LoadLists
End Sub

Public Sub Init()

'On Error Resume Next 'Error Check

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
    
    LstAvailable.BoundColumn = 0
    LstAvailable.ListWidth = LstAvailable.Width - 5
    LstAvailable.ColumnCount = 2
    LstAvailable.ColumnWidths = "0,1"
    
    LstSelected.BoundColumn = 0
    LstSelected.ListWidth = LstSelected.Width - 5
    LstSelected.ColumnCount = 3
    LstSelected.ColumnWidths = "0,1,0"
    
    LoadLists
End Sub

Private Sub LoadLists()
    sqlstring = "SELECT     dbo.RecipeEquipmentTransition.ID, dbo.ProcessClassTransition.ProcessClass_ID, dbo.RecipeEquipmentTransition.Name " + _
                "FROM         dbo.ProcessClassTransition INNER JOIN " + _
                                      "dbo.RecipeEquipmentTransition ON dbo.ProcessClassTransition.ID = dbo.RecipeEquipmentTransition.ProcessClassTransition_ID " + _
                "WHERE     (NOT (dbo.RecipeEquipmentTransition.ID IN " + _
                                          "(SELECT     RecipeEquipmentTransition_ID " + _
                                            "FROM          dbo.RecipeEquipmentTransition_Data))) order by name"
    ListRefreshSQL rs, db, sqlstring, LstAvailable, "id,name", 0
    LstAvailable.RemoveItem (LstAvailable.ListCount - 1)
    
        sqlstring = "SELECT     COALESCE (dbo.RecipeEquipmentTransition_Data.ID, 0) AS ID, LTRIM(dbo.Transition_Index.ID) + ' ' + COALESCE (dbo.RecipeEquipmentTransition.Name,'') AS Name, dbo.Transition_Index.ID as Index_ID " + _
                    "FROM         dbo.ProcessClassTransition INNER JOIN " + _
                                          "dbo.RecipeEquipmentTransition ON dbo.ProcessClassTransition.ID = dbo.RecipeEquipmentTransition.ProcessClassTransition_ID INNER JOIN " + _
                                          "dbo.RecipeEquipmentTransition_Data ON " + _
                                          "dbo.RecipeEquipmentTransition.ID = dbo.RecipeEquipmentTransition_Data.RecipeEquipmentTransition_ID RIGHT OUTER JOIN " + _
                                          "dbo.Transition_Index ON dbo.RecipeEquipmentTransition_Data.Transition_Index_ID = dbo.Transition_Index.ID"
    ListRefreshSQL rs, db, sqlstring, LstSelected, "id,name,index_id", 0
    LstSelected.RemoveItem (LstSelected.ListCount - 1)
    LstSelected.ListIndex = 0
End Sub
Public Property Let SetLD(ByVal vNewValue As LocalDatabase)
    Set ld = vNewValue
End Property
