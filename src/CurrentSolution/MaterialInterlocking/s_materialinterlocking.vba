Option Explicit
'              Tetra PlantMaster
'
'Function:     material interlocking.gfx
'v6.4.30_01
'Created by:   Carey Warren
'Modified by: Carey Warren
'Date: 2015-10-23
'------------------------------------------------------------------------------
Private footer As Display
Private Tag_Group As TagGroup
Private TagsInError As StringList
Private PLCTag As String

Private cgroups As Collection
Private igroups As Integer
Private igroup As Integer

Private jgroups As Integer
Private jgroup As Integer

Private groupspagecount As Integer
Private grouppagecount As Integer
Private selectedgroups As Long
Private selectedgroup As Integer

Private cMatClass As Collection
Private iMatClass As Integer
Private iMat As Integer

Private jMatClass As Integer
Private jMat As Integer

Private MatClasspagecount As Integer
Private Matpagecount As Integer
Private selectedMatClass As String
Private selectedMat As Collection
Private selectedMatIndex As Integer
Private cMat As Collection

Private num As Integer
Private Sub Display_AnimationStart()

    'Declare local variables
    Dim UserName As String
    Dim lGroup As Long
    Dim e As String
    Dim grouplist(40) As String
    Dim i As Integer
    
    
     'write the display name to the current display to trigger navigation
    If Me.TagParameters.Count > 0 Then
        Dim tg As TagGroup
        Dim s As String
        Dim t As Tag
        s = "Clients\" & Me.TagParameters(20) & "\CurrentDisplay"
        Set tg = Application.CreateTagGroup(Me.AreaName, 500)
        tg.Add s
        tg.Active = True
        Set t = tg.Item(s)
        t.Value = Me.Name
        Set t = Nothing
        Set tg = Nothing
    End If
    
    PLCTag = "/Area1/DS1::[PLC01]_STD_GF_MaterialConfig["

    'Error check
    On Error GoTo ErrHandler
    
    'Global data from footer
    Set footer = Me.Application.LoadedDisplays.Item("footer")
  
    'set animation for group selection
    e = "Group" & CStr(Me.TagParameters(1)) & CStr(Me.TagParameters(2))
    Me.Elements(e).BackColor = ColorConstants.vbBlue
        
    num = 10 ' set to the number of buttons i.e. 10
    
    footer.GetMaterialClassCollection cMatClass
    iMatClass = 0

    'Read collection from database
    footer.GetMaterialGroupCollection cgroups
    igroups = 0
    
    groupspagecount = 4 ' 40 groups
    grouppagecount = 2 ' 20 materials
    
    If cMatClass.Count Mod num = 0 Then
        MatClasspagecount = Int(cMatClass.Count / num)
    Else
        MatClasspagecount = Int(cMatClass.Count / num) + 1
    End If
       
    
    ShowGroups
    ShowMaterialClass

Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name, ftDiagSeverityError
End Sub

Private Sub cmdAssign_Released()
    Dim m As Collection
    
    'Error check
    On Error GoTo ErrHandler
        
    Set m = selectedMat.Item(selectedMatIndex)
    LogDiagnosticsMessage PLCTag & CStr(selectedgroups) & "," & CStr(selectedgroup) & "]" & " was set to " & m("PLC_ID") & " by " & Me.Application.CurrentUserName & " on " & footer.ComputerName, ftDiagSeverityAudit

    Set Tag_Group = Application.CreateTagGroup(Me.AreaName, 250)
    Tag_Group.Add PLCTag & CStr(selectedgroups) & "," & CStr(selectedgroup) & "]"
    Tag_Group.Item(1).Value = m("PLC_ID")

    'Close tag group
    Tag_Group.RemoveAll
    Set Tag_Group = Nothing
    
    ShowGroup 0
    
Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error cmdAssign_Released" & Err.Number & " " & Err.Description & " on display " & Name, ftDiagSeverityError
End Sub

Private Sub ShowGroups()
    'populate list of areas
    Dim i As Integer
    Dim j As Integer
    cmdRename.Enabled = False
    cmdAssign.Enabled = False
    cmdDelete.Enabled = False
    selectedgroup = 0
    For i = 1 To num
        j = i + igroups * num
        With Me.Elements("GroupsButton" & CStr(i))
            If j <= cgroups.Count Then
                .Caption = cgroups(j)
                .Visible = True
            Else
                .Caption = ""
                .Visible = False
            End If
        End With
    Next i
     If igroups < groupspagecount - 1 Then
        GroupsDownButton.Enabled = True
    Else
        GroupsDownButton.Enabled = False
    End If
    If igroups > 0 Then
        GroupsUpButton.Enabled = True
    Else
        GroupsUpButton.Enabled = False
    End If
End Sub

Private Sub ShowMaterialClass()
    'populate list of areas
    Dim i As Integer
    Dim j As Integer
    Dim m As Collection
    cmdAssign.Enabled = False
    cmdDelete.Enabled = False
    For i = 1 To num
        j = i + iMatClass * num
        With Me.Elements("MatClassButton" & CStr(i))
            If j <= cMatClass.Count Then
                Set m = cMatClass.Item(j)
                .Caption = m("Name")
                .Visible = True
            Else
                .Caption = ""
                .Visible = False
            End If
        End With
    Next i
    If iMatClass < MatClasspagecount - 1 Then
        MatClassDownButton.Enabled = True
    Else
        MatClassDownButton.Enabled = False
    End If
    If iMatClass > 0 Then
        MatClassUpButton.Enabled = True
    Else
        MatClassUpButton.Enabled = False
    End If
    iMat = 0
    ShowMaterial 1
End Sub

Private Sub ShowGroup(Index As Integer)
    'populate list of units in the area
    Dim j As Integer
    Dim i As Integer
    Dim k As Integer
    cmdAssign.Enabled = False
    cmdDelete.Enabled = False
    'Error check
    On Error GoTo ErrHandler
    SelectGroup 0
    If Index > 0 Then
        selectedgroups = Index + (igroups * num) - 1
    End If
    'read material tags
    Set Tag_Group = Application.CreateTagGroup(Me.AreaName, 250)
    For i = 0 To 9
        k = i + 10 * igroup
        Tag_Group.Add PLCTag & CStr(selectedgroups) & "," & CStr(k) & "]"
    Next
    Tag_Group.Active = True
    If Not Tag_Group.RefreshFromSource(TagsInError) Then
        LogDiagnosticsMessage "Taggroup refresh failed on display " & Name
    End If

    For i = 0 To 9
        With Me.Elements("GroupButton" & CStr(i + 1))
            .Caption = Tag_Group.Item(i + 1).Value & " - " & footer.MaterialName_FromTPMDB(Tag_Group.Item(i + 1).Value)
            .Visible = True
        End With
    Next
    'Close tag group
    Tag_Group.Active = False
    Tag_Group.RemoveAll
    Set Tag_Group = Nothing

    If Index > 0 Then
        jgroups = Index
    End If

    'update color on area button that matches
    For i = 1 To num
        If i = jgroups Then
            Me.Elements("GroupsButton" & CStr(i)).BackColor = &HC00000
        Else
            Me.Elements("GroupsButton" & CStr(i)).BackColor = &HA4A0A0
        End If
    Next i

    If igroup < grouppagecount - 1 Then
        GroupDownButton.Enabled = True
    Else
        GroupDownButton.Enabled = False
    End If
    If igroup > 0 Then
        GroupUpButton.Enabled = True
    Else
        GroupUpButton.Enabled = False
    End If
    cmdRename.Enabled = True
Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " ShowGroup", ftDiagSeverityError

End Sub

Private Sub ShowMaterial(Index As Integer)
     'populate list of units in the area
    Dim j As Integer
    Dim i As Integer
    Dim k As Integer
    Dim m As Collection

    'Error check
    On Error GoTo ErrHandler
    SelectMaterial 0
    cmdAssign.Enabled = False
    cmdDelete.Enabled = False
    If Index > 0 Then
        selectedMatClass = CStr(Index + (iMatClass * num))
    End If
      
'
    If Index > 0 Then
        jMatClass = Index
    End If

    'update color on area button that matches
    For i = 1 To num
        If i = jMatClass Then
            Me.Elements("MatClassButton" & CStr(i)).BackColor = &HC00000
        Else
            Me.Elements("MatClassButton" & CStr(i)).BackColor = &HA4A0A0
        End If
    Next i
    
    'keep area if using up\down button
    If jMatClass > 0 Then
        Set selectedMat = footer.GetMaterialList(Me.Elements("MatClassButton" & CStr(jMatClass)).Caption)
        If selectedMat.Count Mod num = 0 Then
            Matpagecount = Int(selectedMat.Count / num)
        Else
            Matpagecount = Int(selectedMat.Count / num) + 1
        End If
    End If

    For i = 1 To num
        j = i + iMat * num
        With Me.Elements("MaterialButton" & CStr(i))
            If j <= selectedMat.Count Then
                Set m = selectedMat.Item(j)
                .Caption = m("PLC_ID") & " - " & m("Name")
                .Visible = True
            Else
                .Caption = ""
                .Visible = False
            End If
        End With
    Next i
    
    If iMat < Matpagecount - 1 Then
        MaterialDownButton.Enabled = True
    Else
        MaterialDownButton.Enabled = False
    End If
    If iMat > 0 Then
        MaterialUpButton.Enabled = True
    Else
        MaterialUpButton.Enabled = False
    End If
Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " ShowMaterial", ftDiagSeverityError
    
End Sub

Private Sub SelectMaterial(Index As Integer)
    Dim i As Integer
    
    If Index > 0 Then
        selectedMatIndex = Index + (iMat * num)
    End If
    
    'update color on area button that matches
    For i = 1 To num
        If i = Index Then
            Me.Elements("MaterialButton" & CStr(i)).BackColor = &HC00000
        Else
            Me.Elements("MaterialButton" & CStr(i)).BackColor = &HA4A0A0
        End If
    Next i
    
    If selectedgroup > 0 Then
        If selectedMatIndex > 0 Then
            cmdAssign.Enabled = CurrentUserHasCode("O")
        End If
        cmdDelete.Enabled = CurrentUserHasCode("O")
    End If
End Sub

Private Sub SelectGroup(Index As Integer)
    Dim i As Integer

    selectedgroup = Index + (igroup * num) - 1
    'update color on area button that matches
    For i = 1 To num
        If i = Index Then
            Me.Elements("GroupButton" & CStr(i)).BackColor = &HC00000
        Else
            Me.Elements("GroupButton" & CStr(i)).BackColor = &HA4A0A0
        End If
    Next i

    If selectedgroup >= 0 Then
        If selectedMatIndex > 0 Then
            cmdAssign.Enabled = CurrentUserHasCode("O")
        End If
        cmdDelete.Enabled = CurrentUserHasCode("O")
    End If
End Sub

Private Sub Display_AfterAnimationStop()

    'Error check
    On Error GoTo ErrHandler
         
    'Clear
    Set footer = Nothing
    
Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error " & Err.Number & " " & Err.Description & " on display " & Name & " Display_AfterAnimationStop", ftDiagSeverityError
End Sub


Private Sub cmdRename_Released()
     
    'Error check
    On Error GoTo ErrHandler
    
    Dim frm As New frmNewname
    Dim i As Integer
    Dim s As String
    Dim c As String
    Dim d As Long
    
    For i = 1 To num
        If i = jgroups Then
            c = Me.Elements("GroupsButton" & CStr(i)).Caption
            d = InStr(1, c, "-")
            s = Right(c, Len(c) - d - 1)
            frm.NameTB.Text = s
        End If
    Next i

    frm.NameTB.SelStart = 0
    frm.NameTB.SelLength = frm.NameTB.TextLength
    frm.Show
    
    If frm.Result = vbOK Then
        footer.MaterialGroupUpdate selectedgroups, frm.NameTB.Text
        footer.GetMaterialGroupCollection cgroups
        ShowGroups
    End If
    Set frm = Nothing
    
Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error cmdRename_Released " & Err.Number & " " & Err.Description & " on display " & Name, ftDiagSeverityError
End Sub
Private Sub cmdDelete_Released()

    'Error check
    On Error GoTo ErrHandler

    LogDiagnosticsMessage PLCTag & CStr(selectedgroups) & "," & CStr(selectedgroup) & "]" & " was set to 0 by " & Me.Application.CurrentUserName & " on " & footer.ComputerName, ftDiagSeverityAudit

    Set Tag_Group = Application.CreateTagGroup(Me.AreaName, 250)
    Tag_Group.Add PLCTag & CStr(selectedgroups) & "," & CStr(selectedgroup) & "]"
    Tag_Group.Item(1).Value = 0

    'Close tag group
    Tag_Group.RemoveAll
    Set Tag_Group = Nothing
    
    ShowGroup 0
    
Exit Sub

'Error message
ErrHandler:
    LogDiagnosticsMessage "VBA error cmdDelete_Released" & Err.Number & " " & Err.Description & " on display " & Name, ftDiagSeverityError
End Sub
Private Sub GroupsDownButton_Released()
    If igroups >= groupspagecount - 1 Then Exit Sub
    igroups = igroups + 1
    ShowGroups
End Sub
Private Sub GroupsUpButton_Released()
    If igroups <= 0 Then Exit Sub
    igroups = igroups - 1
    ShowGroups
End Sub
Private Sub GroupsButton10_Released()
    ShowGroup 10
End Sub
Private Sub GroupsButton9_Released()
    ShowGroup 9
End Sub
Private Sub GroupsButton8_Released()
    ShowGroup 8
End Sub
Private Sub GroupsButton7_Released()
    ShowGroup 7
End Sub
Private Sub GroupsButton6_Released()
    ShowGroup 6
End Sub
Private Sub GroupsButton5_Released()
    ShowGroup 5
End Sub
Private Sub GroupsButton4_Released()
    ShowGroup 4
End Sub
Private Sub GroupsButton3_Released()
    ShowGroup 3
End Sub
Private Sub GroupsButton2_Released()
    ShowGroup 2
End Sub
Private Sub GroupsButton1_Released()
    ShowGroup 1
End Sub
Private Sub GroupDownButton_Released()
    If igroup >= grouppagecount - 1 Then Exit Sub
    igroup = igroup + 1
    ShowGroup -1
End Sub
Private Sub GroupUpButton_Released()
    If igroup <= 0 Then Exit Sub
    igroup = igroup - 1
    ShowGroup -1
End Sub
Private Sub GroupButton10_Released()
    SelectGroup 10
End Sub
Private Sub GroupButton9_Released()
    SelectGroup 9
End Sub
Private Sub GroupButton8_Released()
    SelectGroup 8
End Sub
Private Sub GroupButton7_Released()
    SelectGroup 7
End Sub
Private Sub GroupButton6_Released()
    SelectGroup 6
End Sub
Private Sub GroupButton5_Released()
    SelectGroup 5
End Sub
Private Sub GroupButton4_Released()
    SelectGroup 4
End Sub
Private Sub GroupButton3_Released()
    SelectGroup 3
End Sub
Private Sub GroupButton2_Released()
    SelectGroup 2
End Sub
Private Sub GroupButton1_Released()
    SelectGroup 1
End Sub
Private Sub MatClassDownButton_Released()
    If iMatClass >= MatClasspagecount - 1 Then Exit Sub
    iMatClass = iMatClass + 1
    ShowMaterialClass
End Sub
Private Sub MatClassUpButton_Released()
    If iMatClass <= 0 Then Exit Sub
    iMatClass = iMatClass - 1
    ShowMaterialClass
End Sub
Private Sub MatClassButton10_Released()
    ShowMaterial 10
End Sub
Private Sub MatClassButton9_Released()
    ShowMaterial 9
End Sub
Private Sub MatClassButton8_Released()
    ShowMaterial 8
End Sub
Private Sub MatClassButton7_Released()
    ShowMaterial 7
End Sub
Private Sub MatClassButton6_Released()
    ShowMaterial 6
End Sub
Private Sub MatClassButton5_Released()
    ShowMaterial 5
End Sub
Private Sub MatClassButton4_Released()
    ShowMaterial 4
End Sub
Private Sub MatClassButton3_Released()
    ShowMaterial 3
End Sub
Private Sub MatClassButton2_Released()
    ShowMaterial 2
End Sub
Private Sub MatClassButton1_Released()
    ShowMaterial 1
End Sub
Private Sub MaterialDownButton_Released()
    If iMat >= Matpagecount - 1 Then Exit Sub
    iMat = iMat + 1
    ShowMaterial -1
End Sub
Private Sub MaterialUpButton_Released()
    If iMat <= 0 Then Exit Sub
    iMat = iMat - 1
    ShowMaterial -1
End Sub
Private Sub MaterialButton10_Released()
    SelectMaterial 10
End Sub
Private Sub MaterialButton9_Released()
    SelectMaterial 9
End Sub
Private Sub MaterialButton8_Released()
    SelectMaterial 8
End Sub
Private Sub MaterialButton7_Released()
    SelectMaterial 7
End Sub
Private Sub MaterialButton6_Released()
    SelectMaterial 6
End Sub
Private Sub MaterialButton5_Released()
    SelectMaterial 5
End Sub
Private Sub MaterialButton4_Released()
    SelectMaterial 4
End Sub
Private Sub MaterialButton3_Released()
    SelectMaterial 3
End Sub
Private Sub MaterialButton2_Released()
    SelectMaterial 2
End Sub
Private Sub MaterialButton1_Released()
    SelectMaterial 1
End Sub
