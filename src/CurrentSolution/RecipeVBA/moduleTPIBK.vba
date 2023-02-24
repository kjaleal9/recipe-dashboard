Private loginstatus As Boolean
Public Const RecipeAccess As Long = 134217728
Private rarray() As String  'record array
Private rarrayt() As String 'record array transpose

Public Property Get FieldStringArray() As String
Dim fieldlist As String
    fieldlist = "TPIBK_StepType_ID," + _
                "Equipment_ID," + _
                "PhaseType_ID," + _
                "Step," + _
                "UserString," + _
                "Transition_Index," + _
                "NextStep," + _
                "Allocation_Type_ID," + _
                "LateBinding," + _
                "Material_ID," + _
                "MaterialClass_ID," + _
                "Material_Name," + _
                "Equipment_Name," + _
                "Phase_Name," + _
                "MainBatchUnit," + _
                "Material_Amount," + _
                "Phase_Mode," + _
                "Equipment_Mode," + _
                "Speed,"
    fieldlist = fieldlist & "Level," + _
                "Temperature," + _
                "Flow," + _
                "Seconds," + _
                "Pressure," + _
                "Enable_Calculation," + _
                "User_REAL[0]," + _
                "User_REAL[1]," + _
                "User_REAL[2]," + _
                "User_REAL[3]," + _
                "User_REAL[4]," + _
                "User_REAL[5]," + _
                "User_REAL[6]," + _
                "User_REAL[7]," + _
                "User_REAL[8]," + _
                "User_REAL[9]," + _
                "User_DINT[0]," + _
                "User_DINT[1]," + _
                "User_DINT[2]," + _
                "User_DINT[3]," + _
                "User_DINT[4],"
    fieldlist = fieldlist & "User_DINT[5]," + _
                "User_DINT[6]," + _
                "User_DINT[7]," + _
                "User_DINT[8]," + _
                "User_DINT[9]," + _
                "User_BOOL[0]," + _
                "User_BOOL[1]," + _
                "User_BOOL[2]," + _
                "User_BOOL[3]," + _
                "User_BOOL[4]," + _
                "User_BOOL[5]," + _
                "User_BOOL[6]," + _
                "User_BOOL[7]," + _
                "User_BOOL[8]," + _
                "User_BOOL[9]," + _
                "Material_Amount_EU," + _
                "Phase_Mode_EU," + _
                "Equipment_Mode_EU," + _
                "Speed_EU,"
    fieldlist = fieldlist & "Level_EU," + _
                "Temperature_EU," + _
                "Flow_EU," + _
                "Seconds_EU," + _
                "Pressure_EU," + _
                "User_Real_EU[0]," + _
                "User_Real_EU[1]," + _
                "User_Real_EU[2]," + _
                "User_Real_EU[3]," + _
                "User_Real_EU[4]," + _
                "User_Real_EU[5]," + _
                "User_Real_EU[6]," + _
                "User_Real_EU[7]," + _
                "User_Real_EU[8]," + _
                "User_Real_EU[9]," + _
                "User_DINT_EU[0]," + _
                "User_DINT_EU[1]," + _
                "User_DINT_EU[2]," + _
                "User_DINT_EU[3]," + _
                "User_DINT_EU[4]," + _
                "User_DINT_EU[5],"
    fieldlist = fieldlist & "User_DINT_EU[6]," + _
                "User_DINT_EU[7]," + _
                "User_DINT_EU[8]," + _
                "User_DINT_EU[9]," + _
                "Material_Amount_Min," + _
                "Phase_Mode_Min," + _
                "Equipment_Mode_Min," + _
                "Speed_Min," + _
                "Level_Min,"
    fieldlist = fieldlist & "Temperature_Min," + _
                "Flow_Min," + _
                "Seconds_Min," + _
                "Pressure_Min," + _
                "User_Real_Min[0]," + _
                "User_Real_Min[1]," + _
                "User_Real_Min[2]," + _
                "User_Real_Min[3]," + _
                "User_Real_Min[4]," + _
                "User_Real_Min[5]," + _
                "User_Real_Min[6]," + _
                "User_Real_Min[7]," + _
                "User_Real_Min[8]," + _
                "User_Real_Min[9]," + _
                "User_DINT_Min[0]," + _
                "User_DINT_Min[1]," + _
                "User_DINT_Min[2]," + _
                "User_DINT_Min[3]," + _
                "User_DINT_Min[4]," + _
                "User_DINT_Min[5]," + _
                "User_DINT_Min[6],"
    fieldlist = fieldlist & "User_DINT_Min[7]," + _
                "User_DINT_Min[8]," + _
                "User_DINT_Min[9]," + _
                "Material_Amount_Max," + _
                "Phase_Mode_Max," + _
                "Equipment_Mode_Max," + _
                "Speed_Max," + _
                "Level_Max," + _
                "Temperature_Max," + _
                "Flow_Max," + _
                "Seconds_Max," + _
                "Pressure_Max," + _
                "User_Real_Max[0]," + _
                "User_Real_Max[1]," + _
                "User_Real_Max[2]," + _
                "User_Real_Max[3]," + _
                "User_Real_Max[4]," + _
                "User_Real_Max[5]," + _
                "User_Real_Max[6]," + _
                "User_Real_Max[7],"
    fieldlist = fieldlist & "User_Real_Max[8]," + _
                "User_Real_Max[9]," + _
                "User_DINT_Max[0]," + _
                "User_DINT_Max[1]," + _
                "User_DINT_Max[2]," + _
                "User_DINT_Max[3]," + _
                "User_DINT_Max[4]," + _
                "User_DINT_Max[5]," + _
                "User_DINT_Max[6]," + _
                "User_DINT_Max[7]," + _
                "User_DINT_Max[8]," + _
                "User_DINT_Max[9]"
    FieldStringArray = fieldlist
End Property

Public Property Get FieldStringArrayDB() As String
Dim fieldlist As String
    fieldlist = "TPIBK_StepType_ID," + _
                "Equipment_ID," + _
                "PhaseType_ID," + _
                "Step," + _
                "UserString," + _
                "Transition_Index," + _
                "NextStep," + _
                "Allocation_Type_ID," + _
                "LateBinding," + _
                "Material_ID," + _
                "MaterialClass_ID," + _
                "Material_Name," + _
                "Equipment_Name," + _
                "Phase_Name," + _
                "MainBatchUnit," + _
                "Material_Amount," + _
                "Phase_Mode," + _
                "Equipment_Mode," + _
                "Speed,"
    fieldlist = fieldlist & "Level," + _
                "Temperature," + _
                "Flow," + _
                "Seconds," + _
                "Pressure," + _
                "Enable_Calculation," + _
                "User_REAL_1," + _
                "User_REAL_2," + _
                "User_REAL_3," + _
                "User_REAL_4," + _
                "User_REAL_5," + _
                "User_REAL_6," + _
                "User_REAL_7," + _
                "User_REAL_8," + _
                "User_REAL_9," + _
                "User_REAL_10," + _
                "User_DINT_1," + _
                "User_DINT_2," + _
                "User_DINT_3," + _
                "User_DINT_4," + _
                "User_DINT_5,"
    fieldlist = fieldlist & "User_DINT_6," + _
                "User_DINT_7," + _
                "User_DINT_8," + _
                "User_DINT_9," + _
                "User_DINT_10," + _
                "User_BOOL_1," + _
                "User_BOOL_2," + _
                "User_BOOL_3," + _
                "User_BOOL_4," + _
                "User_BOOL_5," + _
                "User_BOOL_6," + _
                "User_BOOL_7," + _
                "User_BOOL_8," + _
                "User_BOOL_9," + _
                "User_BOOL_10," + _
                "Material_Amount_EU," + _
                "Phase_Mode_EU," + _
                "Equipment_Mode_EU," + _
                "Speed_EU,"
    fieldlist = fieldlist & "Level_EU," + _
                "Temperature_EU," + _
                "Flow_EU," + _
                "Seconds_EU," + _
                "Pressure_EU," + _
                "User_Real_1_EU," + _
                "User_Real_2_EU," + _
                "User_Real_3_EU," + _
                "User_Real_4_EU," + _
                "User_Real_5_EU," + _
                "User_Real_6_EU," + _
                "User_Real_7_EU," + _
                "User_Real_8_EU," + _
                "User_Real_9_EU," + _
                "User_Real_10_EU," + _
                "User_DINT_1_EU," + _
                "User_DINT_2_EU," + _
                "User_DINT_3_EU," + _
                "User_DINT_4_EU," + _
                "User_DINT_5_EU," + _
                "User_DINT_6_EU,"
    fieldlist = fieldlist & "User_DINT_7_EU," + _
                "User_DINT_8_EU," + _
                "User_DINT_9_EU," + _
                "User_DINT_10_EU," + _
                "Material_Amount_Min," + _
                "Phase_Mode_Min," + _
                "Equipment_Mode_Min," + _
                "Speed_Min," + _
                "Level_Min,"
    fieldlist = fieldlist & "Temperature_Min," + _
                "Flow_Min," + _
                "Seconds_Min," + _
                "Pressure_Min," + _
                "User_Real_1_Min," + _
                "User_Real_2_Min," + _
                "User_Real_3_Min," + _
                "User_Real_4_Min," + _
                "User_Real_5_Min," + _
                "User_Real_6_Min," + _
                "User_Real_7_Min," + _
                "User_Real_8_Min," + _
                "User_Real_9_Min," + _
                "User_Real_10_Min," + _
                "User_DINT_1_Min," + _
                "User_DINT_2_Min," + _
                "User_DINT_3_Min," + _
                "User_DINT_4_Min," + _
                "User_DINT_5_Min," + _
                "User_DINT_6_Min," + _
                "User_DINT_7_Min,"
    fieldlist = fieldlist & "User_DINT_8_Min," + _
                "User_DINT_9_Min," + _
                "User_DINT_10_Min," + _
                "Material_Amount_Max," + _
                "Phase_Mode_Max," + _
                "Equipment_Mode_Max," + _
                "Speed_Max," + _
                "Level_Max," + _
                "Temperature_Max," + _
                "Flow_Max," + _
                "Seconds_Max," + _
                "Pressure_Max," + _
                "User_Real_1_Max," + _
                "User_Real_2_Max," + _
                "User_Real_3_Max," + _
                "User_Real_4_Max," + _
                "User_Real_5_Max," + _
                "User_Real_6_Max," + _
                "User_Real_7_Max," + _
                "User_Real_8_Max,"
    fieldlist = fieldlist & "User_Real_9_Max," + _
                "User_Real_10_Max," + _
                "User_DINT_1_Max," + _
                "User_DINT_2_Max," + _
                "User_DINT_3_Max," + _
                "User_DINT_4_Max," + _
                "User_DINT_5_Max," + _
                "User_DINT_6_Max," + _
                "User_DINT_7_Max," + _
                "User_DINT_8_Max," + _
                "User_DINT_9_Max," + _
                "User_DINT_10_Max"
    FieldStringArrayDB = fieldlist
End Property

Public Property Get ValidLogin() As Variant
    ValidLogin = loginstatus
End Property

Public Property Let ValidLogin(ByVal vNewValue As Variant)
    loginstatus = vNewValue
End Property

Public Function LoginErrorMessage(index As Integer) As String
    Select Case index
        Case 21
            LoginErrorMessage = "User not defined"
        Case 36
            LoginErrorMessage = "User not active"
        Case 22
            LoginErrorMessage = "Incorrect Password"
        Case 24
            LoginErrorMessage = "User not assigned to group"
            
    End Select
End Function

Function NullIf(value As Variant, NullValue As Variant) As Variant
    If value = NullValue Then
        NullIf = Null
    Else
        NullIf = value
    End If
End Function

Function IfNull(value As Variant, Optional NullValue As Variant = "") As Variant
    If IsNull(value) Then
        IfNull = NullValue
    Else
        IfNull = value
    End If
End Function

Public Sub SelectRow(Cmb As MSForms.Control, Col As Integer, valuetofind As Long, ByRef mCmb As String)
On Error GoTo ErrHandlerSQL
    If Cmb.ListCount = 0 Then Exit Sub
    mCmb = Cmb.Column(0, Cmb.ListIndex)
    For i = 0 To Cmb.ListCount - 1
'        Debug.Print i
'        Debug.Print Cmb.Column(col, i)
        If Cmb.Column(Col, i) = valuetofind Then
            mCmb = Cmb.Column(0, i)
            Cmb.ListIndex = i
            Exit Sub
        End If
    Next i
    Cmb.ListIndex = 0
Exit Sub

'Errorhandling
ErrHandlerSQL:
    MsgBox "VBA error " & Err.Number & " " & Err.Description & " on form " & Name
    Exit Sub

End Sub

'*******************************************************************
'***** THE FOLLOWING WAS COPIED FROM THE PFS MODULE IN THE XLS *****
'*******************************************************************

'this function takes the base string and goes throught an array of strings and replaces
'the placeholders i.e. %%1 whith the string in the array.
'goes in reverse order intentionally
Public Function Placeholder(basestring As String, replacements() As String) As String
    For i = 10 To 1 Step -1
        basestring = Replace(basestring, "%%" & CStr(i), replacements(i))
    Next i
    Placeholder = basestring
End Function

Public Sub ListRefreshSQL(rs As adodb.Recordset, db As dbConnect, sql As String, Cmb As Object, fieldlist As String, valuetofind As Variant)
'On Error GoTo ErrHandlerSQL
    i = 0
    j = 0
    k = 0
    fieldarray = Split(fieldlist, ",")

    Set rs = db.getRecords(sql)
    'If Cmb.Name = "Lstparameters" Then
    '    MsgBox sql
    'End If

    ReDim rarrayt(Cmb.ColumnCount, 1)
    If rs Is Nothing Then Exit Sub
    While Not rs.EOF
        For k = 0 To Cmb.ColumnCount - 1
            If IsNull(rs.Fields(fieldarray(k)).value) Then
                rarrayt(k, i) = ""
            Else
                rarrayt(k, i) = CStr(rs.Fields(fieldarray(k)).value)
            End If
        Next k
        If valuetofind = rarrayt(0, i) Then 'look for a value in field and mark it
            j = i
        End If
        rs.MoveNext
        i = i + 1
        ReDim Preserve rarrayt(Cmb.ColumnCount, i)
    Wend
    Cmb.List = Transpose(i, Cmb.ColumnCount, rarrayt)
    Cmb.ListIndex = j
    valuetofind = 0
    'rs.Close
    Set rs = Nothing
Exit Sub

'Errorhandling
ErrHandlerSQL:
    MsgBox Now() & " VBA error SQL " & Err.Number & " " & Err.Description & " on form " & Name & " for ListRefreshSQL()"
End Sub

Public Function Transpose(i, count, rarrayt() As String) As String()
    Dim x, z As Integer
    Dim rarray() As String
    
    ReDim rarray(i, count)
    
    For x = 0 To i
        For z = 0 To count
            rarray(x, z) = rarrayt(z, x)
        Next z
    Next x
    Transpose = rarray
End Function

Public Sub CheckNumericInput(ByVal KeyAscii As MSForms.ReturnInteger, tb As MSForms.TextBox)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Asc("-")
            MsgBox "Check Negative"
            If InStr(1, tb.Text, "-") > 0 Or tb.SelStart > 0 Then
                KeyAscii = 0
            End If
        Case Asc(".")
            If InStr(1, tb.Text, ".") > 0 Then
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0
    End Select
End Sub

Public Sub CheckSpecialInput(ByVal KeyAscii As MSForms.ReturnInteger, tb As MSForms.TextBox)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
            If tb.SelStart = 0 Then
                KeyAscii = 0
            End If
        Case Asc(" ")
            KeyAscii = Asc("_")
            If tb.SelStart = 0 Then
                KeyAscii = 0
            End If
        Case Asc("A") To Asc("Z")
        Case Asc("a") To Asc("z")
        Case Asc("_")
            If tb.SelStart = 0 Then
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0
            Beep
    End Select
End Sub

Function IfEmpty(value As Variant, Optional NullValue As Variant = "") As Variant
    If IsNull(value) Or value = "" Then
        IfEmpty = NullValue
    Else
        IfEmpty = value
    End If
End Function

Function GetValue(Cmb As MSForms.Control, Col, Optional Rtn = "")
    If Cmb.Enabled And Cmb.ListIndex <> -1 Then
        GetValue = IfEmpty(Cmb.Column(Col, Cmb.ListIndex), Rtn)
    Else
        GetValue = Rtn
    End If
End Function
