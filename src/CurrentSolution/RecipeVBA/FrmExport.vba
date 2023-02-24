
Option Explicit
'Option Compare Text

Private rs As adodb.Recordset
Private db As dbConnect
Private valuetofind As Long
Private UserCodeValue As String
Private ld As LocalDatabase
Private BID As String
Private Bver As String


Private Sub CmdExport_Click()
    
    Dim vbDoubleQuote As String
    vbDoubleQuote = """"
    Dim fso As New FileSystemObject
    
    If Not fso.FolderExists(TxtPath.value) Then
        MsgBox "Invalid Path"
        Exit Sub
        
    End If
    
    Dim rs As adodb.Recordset
    Dim par As adodb.Parameter
    Set par = New adodb.Parameter
    par.Type = adVarChar                      'Parameter type
    par.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par.value = BID                 'Parameter value (MaterialClass)
    par.Size = 50                             'Max. size for Parameter value
    Dim par1 As adodb.Parameter
    Set par1 = New adodb.Parameter
    par1.Type = adVarChar                      'Parameter type
    par.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par1.value = BatchVersion                    'Parameter value (MaterialClass)
    par1.Size = 50                             'Max. size for Parameter value
    
    Set rs = db.executeCommand2Par("TPIBK_getRecipeXML", par, par1)
    If Not rs.EOF Then

        If Right(TxtPath.value, 1) <> "\" Then
            TxtPath.value = TxtPath.value & "\"
        End If
        
        
        Dim xdoc As New MSXML2.DOMDocument60
        xdoc.resolveExternals = True

        xdoc.loadXML (rs(0).value)
        If xdoc.parseError <> 0 Then
            Dim strErrText As String
            strErrText = "Your XML Document failed to load" & _
            "due the following error." & vbCrLf & _
            "Error #: " & xdoc.parseError.errorCode & ": " & xdoc.parseError.reason & _
            "Line #: " & xdoc.parseError.Line & vbCrLf & _
            "Line Position: " & xdoc.parseError.linepos & vbCrLf & _
            "Position In File: " & xdoc.parseError.filepos & vbCrLf & _
            "Source Text: " & xdoc.parseError.srcText & vbCrLf & _
            "Document URL: " & xdoc.parseError.url
            MsgBox strErrText
        End If

        xdoc.Save (TxtPath.value & BID & "_" & Bver & ".xml")
        Set xdoc = Nothing
    End If
    rs.Close
    Set rs = Nothing
    Set fso = Nothing
    MsgBox "Export Complete"
    Unload Me
End Sub

Public Sub Init()
On Error GoTo ErrHandler 'Error Check

    Set db = ld.db
    db.ConnectDSN

Exit Sub

'Error message
ErrHandler:
   ' MsgBox "VBA error Animation Start " & Err.Number & " " & Err.Description & " on display " & name
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

Private Sub CmdPath_Click()
'    Dim objFileDialog As Office.FileDialog
'    Set objFileDialog = Application.FileDialog(MsoFileDialogType.msoFileDialogFolderPicker)
'
'    With objFileDialog
'        .AllowMultiSelect = True
'        .ButtonName = "Folder Picker"
'        .Title = "Folder Picker"
'        If (.Show > 0) Then
'        End If
'        If (.SelectedItems.count > 0) Then
'            Call MsgBox(.SelectedItems(1))
'        End If
'    End With
End Sub

Private Sub TxtPath_Change()
    Dim fso As New FileSystemObject
    
    If Not fso.FolderExists(TxtPath.value) Then
        lblInvalid.Caption = "Invalid Path!"
    Else
        lblInvalid.Caption = ""
    End If
    Set fso = Nothing
End Sub
