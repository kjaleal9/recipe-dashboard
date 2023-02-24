

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

    
    If lblInvalid.Caption <> "" Then
        MsgBox "Invalid File"
        Exit Sub
        
    End If
    
    Dim xdoc As New MSXML2.DOMDocument60
    xdoc.Load (TxtPath.value)
    
    Dim rs As adodb.Recordset
    Dim par As adodb.Parameter
    Set par = New adodb.Parameter
    par.Type = adLongVarChar                      'Parameter type
    par.Direction = adParamInput              'Parameter direction (IN, OUT, IN/OUT,)
    par.value = xdoc.xml                 'Parameter value (MaterialClass)
    par.Size = -1                            'Max. size for Parameter value
    db.executeCommand1Par "TPIBK_ImportRecipeXML", par

    Set xdoc = Nothing
    MsgBox "Import Complete"
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

Private Sub TxtPath_Change()
    Dim fso As New FileSystemObject
    
    If Not fso.FileExists(TxtPath.value) Then
        lblInvalid.Caption = "Invalid File!"
    ElseIf Not TxtPath.value Like "*.xml" Then
        lblInvalid.Caption = ""
    Else
        lblInvalid.Caption = ""
    End If
    
    Set fso = Nothing
End Sub
