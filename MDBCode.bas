Attribute VB_Name = "MDBCode"
Option Explicit
'Dim WS As Workspace
Dim RS As adodb.Recordset
Dim Cnn As New adodb.Connection

Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, _
ByVal nIDEvent As Long, ByVal uElapse As Long, _
ByVal lpTimerProc As Long) As Long

Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, _
ByVal nIDEvent As Long) As Long

Private TimerID As Long

Const NOCUSTOMER As Long = 6
Const NOCONTACT As Long = 10
Const NOPRODUCT As Long = 14
Const NOCODE As Long = 6
Function GetData(sQry As String, EmplID As Long) As Collection
'Get a db connection

End Function
Function GetListRS(sSql As String, oData As Object)
'=========================================================
'GetListRS passes a sql string to retrieve listbox data
'and fills the class collection
'sSql is sql SELECT query
'oData is collection class for each type of list data
'=========================================================

Dim ErrorCondition As Integer
On Error GoTo ErrorHandler    ' Enable error trapping.

If Not (Cnn Is Nothing) Then
    ConnectDB
End If

Set RS = New adodb.Recordset
RS.Open sSql, Cnn, , , adCmdUnknown

If Not (oData Is CLinks) Then
    Do Until RS.EOF
        oData.Add _
        ID:=RS!ID, _
        sName:=RS!sName
        
        RS.MoveNext
    Loop
Else 'is a link and needs more info
    Do Until RS.EOF
        oData.Add _
        ID:=RS!ID, _
        CompanyID:=RS!CompanyID, _
        ContactId:=RS!ContactId
        
        RS.MoveNext
    Loop
End If

Set RS = Nothing

On Error GoTo 0 ' Disable error trapping.
GetListRS = 1
Exit Function

ErrorHandler:
    ErrorCondition = True
    MsgBox Err.Number & vbCrLf & Err.Description & sSql, vbExclamation, "GetListRS"
    Resume Next
End Function
'Function GetLinkRecordset(sSql As String, colData As Object) As Integer
''=========================================================
''Function should be obsolete and will be deleted upon release
''=========================================================
'Dim ErrorCondition As Integer
'
'On Error GoTo DBErrorHandler    ' Enable error trapping.
'If Not ErrorCondition Then
'    On Error GoTo TableErrorHandler ' Enable error trapping.
'    ConnectDB
'    Set RS = New adodb.Recordset
'    RS.Open sSql, Cnn, , , adCmdUnknown
'    RS.MoveFirst
'End If
'
'Do Until RS.EOF
'    colData.Add ID:=RS.Fields("ID").Value, _
'    CompanyID:=RS.Fields("CompanyID").Value, _
'    ContactID:=RS.Fields("ContactID").Value
'    RS.MoveNext
'Loop
'
'Set RS = Nothing
'CloseDB
'
'On Error GoTo 0 ' Disable error trapping.
'GetLinkRecordset = 1
'Exit Function
'
'DBErrorHandler:
'TableErrorHandler:
'    ErrorCondition = True
'    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation
'    Resume Next
'
'End Function
'Function AppendCall(sTable As String, iCustomerID As Long, iContactID As Long, iCallCodeID As Long, iProductID As Long, iEmployeeID As Long, dNoteDate As Date, sNote As String, dEntryDate As Date, iCallTime As Integer) As Integer
''=========================================================
''Function should be obsolete and will be deleted upon release
''=========================================================
'On Error GoTo ErrorHandler    ' Enable error trapping.
'AppendCall = 0
'
'If Not (Cnn Is Nothing) Then
'    ConnectDB
'End If
'
'Set RS = New adodb.Recordset
'
'With RS
'    .CursorType = adOpenKeyset
'    .LockType = adLockOptimistic
'    .Open sTable, Cnn, , , adCmdUnknown
'
'    .AddNew
'        RS!CustomerID = iCustomerID
'        RS!ContactID = iContactID
'        RS!CallCodeId = iCallCodeID
'        RS!productid = iProductID
'        RS!EmployeeID = iEmployeeID
'        RS!NoteDate = dNoteDate
'        RS!note = sNote
'        RS!EntryDate = dEntryDate
'        RS!CallTime = iCallTime
'    .Update
'End With 'RS
'Set RS = Nothing
'
'On Error GoTo 0 ' Disable error trapping.
'AppendCall = 1
'Exit Function
'
'ErrorHandler:
'    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "DBModule - AppendCall"
'    Resume Next
'
'End Function
Function GetHistoryRS(sQry As String, oCalls As Object) As Integer
Dim CallTime As Integer

On Error GoTo ErrorHandler    ' Enable error trapping.
GetHistoryRS = 0 'Assume failure

If Not (Cnn Is Nothing) Then
    ConnectDB
End If

Set RS = New adodb.Recordset

With RS
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open sQry, Cnn, adOpenForwardOnly, , adCmdUnknown
    
    Do Until .EOF
        If RS!iCallTime > 0 Then
            CallTime = RS!iCallTime
        Else
            CallTime = 0
        End If
        
        oCalls.Add _
            sLastName:=RS!LastName, _
            sContactName:=RS!ContactName, _
            sCompanyName:=RS!CompanyName, _
            sProductName:=RS!ProductName, _
            sCallType:=RS!CallType, _
            dNoteDate:=RS!NoteDate, _
            sNote:=RS!sNote, _
            iCallTime:=CallTime
        .MoveNext
    Loop
End With 'RS

Set RS = Nothing

On Error GoTo 0 ' Disable error trapping.
GetHistoryRS = 1
Exit Function

ErrorHandler:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation
    Resume Next
End Function
Function AppendData(sTable As String, oData As Object) As Integer
'================================================================
'This function appends NEW records to the Call Log database.  The
'new records can be calls or new customers/contacts/call codes.
'
'First it checks for an existing db connection then
'Creates a opens the specified db table (sTable)
'Appends record depending or the sTable selection
'
'Return is the new ID for the record; 0 is a failure.
'================================================================
Dim ErrorCondition As Integer
Dim vCounter As Variant
Dim dEntryDate As Date

On Error GoTo ErrorHandler    ' Enable error trapping.
AppendData = 0   'Assume failure

If Not (Cnn Is Nothing) Then
    ConnectDB
End If

Set RS = New adodb.Recordset

RS.CursorType = adOpenKeyset
RS.LockType = adLockOptimistic
RS.Open sTable, Cnn, , , adCmdUnknown

dEntryDate = Now()

For Each vCounter In oData
    RS.AddNew
    With vCounter
        Select Case sTable
            Case "SupportCalls"
                RS!NoteDate = .NoteDate
                RS!CallTime = .CallTime
            'Required for DB integrety, all for time only
                If .CustomerID = 0 Then
                    RS!CustomerID = NOCUSTOMER
                Else
                    RS!CustomerID = .CustomerID
                End If
                If .ContactId = 0 Then
                    RS!ContactId = NOCONTACT
                Else
                    RS!ContactId = .ContactId
                End If
                If .productid = 0 Then
                    RS!productid = NOPRODUCT
                Else
                    RS!productid = .productid
                End If
                If .CallCodeId = 0 Then
                    RS!CallCodeId = NOCODE
                Else
                    RS!CallCodeId = .CallCodeId
                End If
                If .note = "" Then
                    RS!note = ""
                Else
                    RS!note = .note
                End If
            Case "Company"
                RS!CompanyName = .CompanyName
                RS!Address = .Address
                RS!City = .City
                RS!State = .State
                RS!Zip = .Zip
                RS!Phone = .Phone
                RS!Country = .Country
                RS!Fax = .Fax
            Case "Contact"
                RS!FirstName = .FirstName
                RS!LastName = .LastName
                RS!Title = .Title
            Case "CallCode"
                RS!CallType = .CallType
            Case "Contact Link"
                RS!CompanyID = .Company
                RS!ContactId = .Contact
        End Select
        
        RS!EmplID = .EmplID
        RS!DateEntered = dEntryDate
    
    End With 'vCounter
    RS.Update

AppendData = RS!ID
Next vCounter

Set RS = Nothing
On Error GoTo 0 ' Disable error trapping.
Exit Function

ErrorHandler:
    ErrorCondition = True
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "DBModule - AppendCall"
    Resume Next

End Function
Sub ConnectDB()
'================================================================
'Need to add code to allow the user or other object to set the location
'of the database.
'================================================================
Dim sConn As String
On Error GoTo ErrorHandler

' Open database.
Set Cnn = New adodb.Connection
Cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\Call Log\CALL LOADING.MDB;Persist Security Info=False"
sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\Call Log\CALL LOADING.MDB;Persist Security Info=False"
Cnn.Open sConn
Exit Sub

ErrorHandler:
sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\edi-clt-001\prod_sup\CALL LOADING.MDB;Persist Security Info=False"
Cnn.Open sConn
Resume Next

MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, " Error: ConnectDB"
End Sub
Sub CloseDB()
Cnn.Close
Set Cnn = Nothing
End Sub
Sub CheckIdleConnection()
TimerID = SetTimer(0, 0, 50, AddressOf TimerProc)
If TimerID = 0 Then
    MsgBox "No timers available!"
    CloseDB
End If
End Sub
Sub StartConnectionTimer()
TimerID = SetTimer(0, 0, 50, AddressOf TimerProc)
If TimerID = 0 Then
    MsgBox "No timers available!"
End If
End Sub
Public Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal lngSysTimer As Long)
KillTimer 0, TimerID
CheckIdleConnection
End Sub
