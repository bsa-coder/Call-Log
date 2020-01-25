VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Dim WS As Workspace
Dim RS As ADODB.Recordset
Dim Cnn As New ADODB.Connection

Const NOCUSTOMER As Long = 6
Const NOCONTACT As Long = 10
Const NOPRODUCT As Long = 14
Const NOCODE As Long = 6

Const clCOMPANY As Integer = 0
Const clCONTACT As Integer = 1
Const clPRODUCT As Integer = 2 ''
Const clCALLCODE As Integer = 3 ''
Const clCALL As Integer = 4
Const clLINK As Integer = 5
Const clHISTORY As Integer = 6
Const clCALLTYPE As Integer = 7
Const clEMPLOYEE As Integer = 8
Const clCOMPANYUPDATE As Integer = 9
Const clCONTACTUPDATE As Integer = 10
Function AppendData(sTable As String, rsData As ADODB.Recordset) As Long
'================================================================
'This function appends NEW records to the Call Tracker database.  The
'new records can be calls or new customers/contacts/call codes.
'
'First it checks for an existing db connection then
'Creates a opens the specified db table (sTable)
'Appends record depending or the sTable selection
'
'Return is the new ID for the record; 0 is a failure.
'================================================================
Dim ErrorCondition As Integer
Dim dEntryDate As Date
'Dim vCounter As Variant

On Error GoTo ErrorHandler    ' Enable error trapping.
AppendData = 0   'Assume failure

If Not (Cnn Is Nothing) Then
    ConnectDB
End If

Set RS = New ADODB.Recordset

RS.CursorType = adOpenKeyset
RS.LockType = adLockOptimistic
RS.Open sTable, Cnn, , , adCmdUnknown

dEntryDate = Now()

If ((sTable = "Company" Or sTable = "Contact") _
    And rsData!ID <> 0) Then 'Editing Record
    RS.Find "ID = " & rsData!ID
Else
    RS.AddNew
End If

RS!EmplID = rsData!EmplID
RS!DateEntered = dEntryDate

Select Case sTable
    Case "Company"
        RS!CompanyName = Trim(Left(rsData!sName, 50))
        RS!Address = Trim(rsData!Address)
        RS!City = Trim(rsData!City)
        RS!State = Trim(rsData!State)
        RS!Zip = Trim(rsData!Zip)
        RS!Phone = Trim(rsData!Phone)
        RS!Country = Trim(rsData!Country)
        RS!Fax = Trim(rsData!Fax)
        RS!Type = Trim(rsData!Type)
    Case "Contact"
        RS!FirstName = Trim(rsData!FirstName)
        RS!LastName = Trim(rsData!LastName)
        RS!Title = Trim(rsData!Title)
        RS!Phone = Trim(rsData!Phone)
        RS!Skill = Trim(rsData!Skill)
        RS!Training = Trim(rsData!Training)
    Case "[Contact Link]"
        RS!CompanyID = rsData!CompanyID
        RS!ContactID = rsData!ContactID
    Case "CallCode"
        RS!CallType = Trim(rsData!CallType)
        
    Case "SupportCalls"
        If rsData!CallCodeId = 0 Then
            RS!CallCodeId = NOCODE
        Else
            RS!CallCodeId = rsData!CallCodeId
        End If
        RS!CallTime = rsData!CallTime
    'Required for DB integraty, all for time only calls
        If rsData!CaseID = "" Then
            RS!CaseID = 0
        Else
            RS!CaseID = rsData!CaseID
        End If
        If rsData!CompanyID = 0 Then
            RS!CompanyID = NOCUSTOMER
        Else
            RS!CompanyID = rsData!CompanyID
        End If
        If rsData!ContactID = 0 Then
            RS!ContactID = NOCONTACT
        Else
            RS!ContactID = rsData!ContactID
        End If
        If rsData!Note = "" Then
            RS!Note = " "
        Else
            RS!Note = Trim(rsData!Note)
        End If
        RS!NoteDate = rsData!NoteDate
        If rsData!ProductId = 0 Then
            RS!ProductId = NOPRODUCT
        Else
            RS!ProductId = rsData!ProductId
        End If
End Select

RS.Update

AppendData = RS!ID
            
If sTable = "SupportCalls" Then
    If RS!CaseID = 0 Then
        RS!CaseID = RS!ID
        RS.Update
    End If
End If

Set RS = Nothing
On Error GoTo 0 ' Disable error trapping.
Exit Function

ErrorHandler:
    ErrorCondition = True
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "DB::AppendData - Error"
    Err.Raise 30000 + Err.Number, "DBServer-AppendData", Err.Description
End Function

Sub ConnectDB()
Dim sProvider As String
Dim sSource As String
Dim sSecurity As String
Dim sConn As String

On Error GoTo ErrorHandler

' Open database.
Set Cnn = New ADODB.Connection
sProvider = "Provider=Microsoft.Jet.OLEDB.4.0"
sSource = "Data Source=" & App.Path & "\CALL LOADING.MDB"
sSecurity = "Persist Security Info=False"
sConn = sProvider & ";" & sSource & ";" & sSecurity

Cnn.Open sConn
Exit Sub

ErrorHandler:
MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, " Error: ConnectDB"
Err.Raise 30000 + Err.Number, "DBServer-ConnectDB", Err.Description
Resume Next

End Sub
Sub CloseDB()
If (Cnn.State = adStateOpen) Then
    Cnn.Close
End If
Set Cnn = Nothing
End Sub
Public Function GetListRS(sQuery As String, ByRef rsData As ADODB.Recordset, Index As Integer)
'=========================================================
'GetListRS passes a sql string to retrieve listbox data
'and fills the class collection
'sQuery is sql SELECT query
'oData is collection class for each type of list data
'=========================================================

Dim ErrorCondition As Integer
On Error GoTo ErrorHandler    ' Enable error trapping.

If Not (Cnn Is Nothing) Then
    ConnectDB
End If

Set RS = New ADODB.Recordset
RS.Open sQuery, Cnn, , , adCmdUnknown
Set rsData = RS

Set RS = Nothing

On Error GoTo 0 ' Disable error trapping.

GetListRS = 1
Exit Function

ErrorHandler:
    ErrorCondition = True
    MsgBox Err.Number & vbCrLf & Err.Description & " " & sQuery, vbExclamation, "GetListRS"
    Err.Raise 30000 + Err.Number, "DBServer-GetListRS", Err.Description
    Resume Next
End Function
'Function AppendCall(sTable As String, iCustomerID As Long, iContactID As Long, iCallCodeID As Long, iProductID As Long, iEmployeeID As Long, dNoteDate As Date, sNote As String, dEntryDate As Date, iCallTime As Integer, iCaseID As Long) As Integer
'Dim ErrorCondition As Integer
'
'On Error GoTo ErrorHandler    ' Enable error trapping.
'AppendCall = 0
'
'If Not (Cnn Is Nothing) Then
'    ConnectDB
'End If
'
'Set RS = New ADODB.Recordset
'
'RS.CursorType = adOpenKeyset
'RS.LockType = adLockOptimistic
'RS.Open sTable, Cnn, , , adCmdUnknown
'
'RS.AddNew
'    RS!CustomerID = iCustomerID
'    RS!ContactID = iContactID
'    RS!CallCodeId = iCallCodeID
'    RS!ProductId = iProductID
'    RS!EmployeeID = iEmployeeID
'    RS!NoteDate = dNoteDate
'    RS!Note = sNote
'    RS!EntryDate = dEntryDate
'    RS!CallTime = iCallTime
'    RS!CaseID = iCaseID
'RS.Update
'MsgBox RS!CaseID & vbCrLf & RS!ID, vbInformation, "Remote Server - AppendCall"
'If RS!CaseID = 0 Then RS!CaseID = RS!ID
'
'AppendCall = RS!CaseID
'
'Set RS = Nothing
'
'On Error GoTo 0 ' Disable error trapping.
'Exit Function
'
'ErrorHandler:
'    ErrorCondition = True
'    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "DBModule - AppendCall"
'    Resume Next
'
'End Function
Function GetHistoryRS(sQuery As String, rsCalls As ADODB.Recordset) As Integer
Dim ErrorCondition As Integer

On Error GoTo ErrorHandler    ' Enable error trapping.
GetHistoryRS = 0

If Not (Cnn Is Nothing) Then
    ConnectDB
End If

Set RS = New ADODB.Recordset

RS.CursorType = adOpenKeyset
RS.LockType = adLockOptimistic
RS.Open sQuery, Cnn, adOpenForwardOnly, , adCmdUnknown
Set rsCalls = RS

Set RS = Nothing

On Error GoTo 0 ' Disable error trapping.
GetHistoryRS = 1
Exit Function

ErrorHandler:
    ErrorCondition = True
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation
    Err.Raise 30000 + Err.Number, "DBServer-GetHistoryRS", Err.Description
End Function
Function GetEmployeeRS(sQuery As String, oData As Object) As Integer
Dim ErrorCondition As Integer

On Error GoTo DBErrorHandler    ' Enable error trapping.
GetEmployeeRS = 0

If Not ErrorCondition Then
    On Error GoTo TableErrorHandler ' Enable error trapping.
    ConnectDB
    Set RS = New ADODB.Recordset

    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    RS.Open sQuery, Cnn, , , adCmdUnknown
End If

Do Until RS.EOF
    oData.Add _
        ID:=RS!ID, sName:=RS!sName
'        ID:=RS.Fields("ID").Value, _
'        sName:=RS.Fields("sName").Value
    RS.MoveNext
Loop

Set RS = Nothing

On Error GoTo 0 ' Disable error trapping.
GetEmployeeRS = 1
Exit Function

DBErrorHandler:
TableErrorHandler:
    ErrorCondition = True
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation
    Err.Raise 30000 + Err.Number, "DBServer-GetEmployeeRS", Err.Description
    Resume Next
End Function
Function AppendEntries(sTable As String, oCol As Object) As Integer
'==================================================================
' This function appends entries from Call Tracker V1 which only records
' the call time.
'==================================================================
Dim ErrorCondition As Integer
Dim vCounter As Variant
Dim dEntryDate As Date

On Error GoTo DBErrorHandler    ' Enable error trapping.
AppendEntries = 0

If Not ErrorCondition Then
    On Error GoTo TableErrorHandler ' Enable error trapping.
    ConnectDB
    Set RS = New ADODB.Recordset

    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    RS.Open sTable, Cnn, , , adCmdUnknown
End If

dEntryDate = Now()
If oCol.Count > 0 Then
    For Each vCounter In oCol
        RS.AddNew
            RS!NoteDate = vCounter.EDate
            RS!CallTime = CInt(vCounter.ETime)
            RS!EmplID = vCounter.EmplID
            RS!DateEntered = dEntryDate
'Required for DB integrety, all for time only
            RS!CompanyID = 6
            RS!ContactID = 10
            RS!ProductId = 14
            RS!CallCodeId = 6
        RS.Update
    Next vCounter
End If

Set RS = Nothing

On Error GoTo 0 ' Disable error trapping.
AppendEntries = 1
Exit Function

DBErrorHandler:
TableErrorHandler:
    ErrorCondition = True
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "DBModule - AppendCall"
    Err.Raise 30000 + Err.Number, "DBServer-AppendEntries", Err.Description
    Resume Next

End Function
'Function AppendCustomer(CompanyName As String, Address As String, City As String, State As String, Zip As String, Country As String, Phone As String, Fax As String, iEmplID As Long, sType As String) As Long
'Dim ErrorCondition As Integer
'Dim dEntryDate As Date
'Dim sTable As String
'
'On Error GoTo DBErrorHandler    ' Enable error trapping.
'AppendCustomer = 0
'sTable = "Company"
'
'If Not ErrorCondition Then
'    On Error GoTo TableErrorHandler ' Enable error trapping.
'    ConnectDB
'    Set RS = New ADODB.Recordset
'
'    RS.CursorType = adOpenKeyset
'    RS.LockType = adLockOptimistic
'    RS.Open sTable, Cnn, , , adCmdUnknown
'
'    dEntryDate = Now()
'
'    RS.AddNew
'        RS!CompanyName = CompanyName
'        RS!Address = Address
'        RS!City = City
'        RS!State = Left(State, 2)
'        RS!Zip = Zip
'        RS!Phone = Phone
'        RS!Country = Country
'        RS!Fax = Fax
'        RS!EntryDate = dEntryDate
'        RS!EmplID = iEmplID
'        If sType <> vbNull Then RS!Type = sType
'    RS.Update
'
'    AppendCustomer = RS!ID
'
'    Set RS = Nothing
'
'    On Error GoTo 0 ' Disable error trapping.
'    Exit Function
'End If
'
'DBErrorHandler:
'TableErrorHandler:
'    ErrorCondition = True
'    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "DBModule - AppendCustomer"
'    Exit Function
'End Function
'Function AppendContact(FirstName As String, LastName As String, Title As String, iEmplID As Long, Skill As String, Training As String) As Long
'Dim ErrorCondition As Integer
'Dim dEntryDate As Date
'Dim sTable As String
'
'On Error GoTo DBErrorHandler    ' Enable error trapping.
'AppendContact = 0
'sTable = "Contact"
'
'If Not ErrorCondition Then
'    On Error GoTo TableErrorHandler ' Enable error trapping.
'    ConnectDB
'    Set RS = New ADODB.Recordset
'
'    RS.CursorType = adOpenKeyset
'    RS.LockType = adLockOptimistic
'    RS.Open sTable, Cnn, , , adCmdUnknown
'End If
'
'dEntryDate = Now()
'
'RS.AddNew
'    RS!FirstName = FirstName
'    RS!LastName = LastName
'    RS!Title = Title
'    RS!EmplID = iEmplID
'    RS!DateEntered = dEntryDate
'    If Skill <> vbNull Then RS!Skill = Skill
'    If Training <> vbNull Then RS!Training = Training
'RS.Update
'
'AppendContact = RS!ID
'
'Set RS = Nothing
'
'Exit Function
'
'DBErrorHandler:
'TableErrorHandler:
'    ErrorCondition = True
'    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "DBModule - AppendCustomer"
'    Exit Function
'End Function
'Function AppendCallCode(CallType As String, iEmplID As Long) As Long
'Dim ErrorCondition As Integer
'Dim dEntryDate As Date
'Dim sTable As String
'
'On Error GoTo DBErrorHandler    ' Enable error trapping.
'AppendCallCode = 0
'sTable = "CallCode"
'
'If Not ErrorCondition Then
'    On Error GoTo TableErrorHandler ' Enable error trapping.
'    ConnectDB
'    Set RS = New ADODB.Recordset
'
'    RS.CursorType = adOpenKeyset
'    RS.LockType = adLockOptimistic
'    RS.Open sTable, Cnn, , , adCmdUnknown
'End If
'
'dEntryDate = Now()
'
'RS.AddNew
'    RS!CallType = CallType
'    RS!EmplID = iEmplID
'    RS!DateEntered = dEntryDate
'RS.Update
'
'AppendCallCode = RS!ID
'
'Set RS = Nothing
'
'Exit Function
'
'DBErrorHandler:
'TableErrorHandler:
'    ErrorCondition = True
'    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "DBModule - AppendCustomer"
'    Exit Function
'End Function
'Function AppendCompanyLink(Company As Long, Contact As Long, iEmplID As Long) As Long
'Dim ErrorCondition As Integer
'Dim dEntryDate As Date
'Dim sTable As String
'
'On Error GoTo DBErrorHandler    ' Enable error trapping.
'AppendCompanyLink = 0
'sTable = "[Contact Link]"
'
'If Not ErrorCondition Then
'    On Error GoTo TableErrorHandler ' Enable error trapping.
'    ConnectDB
'    Set RS = New ADODB.Recordset
'
'    RS.CursorType = adOpenKeyset
'    RS.LockType = adLockOptimistic
'    RS.Open sTable, Cnn, , , adCmdUnknown
'End If
'
'dEntryDate = Now()
'
'RS.AddNew
'    RS!CompanyID = Company
'    RS!ContactID = Contact
'    RS!EmplID = iEmplID
'    RS!DateEntered = dEntryDate
'RS.Update
'
'AppendCompanyLink = RS!ID
'
'Set RS = Nothing
'
'Exit Function
'
'DBErrorHandler:
'TableErrorHandler:
'    ErrorCondition = True
'    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "DBModule - AppendLink"
'    Exit Function
'End Function
Function EditData(sTable As String, oData As Object) As Integer
'================================================================
'This function edits EXISTING records to the Call Tracker database.  The
'records can be customers/contacts.
'
'First it checks for an existing db connection then
'Creates a opens the specified db table (sTable)
'Finds the record using the data ID for the sTable selection
'Installs the changes
'
'Return is the data ID for the record; 0 is a failure.
'================================================================
Dim ErrorCondition As Integer
Dim vCounter As Variant
Dim dEntryDate As Date

On Error GoTo ErrorHandler    ' Enable error trapping.
EditData = 0   'Assume failure

If Not (Cnn Is Nothing) Then
    ConnectDB
End If

Set RS = New ADODB.Recordset

RS.CursorType = adOpenKeyset
RS.LockType = adLockOptimistic
RS.Open sTable, Cnn, , , adCmdUnknown

dEntryDate = Now()

For Each vCounter In oData
    
    RS.Find "ID = '" & vCounter.ID & "'"
    
    With vCounter
        Select Case sTable
            Case "SupportCalls"
            Case "Company"
                RS!CompanyName = .CompanyName
                RS!Address = .Address
                RS!City = .City
                RS!State = .State
                RS!Zip = .Zip
                RS!Phone = .Phone
                RS!Country = .Country
                RS!Fax = .Fax
                RS!Type = .Type
            Case "Contact"
                RS!FirstName = .FirstName
                RS!LastName = .LastName
                RS!Title = .Title
                RS!Skill = .Skill
                RS!Training = .Training
            Case "CallCode"
            Case "Contact Link"
                RS!CompanyID = .Company
                RS!ContactID = .Contact
        End Select
        
        RS!EmplID = .EmplID
        RS!DateEntered = dEntryDate
    
    End With 'vCounter
    RS.Update

EditData = RS!ID
            
Next vCounter

Set RS = Nothing
On Error GoTo 0 ' Disable error trapping.
Exit Function

ErrorHandler:
    ErrorCondition = True
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "DBModule - EditData"
    Err.Raise 30000 + Err.Number, "DBServer-EditData", Err.Description

End Function
Private Sub Class_Terminate()
    CloseDB
End Sub
