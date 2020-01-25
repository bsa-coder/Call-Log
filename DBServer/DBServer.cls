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
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B4FC2D2032A"
Option Explicit
'Dim WS As Workspace
'##ModelId=3B4FC2D3008F
Dim RS As ADODB.Recordset
'##ModelId=3B4FC2D300A3
Dim Cnn As New ADODB.Connection

'##ModelId=3B4FC2D300AB
Const NOCUSTOMER As Long = 6
'##ModelId=3B4FC2D300FB
Const NOCONTACT As Long = 10
'##ModelId=3B4FC2D30173
Const NOPRODUCT As Long = 14
'##ModelId=3B4FC2D301D7
Const NOCODE As Long = 6

'##ModelId=3B4FC2D30246
Const clCOMPANY As Integer = 0
'##ModelId=3B4FC2D302B4
Const clCONTACT As Integer = 1
'##ModelId=3B4FC2D30336
Const clPRODUCT As Integer = 2 ''
'##ModelId=3B4FC2D303B8
Const clCALLCODE As Integer = 3 ''
'##ModelId=3B4FC2D4005C
Const clCALL As Integer = 4
'##ModelId=3B4FC2D400F3
Const clLINK As Integer = 5
'##ModelId=3B4FC2D40193
Const clHISTORY As Integer = 6
'##ModelId=3B4FC2D4023D
Const clCALLTYPE As Integer = 7
'##ModelId=3B4FC2D402E7
Const clEMPLOYEE As Integer = 8
'##ModelId=3B4FC2D4039B
Const clCOMPANYUPDATE As Integer = 9
'##ModelId=3B4FC2D5007C
Const clCONTACTUPDATE As Integer = 10
'##ModelId=3B4FC2D5014E
Const clPRODUCTHISTORY As Integer = 11
'##ModelId=3B4FC2D50220
Const clCONTACT2 As Integer = 12
'##ModelId=3B4FC2D502FD
Const clCOMPANY2 As Integer = 13

'##ModelId=3B4FC2D60005
Dim iCurrentRecordID As Long
'##ModelId=3B4FC2D600EB
Dim sCurrentTable As String
'##ModelId=3B4FC2D601DC
Dim dEntryDate As Date

'##ModelId=3B4FC2DB005C
Public Event OnAppendDone()
Public Event OnNewCustomerAppendDone()

'##ModelId=3B4FC2D602D6
Function AppendData(sTable As String, rsData As ADODB.Recordset) As Long
'================================================================
'This function appends (or updates) records to the Call Tracker
'database.  The new records can be calls/customers/contacts/callcodes
'
'First it checks for an existing db connection then
'Creates a opens the specified db table (sTable)
'Appends record depending or the sTable selection
'
'Return is the new ID for the record; 0 is a failure.
'================================================================

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

If ((sTable = "Company" Or sTable = "Contact" Or sTable = "SupportCalls") And rsData!ID <> 0) Then 'Editing Record
    RS.Find "ID = " & rsData!ID 'Editing a record
Else
    RS.AddNew
End If

RS!EmplID = rsData!EmplID
RS!DateEntered = dEntryDate

Select Case sTable
    Case "Company"
        RS!CompanyName = Trim(Left(rsData!sName, 100))
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

iCurrentRecordID = RS!ID
sCurrentTable = sTable
AppendData = RS!ID
            
If sTable = "SupportCalls" Then
    If RS!CaseID = 0 Then
        RS!CaseID = RS!ID
        RS.Update
    End If
End If

CheckAppend

Set RS = Nothing
On Error GoTo 0 ' Disable error trapping.
Exit Function

ErrorHandler:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "DB::AppendData - Error"
    Resume Next
End Function
'##ModelId=3B4FC2D700E3
Private Sub CheckAppend()
Dim sQuery As String
Dim sSelect As String
Dim sWhere As String
Dim sFrom As String
Dim sOrderBy As String
Dim iCounter As Integer

sSelect = "SELECT [DateEntered] "
sFrom = "FROM " & sCurrentTable & " "
sWhere = "WHERE (ID = " & iCurrentRecordID & ")"
sOrderBy = ";"
sQuery = sSelect & sFrom & sWhere & sOrderBy

If Not (Cnn Is Nothing) Then
    ConnectDB
End If

Set RS = New ADODB.Recordset

RS.CursorType = adOpenKeyset
RS.LockType = adLockOptimistic
RS.Open sQuery, Cnn, , , adCmdUnknown

Do While RS!DateEntered <> dEntryDate
    RS.Open sQuery, Cnn, , , adCmdUnknown
    RS.Close
    For iCounter = 0 To 1000
        iCounter = iCounter + 1
    Next iCounter
Loop

If sCurrentTable = "Company" Then
    RaiseEvent OnNewCustomerAppendDone
Else
    RaiseEvent OnAppendDone
End If

End Sub
'##ModelId=3B4FC2D70183
Public Function OnAppendComplete() As Boolean
OnAppendComplete = True
RaiseEvent OnAppendDone
End Function
'##ModelId=3B4FC2D701FB
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

Do While Not (Cnn.State = adStateOpen)
    Exit Do
Loop
'If Cnn.Properties.Item(1).Value Then Resume Next
Exit Sub

ErrorHandler:
    MsgBox Err.Number & " from " & Err.Source & vbCrLf & Err.Description, vbCritical, "DB-GetListRS"
    Resume Next

End Sub
'##ModelId=3B4FC2D70273
Sub CloseDB()
If (Cnn.State = adStateOpen) Then
    Cnn.Close
End If
Set Cnn = Nothing
End Sub
'##ModelId=3B4FC2D702E2
Public Function GetListRS(sQuery As String, ByRef rsData As ADODB.Recordset, iIndex As Integer)
'=========================================================
'GetListRS passes a sql string to retrieve listbox data
'and fills the recordset
'sQuery is sql SELECT query
'rsData to hold the data
'=========================================================
Dim iCounter As Long
On Error GoTo ErrorHandler    ' Enable error trapping.
GetListRS = 0 'assume failure

'For iCounter = 0 To 100000
'    iCounter = iCounter + 1
'Next iCounter

If Not (Cnn Is Nothing) Then
    ConnectDB
End If

Set RS = New ADODB.Recordset


RS.CursorLocation = adUseClient
RS.CursorType = adOpenKeyset
RS.LockType = adLockOptimistic
'RS.Open sQuery, Cnn, adOpenForwardOnly, , adCmdUnknown
'Set RS = Nothing
'Set RS = New ADODB.Recordset
'RS.Open sQuery, Cnn, adOpenForwardOnly, , adCmdUnknown
RS.Open sQuery, Cnn, adOpenStatic, adLockReadOnly ', adCmdTable

'If CheckDirty Then 'try again
'    RS.Open sQuery, Cnn, adOpenForwardOnly, , adCmdUnknown
'End If

Set RS.ActiveConnection = Nothing
Set rsData = RS
Set RS = Nothing

On Error GoTo 0 ' Disable error trapping.

GetListRS = 1
Exit Function

ErrorHandler:
    MsgBox Err.Number & " from " & Err.Source & vbCrLf & Err.Description & vbCrLf & sQuery, vbCritical, "DB-GetListRS"
    Resume Next
End Function
'##ModelId=3B4FC2D801CB
Private Function CheckDirty() As Integer
Dim ImDirty As Integer
CheckDirty = 0

If Not RS!Dirty Is Nothing Then ImDirty = RS!Dirty + 1
RS.MoveFirst
CheckDirty = ImDirty - CInt(ImDirty / 2) * 2

End Function
'##ModelId=3B4FC2D8027F
Function GetHistoryRS(sQuery As String, rsCalls As ADODB.Recordset) As Integer
'=========================================================
'OBSOLETE GetHistoryRS is an exact duplicate of GetListRS
'So do not use it.
'=========================================================
GetHistoryRS = GetListRS(sQuery, rsCalls, 1)
'On Error GoTo ErrorHandler    ' Enable error trapping.
'GetHistoryRS = 0
'
'If Not (Cnn Is Nothing) Then
'    ConnectDB
'End If
'
'Set RS = New ADODB.Recordset
'
'RS.CursorType = adOpenKeyset
'RS.LockType = adLockOptimistic
'RS.Open sQuery, Cnn, adOpenForwardOnly, , adCmdUnknown
'
'Set rsCalls = RS
'Set RS = Nothing
'
'On Error GoTo 0 ' Disable error trapping.
'GetHistoryRS = 1
'Exit Function
'
'ErrorHandler:
'    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation
'    Resume Next
End Function
'##ModelId=3B4FC2D900FA
Function GetEmployeeRS(sQuery As String, oData As Object) As Integer
Dim ErrorCondition As Integer

On Error GoTo DBErrorHandler    ' Enable error trapping.
GetEmployeeRS = 0

If Not ErrorCondition Then
    On Error GoTo TableErrorHandler ' Enable error trapping.
    ConnectDB
    Set RS = New ADODB.Recordset

    RS.CursorLocation = adUseClient
    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    RS.Open sQuery, Cnn, , , adCmdUnknown
    
    Set RS.ActiveConnection = Nothing
    
End If

Do Until RS.EOF
    oData.Add ID:=RS!ID, sName:=RS!sName
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
    Resume Next
End Function
'##ModelId=3B4FC2D9030C
Function AppendEntries(sTable As String, oCol As Object) As Integer
'==================================================================
' This function appends entries from Call Tracker V1 which only records
' the call time.
'==================================================================
Dim ErrorCondition As Integer
Dim vCounter As Variant
Dim dEntryDate As Date
Dim oEntries As Object

On Error GoTo DBErrorHandler    ' Enable error trapping.
AppendEntries = 0
'Set oEntries = oCol

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
            RS!Note = " "
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
    Resume Next

End Function
'##ModelId=3B4FC2DA0141
Function EditData(sTable As String, oData As Object) As Integer
'================================================================
'This function is obsolete.
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
EditData = AppendData(sTable, oData)
'Dim ErrorCondition As Integer
'Dim vCounter As Variant
'Dim dEntryDate As Date
'
'On Error GoTo ErrorHandler    ' Enable error trapping.
'EditData = 0   'Assume failure
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
'dEntryDate = Now()
'
'For Each vCounter In oData
'
'    RS.Find "ID = '" & vCounter.ID & "'"
'
'    With vCounter
'        Select Case sTable
'            Case "SupportCalls"
'            Case "Company"
'                RS!CompanyName = .CompanyName
'                RS!Address = .Address
'                RS!City = .City
'                RS!State = .State
'                RS!Zip = .Zip
'                RS!Phone = .Phone
'                RS!Country = .Country
'                RS!Fax = .Fax
'                RS!Type = .Type
'            Case "Contact"
'                RS!FirstName = .FirstName
'                RS!LastName = .LastName
'                RS!Title = .Title
'                RS!Skill = .Skill
'                RS!Training = .Training
'            Case "CallCode"
'            Case "Contact Link"
'                RS!CompanyID = .Company
'                RS!ContactID = .Contact
'        End Select
'
'        RS!EmplID = .EmplID
'        RS!DateEntered = dEntryDate
'
'    End With 'vCounter
'    RS.Update
'
'EditData = RS!ID
'
'Next vCounter
'
'Set RS = Nothing
'On Error GoTo 0 ' Disable error trapping.
'Exit Function
'
'ErrorHandler:
'    ErrorCondition = True
'    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "DBModule - EditData"
'    Resume Next
'
End Function
'##ModelId=3B4FC2DA039A
Private Sub Class_Terminate()
    CloseDB
If Not RS Is Nothing Then
    Set RS = Nothing
End If

End Sub

