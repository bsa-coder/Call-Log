Attribute VB_Name = "DBModule"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EBC8610046"
Option Explicit
'##ModelId=39EBC86100D5
'Dim WS As Workspace
'Dim DB As ADODB
'##ModelId=39EBC86100E7
Dim RS As adodb.Recordset
'##ModelId=39EBC86100F3
Dim Cnn As New adodb.Connection

'##ModelId=39EBC86100FB
Sub ConnectDB()
Dim sConn As String
On Error GoTo ErrorHandler

' Open database.
Set Cnn = New adodb.Connection
Cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\Call Log\CALL LOADING.MDB;Persist Security Info=False"
sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\Call Log\CALL LOADING.MDB;Persist Security Info=False"
Cnn.Open sConn
Exit Sub

ErrorHandler:
MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, " Error: ConnectDB"
Resume Next

End Sub
'##ModelId=39EBC861012D
Sub CloseDB()
Cnn.Close
Set Cnn = Nothing
End Sub
'##ModelId=39EBC861015F
Function GetListRS2(sSql As String, oData As Object)
Dim ErrorCondition As Integer

On Error GoTo DBErrorHandler    ' Enable error trapping.
If Not ErrorCondition Then
    On Error GoTo TableErrorHandler ' Enable error trapping.
    ConnectDB
    Set RS = New adodb.Recordset
    RS.Open sSql, Cnn, , , adCmdUnknown
End If

Do Until RS.EOF
    oData.Add sName:=RS.Fields("sName").Value, ID:=RS.Fields("ID").Value ', sKey:=RS.Fields("ID").Value
    RS.MoveNext
Loop

Set RS = Nothing
CloseDB

On Error GoTo 0 ' Disable error trapping.
GetListRS2 = 1
Exit Function

DBErrorHandler:
TableErrorHandler:
    ErrorCondition = True
    MsgBox Err.Number & vbCrLf & Err.Description & sSql, vbExclamation, "GetListRS"
    Resume Next

End Function
'##ModelId=3A0F61E202AE
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

If Not (TypeOf oData Is CLinks) Then
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
        ContactID:=RS!ContactID
        
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
'##ModelId=39EBC861024F
Function GetLinkRecordset2(sSql As String, colData As Object) As Integer
Dim ErrorCondition As Integer

On Error GoTo DBErrorHandler    ' Enable error trapping.
If Not ErrorCondition Then
    On Error GoTo TableErrorHandler ' Enable error trapping.
    ConnectDB
    Set RS = New adodb.Recordset
    RS.Open sSql, Cnn, , , adCmdUnknown
    RS.MoveFirst
End If

Do Until RS.EOF
    colData.Add ID:=RS.Fields("ID").Value, CompanyID:=RS.Fields("CompanyID").Value, ContactID:=RS.Fields("ContactID").Value
    RS.MoveNext
Loop

Set RS = Nothing
CloseDB

On Error GoTo 0 ' Disable error trapping.
GetLinkRecordset2 = 1
Exit Function

DBErrorHandler:
TableErrorHandler:
    ErrorCondition = True
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation
    Resume Next

End Function
'##ModelId=39EBC8610372
Function AppendCall(sTable As String, iCustomerID As Long, iContactID As Long, iCallCodeID As Long, iProductID As Long, iEmployeeID As Long, dNoteDate As Date, sNote As String, dEntryDate As Date, iCallTime As Integer) As Integer
Dim ErrorCondition As Integer

On Error GoTo DBErrorHandler    ' Enable error trapping.
AppendCall = 0

If Not ErrorCondition Then
    On Error GoTo TableErrorHandler ' Enable error trapping.
    ConnectDB
    Set RS = New adodb.Recordset

    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    RS.Open sTable, Cnn, , , adCmdUnknown
End If

RS.AddNew
    RS!CustomerID = iCustomerID
    RS!ContactID = iContactID
    RS!callcodeid = iCallCodeID
    RS!ProductID = iProductID
    RS!EmployeeID = iEmployeeID
    RS!NoteDate = dNoteDate
    RS!Note = sNote
    RS!EntryDate = dEntryDate
    RS!CallTime = iCallTime
RS.Update

Set RS = Nothing
CloseDB

On Error GoTo 0 ' Disable error trapping.
AppendCall = 1
Exit Function

DBErrorHandler:
TableErrorHandler:
    ErrorCondition = True
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "DBModule - AppendCall"
    Resume Next

End Function
'##ModelId=39EBC8620355
Function GetHistoryRS(sQry As String, oCalls As Object) As Integer
Dim CallTime As Integer

On Error GoTo ErrorHandler    ' Enable error trapping.
GetHistoryRS = 0

If Not (Cnn Is Nothing) Then
    ConnectDB
End If

Set RS = New adodb.Recordset

RS.CursorType = adOpenKeyset
RS.LockType = adLockOptimistic
RS.Open sQry, Cnn, adOpenForwardOnly, , adCmdUnknown

Do Until RS.EOF
    If (RS.Fields("iCallTime").Value > 0) Then CallTime = RS.Fields("iCallTime").Value Else CallTime = 0
    oCalls.Add _
        sLastName:=RS.Fields("LastName").Value, _
        sContactName:=RS.Fields("ContactName").Value, _
        sCompanyName:=RS.Fields("CompanyName").Value, _
        sProductName:=RS.Fields("ProductName").Value, _
        sCallType:=RS.Fields("CallType").Value, _
        dNoteDate:=RS.Fields("NoteDate").Value, _
        sNote:=RS.Fields("sNote").Value, _
        iCallTime:=CallTime
    RS.MoveNext
Loop

Set RS = Nothing
CloseDB

On Error GoTo 0 ' Disable error trapping.
GetHistoryRS = 1
Exit Function

ErrorHandler:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation
    Resume Next
End Function
'##ModelId=39EBC863005D
Function GetEmployeeRS(sQuery As String, oData As Object) As Integer
Dim ErrorCondition As Integer

On Error GoTo DBErrorHandler    ' Enable error trapping.
GetEmployeeRS = 0

If Not ErrorCondition Then
    On Error GoTo TableErrorHandler ' Enable error trapping.
    ConnectDB
    Set RS = New adodb.Recordset

    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    RS.Open sQuery, Cnn, , , adCmdUnknown
End If

Do Until RS.EOF
    oData.Add _
        ID:=RS.Fields("ID").Value, _
        sName:=RS.Fields("sName").Value
    RS.MoveNext
Loop

Set RS = Nothing
CloseDB

On Error GoTo 0 ' Disable error trapping.
GetEmployeeRS = 1
Exit Function

DBErrorHandler:
TableErrorHandler:
    ErrorCondition = True
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation
    Resume Next
End Function
'##ModelId=39EBC8630158
Function AppendEntries(sTable As String, oCol As Object) As Integer
Dim ErrorCondition As Integer
Dim vCounter As Variant
Dim dEntryDate As Date

On Error GoTo DBErrorHandler    ' Enable error trapping.
AppendEntries = 0

If Not ErrorCondition Then
    On Error GoTo TableErrorHandler ' Enable error trapping.
    ConnectDB
    Set RS = New adodb.Recordset

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
            RS!EmployeeID = vCounter.EmplID
            RS!EntryDate = dEntryDate
'Required for DB integrety, all for time only
            RS!CustomerID = 6
            RS!ContactID = 10
            RS!ProductID = 14
            RS!callcodeid = 6
        RS.Update
    Next vCounter
End If

Set RS = Nothing
CloseDB

On Error GoTo 0 ' Disable error trapping.
AppendEntries = 1
Exit Function

DBErrorHandler:
TableErrorHandler:
    ErrorCondition = True
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "DBModule - AppendCall"
    Resume Next

End Function
'##ModelId=39EBC8630252
Function AppendCustomer(CompanyName As String, Address As String, City As String, State As String, Zip As String, Country As String, Phone As String, Fax As String, iEmplID As Long) As Long
Dim ErrorCondition As Integer
Dim dEntryDate As Date
Dim sTable As String

On Error GoTo DBErrorHandler    ' Enable error trapping.
AppendCustomer = 0
sTable = "Company"

If Not ErrorCondition Then
    On Error GoTo TableErrorHandler ' Enable error trapping.
    ConnectDB
    Set RS = New adodb.Recordset

    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    RS.Open sTable, Cnn, , , adCmdUnknown

    dEntryDate = Now()
    
    RS.AddNew
'        RS!CompanyName = "Gaston Mining"
'        RS!Address = "Main Street"
'        RS!City = "Flagstaff"
'        RS!State = "SD"
'        RS!Zip = "44444"
'        RS!Phone = "333-555-5555"
'        RS!Country = "Country"
'        RS!Fax = "333-555-4444"
'        RS!EntryDate = dEntryDate
'        RS!EmplID = 5
        
        RS!CompanyName = CompanyName
        RS!Address = Address
        RS!City = City
        RS!State = Left(State, 2)
        RS!Zip = Zip
        RS!Phone = Phone
        RS!Country = Country
        RS!Fax = Fax
        RS!EntryDate = dEntryDate
        RS!EmplID = iEmplID
    RS.Update
    
    AppendCustomer = RS!ID
    
    Set RS = Nothing
    CloseDB

    On Error GoTo 0 ' Disable error trapping.
    Exit Function
End If

DBErrorHandler:
TableErrorHandler:
    ErrorCondition = True
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "DBModule - AppendCustomer"
    Exit Function
End Function
'##ModelId=39EBC864020D
Function AppendContact(FirstName As String, LastName As String, Title As String, iEmplID As Long) As Long
Dim ErrorCondition As Integer
Dim dEntryDate As Date
Dim sTable As String

On Error GoTo DBErrorHandler    ' Enable error trapping.
AppendContact = 0
sTable = "Contact"

If Not ErrorCondition Then
    On Error GoTo TableErrorHandler ' Enable error trapping.
    ConnectDB
    Set RS = New adodb.Recordset

    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    RS.Open sTable, Cnn, , , adCmdUnknown
End If

dEntryDate = Now()

RS.AddNew
    RS!FirstName = FirstName
    RS!LastName = LastName
    RS!Title = Title
    RS!EmplID = iEmplID
    RS!DateEntered = dEntryDate
RS.Update

AppendContact = RS!ID

Set RS = Nothing
CloseDB

Exit Function

DBErrorHandler:
TableErrorHandler:
    ErrorCondition = True
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "DBModule - AppendCustomer"
    Exit Function
End Function
'##ModelId=39EBC86403DA
Function AppendCallCode(CallType As String, iEmplID As Long) As Long
Dim ErrorCondition As Integer
Dim dEntryDate As Date
Dim sTable As String

On Error GoTo DBErrorHandler    ' Enable error trapping.
AppendCallCode = 0
sTable = "CallCode"

If Not ErrorCondition Then
    On Error GoTo TableErrorHandler ' Enable error trapping.
    ConnectDB
    Set RS = New adodb.Recordset

    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    RS.Open sTable, Cnn, , , adCmdUnknown
End If

dEntryDate = Now()

RS.AddNew
    RS!CallType = CallType
    RS!EmplID = iEmplID
    RS!DateEntered = dEntryDate
RS.Update

AppendCallCode = RS!ID

Set RS = Nothing
CloseDB

Exit Function

DBErrorHandler:
TableErrorHandler:
    ErrorCondition = True
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "DBModule - AppendCustomer"
    Exit Function
End Function
'##ModelId=39EBC86500F6
Function AppendCompanyLink(Company As Long, Contact As Long, iEmplID As Long) As Long
Dim ErrorCondition As Integer
Dim dEntryDate As Date
Dim sTable As String

On Error GoTo DBErrorHandler    ' Enable error trapping.
AppendCompanyLink = 0
sTable = "[Contact Link]"

If Not ErrorCondition Then
    On Error GoTo TableErrorHandler ' Enable error trapping.
    ConnectDB
    Set RS = New adodb.Recordset

    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    RS.Open sTable, Cnn, , , adCmdUnknown
End If

dEntryDate = Now()

RS.AddNew
    RS!CompanyID = Company
    RS!ContactID = Contact
    RS!EmplID = iEmplID
    RS!DateEntered = dEntryDate
RS.Update

AppendCompanyLink = RS!ID

Set RS = Nothing
CloseDB

Exit Function

DBErrorHandler:
TableErrorHandler:
    ErrorCondition = True
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "DBModule - AppendLink"
    Exit Function
End Function


