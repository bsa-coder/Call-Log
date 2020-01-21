VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BLServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

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
Const clPRODUCTHISTORY As Integer = 11

Dim rsData As ADODB.Recordset
Dim rsLBData As ADODB.Recordset

Dim DBServer As CallTrackerDBServer.DBServer
'local variable(s) to hold property value(s)
Private mvarStatus As Integer 'local copy
'Public Property Let Status(ByVal vData As Integer)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.Status = 5
'    mvarStatus = vData
'End Property
Public Property Get Status() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Status
    Status = mvarStatus
End Property
Private Function MakeQuery(Index As Integer, Optional Filter As Long) As String
Dim sQuery As String
Dim sSelect As String
Dim sFrom As String
Dim sWhere As String
Dim sOrderBy As String
Dim sFilter As String

On Error GoTo ErrorHandler
mvarStatus = 10 'MakeQuery started

MakeQuery = "" ' assume failure

Select Case Index
    Case clCOMPANY
        If Filter <> 0 Then
            sFilter = " AND ([Company]![ID]=" & CStr(Filter) & ")"
        Else
            sFilter = ""
        End If
        
        sSelect = "SELECT Company.ID as ID, [Company]![CompanyName] & ', ' & "
        sSelect = sSelect & "[Company]![City] & ', ' & [Company]![State] & ', ' & "
        sSelect = sSelect & "[Company]![Address] AS sName, [Company]![Type] as cType "
        sFrom = "FROM Company "
        sWhere = "WHERE (([Company]![ID]<>" & NOCUSTOMER & ")" & sFilter & ") "
        sOrderBy = "ORDER BY [Company]![CompanyName] & ', ' & [Company]![State] & ', ' & [Company]![City];"
    Case clCONTACT 'working
        If Filter <> 0 Then
            sFilter = " AND ([Contact]![ID]=" & CStr(Filter) & ")"
        Else
            sFilter = ""
        End If
        
        sSelect = "SELECT [Contact]![ID] AS ID, [Contact]![FirstName] & ' ' & "
        sSelect = sSelect & "[Contact]![LastName] & ', ' & [Contact]![Title] & ' :: ' & "
        sSelect = sSelect & "[Contact]![Phone] AS sName, [Contact]![Skill] as Skill, [Contact]![Training] as Training "
        sFrom = "FROM Contact "
        sWhere = "WHERE (([Contact]![ID]<>" & NOCONTACT & ")" & sFilter & ") "
        sOrderBy = "ORDER BY [Contact]![FirstName] & ' ' & [Contact]![LastName] & ', ' & [Contact]![Title];"
    Case clPRODUCT 'working
        sSelect = "SELECT [Product]![ProductName] AS sName, Product.ID "
        sFrom = "From Product "
        sWhere = "WHERE ([Product]![ID]<>" & NOPRODUCT & ") "
        sOrderBy = "ORDER BY [Product]![ProductName];"
    Case clCALLCODE 'working
        If Filter <> 0 Then
            sFilter = " AND ([CallCode]![ID]=" & CStr(Filter) & ")"
        Else
            sFilter = ""
        End If
        
        sSelect = "SELECT [CallCode]![CallType] AS sName, CallCode.ID "
        sFrom = "From CallCode "
        sWhere = "WHERE (([CallCode]![ID]<>" & NOCODE & ")" & sFilter & ") "
        sOrderBy = "ORDER BY [CallCode]![CallType];"
    Case clLINK 'working
        sSelect = "SELECT [Contact Link].* "
        sFrom = "FROM [Contact Link];"
        sWhere = ""
        sOrderBy = ""
    Case clEMPLOYEE 'working
        sSelect = "SELECT Employees.ID AS ID, "
        sSelect = sSelect & "IIf(Len([Employees]![FirstName])=1,[Employees]![MiddleName],[Employees]![FirstName]) & ' ' & [Employees]![LastName] AS sName "
        sFrom = "FROM Employees "
        sWhere = ""
        sOrderBy = "ORDER BY Employees.LastName;"
    Case clHISTORY 'working
        sSelect = "SELECT Employees.LastName AS LastName, "
        sSelect = sSelect & "[Contact]![FirstName] & ' ' & "
        sSelect = sSelect & "[Contact]![LastName] & ', ' & "
        sSelect = sSelect & "[Contact]![Title] AS ContactName, "
        sSelect = sSelect & "[Contact]![Phone] AS Phone, "
        sSelect = sSelect & "Company.CompanyName AS CompanyName, "
        sSelect = sSelect & "Product.ProductName AS ProductName, "
        sSelect = sSelect & "CallCode.CallType AS CallType, "
        sSelect = sSelect & "SupportCalls.NoteDate AS NoteDate, "
        sSelect = sSelect & "SupportCalls.Note AS sNote, "
        sSelect = sSelect & "SupportCalls.CallTime AS iCallTime, "
        sSelect = sSelect & "SupportCalls.CaseID AS CaseID "
        
        sFrom = "FROM Product INNER JOIN "
        sFrom = sFrom & "(Employees INNER JOIN "
        sFrom = sFrom & "(Contact INNER JOIN "
        sFrom = sFrom & "(Company INNER JOIN "
        sFrom = sFrom & "(CallCode INNER JOIN SupportCalls "
        sFrom = sFrom & "ON CallCode.ID = SupportCalls.CallCodeID) "
        sFrom = sFrom & "ON Company.ID = SupportCalls.CompanyID) "
        sFrom = sFrom & "ON Contact.ID = SupportCalls.ContactID) "
        sFrom = sFrom & "ON Employees.ID = SupportCalls.EmplID) "
        sFrom = sFrom & "ON Product.ID = SupportCalls.ProductID "
    
        sWhere = "Where (((Company.ID) = " & Filter & ")) "
        
        sOrderBy = "ORDER BY SupportCalls.NoteDate DESC;"
    Case clCOMPANYUPDATE
        If Filter <> 0 Then
            sFilter = " AND ([Company]![ID]=" & CStr(Filter) & ")"
        Else
            sFilter = ""
        End If
        
        sSelect = "SELECT * "
        sFrom = "FROM Company "
        sWhere = "WHERE (([Company]![ID]<>" & NOCUSTOMER & ")" & sFilter & ") "
        sOrderBy = "ORDER BY [Company]![CompanyName] & ', ' & [Company]![State] & ', ' & [Company]![City];"
    Case clCONTACTUPDATE
        If Filter <> 0 Then
            sFilter = " AND ([Contact]![ID]=" & CStr(Filter) & ")"
        Else
            sFilter = ""
        End If
        
        sSelect = "SELECT * "
        sFrom = "FROM Contact "
        sWhere = "WHERE (([Contact]![ID]<>" & NOCONTACT & ")" & sFilter & ") "
        sOrderBy = "ORDER BY [Contact]![FirstName] & ' ' & [Contact]![LastName] & ', ' & [Contact]![Title];"
    Case clPRODUCTHISTORY
        sSelect = "SELECT Employees.LastName AS LastName, "
        sSelect = sSelect & "[Contact]![FirstName] & ' ' & "
        sSelect = sSelect & "[Contact]![LastName] & ', ' & "
        sSelect = sSelect & "[Contact]![Title] AS ContactName, "
        sSelect = sSelect & "[Contact]![Phone] AS Phone, "
        sSelect = sSelect & "Company.CompanyName AS CompanyName, "
        sSelect = sSelect & "Product.ProductName AS ProductName, "
        sSelect = sSelect & "CallCode.CallType AS CallType, "
        sSelect = sSelect & "SupportCalls.NoteDate AS NoteDate, "
        sSelect = sSelect & "SupportCalls.Note AS sNote, "
        sSelect = sSelect & "SupportCalls.CallTime AS iCallTime, "
        sSelect = sSelect & "SupportCalls.CaseID AS CaseID, "
        sSelect = sSelect & "SupportCalls.CallCodeID AS CallCodeID "
        
        sFrom = "FROM Product INNER JOIN "
        sFrom = sFrom & "(Employees INNER JOIN "
        sFrom = sFrom & "(Contact INNER JOIN "
        sFrom = sFrom & "(Company INNER JOIN "
        sFrom = sFrom & "(CallCode INNER JOIN SupportCalls "
        sFrom = sFrom & "ON CallCode.ID = SupportCalls.CallCodeID) "
        sFrom = sFrom & "ON Company.ID = SupportCalls.CompanyID) "
        sFrom = sFrom & "ON Contact.ID = SupportCalls.ContactID) "
        sFrom = sFrom & "ON Employees.ID = SupportCalls.EmplID) "
        sFrom = sFrom & "ON Product.ID = SupportCalls.ProductID "
    
        sWhere = "Where (((Product.ID) = " & Filter & ")) "
        
        sOrderBy = "ORDER BY SupportCalls.NoteDate DESC;"
End Select

mvarStatus = 15 'MakeQuery ended

MakeQuery = sSelect & sFrom & sWhere & sOrderBy
Exit Function

ErrorHandler:

End Function
'===================================================================
Public Function GetLbData(iDataIndex As Integer, Optional Filter As Long) As ADODB.Recordset
'==========================================================
'Function is called by presentation layer which passes a
'listbox and collection object that holds the listbox info
'Filter is used for showing filter lists of contacts for a company
'
'The return value is 0 for any fault and 1 for pass
'==========================================================

Dim sQuery As String
mvarStatus = 20 'GetLbData started

On Error GoTo GetDataError

'MsgBox iDataIndex, vbInformation, "BLServer"

If Not ((iDataIndex = clLINK) Or (iDataIndex = clCOMPANYUPDATE) Or (iDataIndex = clCONTACTUPDATE)) Then 'get links
    sQuery = MakeQuery(iDataIndex)
Else
    sQuery = MakeQuery(iDataIndex, Filter)
End If

'Pass string and recordset to DBLogic module
If DBServer.GetListRS(sQuery, rsLBData, iDataIndex) Then
    Set GetLbData = rsLBData
    If GetLbData Is Nothing Then
        MsgBox "GetLBData not created", vbCritical, "BLServer - GetLBData"
        GoTo GetDataError
    End If
End If

mvarStatus = 25 'GetLbData started
On Error GoTo 0
Exit Function

GetDataError:
mvarStatus = 90 'GetLbData started
MsgBox "Error: " & mvarStatus, vbCritical, "BLServer - GetLBData"
MsgBox Err.Number & vbCrLf & Err.Description, , "GetCustomers"
End Function
Public Function AddCall(iCompanyID As Long, iContactID As Long, _
iCallCodeID As Long, iProductID As Long, iEmplID As Long, _
dNoteDate As Date, sNote As String, iCallTime As Integer, iCaseID As Long) As Long
'========================================================================
'This function gathers call information from the form
'Then accesses the DB layer to append the call
'========================================================================
Dim dEntryDate As Date
Dim sTable As String
Dim NewCalls As CCalls
Dim rsTemp As ADODB.Recordset

AddCall = 0 'assume failure

dEntryDate = Now()

'-------------------------------------------------------------------------
'Validate call data
If iCallCodeID = 0 Then Exit Function
If iCallTime = 0 Then Exit Function
If iCaseID = 0 Then iCaseID = 0
If iCompanyID = 0 Then Exit Function
If iContactID = 0 Then Exit Function
If iEmplID = 0 Then Exit Function
If sNote = "" Then Exit Function
If dNoteDate = vbNull Then dNoteDate = Now()
If iProductID = 0 Then Exit Function
'-------------------------------------------------------------------------
'MsgBox "iCompanyID = " & iCompanyID & vbCrLf & "iContactID = " & iContactID & vbCrLf & _
"iCallCodeID = " & iCallCodeID & vbCrLf & "iProductID = " & iProductID & vbCrLf & _
"iEmplID = " & iEmplID & vbCrLf & "dNoteDate = " & dNoteDate & vbCrLf & _
"sNote = " & sNote & vbCrLf & "iCallTime = " & iCallTime & vbCrLf & _
"iCaseID = " & iCaseID, , "BL::AddCall - 1"

Set rsTemp = CreateRS(clCALL)

rsTemp.AddNew
    rsTemp!ID = 0
    rsTemp!CallCodeId = iCallCodeID
    rsTemp!CallTime = iCallTime
    rsTemp!CaseID = iCaseID
    rsTemp!CompanyID = iCompanyID
    rsTemp!ContactID = iContactID
    rsTemp!EmplID = iEmplID
    rsTemp!EntryDate = dEntryDate
    rsTemp!ProductID = iProductID
    rsTemp!Note = sNote
    rsTemp!NoteDate = dNoteDate
rsTemp.Update

sTable = "SupportCalls"

AddCall = DBServer.AppendData(sTable, rsTemp)

If AddCall = 0 Then
    MsgBox "Failed to add call", vbExclamation, "BLogic::AddCall - Error"
    Exit Function
End If

End Function
Public Function AddCustomer(CompanyName As String, Address As String, City As String, State As String, Zip As String, Country As String, Phone As String, Fax As String, iEmplID As Long, cType As String, Optional ID As Long) As Long
Dim sTable As String
Dim rsTemp As ADODB.Recordset
Dim iTemp As Long

AddCustomer = 0

'-------------------------------------------------
'Validate field data
If CompanyName = "" Then Exit Function
If iEmplID = 0 Then Exit Function
If Phone = "" Then Exit Function

If Address = "" Then Address = " "
If City = "" Then City = " "
If State = "" Then State = " "
If Zip = "" Then Zip = " "
If Country = "" Then Country = " "
If Fax = "" Then Fax = " "
If cType = "" Then cType = "U"
'-------------------------------------------------

Set rsTemp = CreateRS(clCOMPANY)

rsTemp.AddNew
    If ID <> 0 Then rsTemp!ID = ID Else rsTemp!ID = 0
    rsTemp!sName = CompanyName
    rsTemp!Address = Address
    rsTemp!City = City
    rsTemp!State = State
    rsTemp!Zip = Zip
    rsTemp!Country = Country
    rsTemp!Phone = Phone
    rsTemp!Fax = Fax
    rsTemp!Type = cType
    rsTemp!EmplID = iEmplID
rsTemp.Update

sTable = "Company"

AddCustomer = DBServer.AppendData(sTable, rsTemp)

If AddCustomer = 0 Then
    MsgBox "Failed to add customer", vbInformation, "BLogic - AddCustomer"
    Exit Function
End If

End Function
Public Function AddContact(FirstName As String, LastName As String, Title As String, iEmplID As Long, Phone As String, Skill As String, Training As String, Optional ID As Long) As Long
Dim sTable As String
Dim rsTemp As ADODB.Recordset
Dim iTemp As Long

AddContact = 0

'-------------------------------------------------
'Validate field data
If iEmplID = 0 Then Exit Function
If FirstName = "" Then Exit Function

If LastName = "" Then LastName = " "
If Title = "" Then Title = " "
If Phone = "" Then Phone = " "
If Skill = "" Then Skill = "1"
If Training = "" Then Training = "N"
'-------------------------------------------------

Set rsTemp = CreateRS(clCONTACT)

rsTemp.AddNew
    If ID <> 0 Then rsTemp!ID = ID Else rsTemp!ID = 0
    rsTemp!FirstName = FirstName
    rsTemp!LastName = LastName
    rsTemp!Title = Title
    rsTemp!EmplID = iEmplID
    rsTemp!Phone = Phone
    rsTemp!Skill = Skill
    rsTemp!Training = Training
rsTemp.Update

sTable = "Contact"

AddContact = DBServer.AppendData(sTable, rsTemp)

If AddContact = 0 Then
    MsgBox "Failed to add contact", vbInformation, "BL - AddContact"
    Exit Function
End If

Set rsTemp = Nothing

End Function
Public Function AddCompanyLink(iCompanyID As Long, iContactID As Long, iEmplID As Long) As Long
Dim sTable As String
Dim rsTemp As ADODB.Recordset
Dim iTemp As Long

On Error GoTo ErrorHandler
AddCompanyLink = 0

'-------------------------------------------------
'Validate field data
If iCompanyID = 0 Then Exit Function
If iContactID = 0 Then Exit Function
If iEmplID = 0 Then Exit Function
'-------------------------------------------------
'If Not rsTemp = Nothing Then
'    Set rsTemp = Nothing
'End If

Set rsTemp = CreateRS(clLINK)

rsTemp.AddNew
    rsTemp!CompanyID = iCompanyID
    rsTemp!ContactID = iContactID
    rsTemp!EmplID = iEmplID
    rsTemp!ID = 0
rsTemp.Update

sTable = "[Contact Link]"

AddCompanyLink = DBServer.AppendData(sTable, rsTemp)

'AddCompanyLink = DBServer.AppendCompanyLink(iCompany, iContact, iEmplID)
If AddCompanyLink = 0 Then
    MsgBox "Failed to add link between contact and company", vbInformation, "BL::AddCompanyLink"
    Exit Function
End If

Set rsTemp = Nothing

Exit Function

ErrorHandler:
MsgBox "Error entering Link;" & vbCrLf & "Err: " & Err.Number & vbCrLf & "Error Description: " & Err.Description, vbExclamation, "AddCompanyLink"
Resume Next

End Function
Public Function AddCallCode(sCallType As String, iEmplID As Long) As Long
Dim sTable As String
Dim rsTemp As ADODB.Recordset
Dim iTemp As Long

AddCallCode = 0

'-------------------------------------------------
'Validate field data
If sCallType = "" Then Exit Function
If iEmplID = 0 Then Exit Function
'-------------------------------------------------
Set rsTemp = CreateRS(clCALLCODE)

rsTemp.AddNew
    rsTemp!CallType = sCallType
    rsTemp!EmplID = iEmplID
rsTemp.Update

sTable = "CallCode"

'AddCallCode = DBServer.AppendCallCode(iCallCode, iEmplID)
AddCallCode = DBServer.AppendData(sTable, rsTemp)

If AddCallCode = 0 Then
    MsgBox "Failed to add new call code.", vbInformation, "BLogic - AddCallCode"
    Exit Function
End If

End Function
Public Function GetHistory(iID As Long) As ADODB.Recordset
Dim rsData As ADODB.Recordset

'Pass string and recordset to DBLogic module
If DBServer.GetHistoryRS(MakeQuery(clHISTORY, iID), rsData) Then
    Set GetHistory = rsData
    If GetHistory Is Nothing Then
        MsgBox "GetHistory not created", vbCritical, "BLServer - GetHistory"
        GoTo GetDataError
    End If
End If

On Error GoTo 0
Exit Function

GetDataError:
Set rsData = Nothing
End Function
Public Function GetProductHistory(iID As Long) As ADODB.Recordset
Dim rsData As ADODB.Recordset

'Pass string and recordset to DBLogic module
If DBServer.GetHistoryRS(MakeQuery(clPRODUCTHISTORY, iID), rsData) Then
    Set GetProductHistory = rsData
    If GetProductHistory Is Nothing Then
        MsgBox "GetProductHistory not created", vbCritical, "BLServer - GetProductHistory"
        GoTo GetDataError
    End If
End If

On Error GoTo 0
Exit Function

GetDataError:
Set rsData = Nothing
End Function
'===================================================================
Public Function AddEntries(oCol As Object) As Integer
'This function gathers call information from the form
'Then it accesses the DB layer to append the call
Dim sTable As String

AddEntries = 0

sTable = "SupportCalls"

If DBServer.AppendEntries(sTable, oCol) = 0 Then MsgBox "Failed to add entries", vbExclamation, "AddEntries Error"

AddEntries = 1
End Function

Private Sub Class_Initialize()
mvarStatus = 99 'Not started
If (DBServer Is Nothing) Then
    Set DBServer = CreateObject("CallTrackerDBServer.DBServer")
    mvarStatus = 80 'DB Server started
End If
End Sub
Private Sub Class_Terminate()
If Not (DBServer Is Nothing) Then
    Set DBServer = Nothing
    mvarStatus = 99 'DB Server stopped
End If
End Sub
Private Function CreateRS(iIndex As Integer) As ADODB.Recordset
Dim rs As ADODB.Recordset
'========================================================================
'This function creates an ADO recordset used to pass contact/customer data to
'the DBServer.
'iIndex selects the proper rs.
'========================================================================
Set rs = New ADODB.Recordset
With rs.Fields
    .Append "ID", adBigInt '
    .Append "Address", adChar, 100
    .Append "CallCodeID", adBigInt '
    .Append "CallType", adChar, 50
    .Append "CallTime", adInteger '
    .Append "CaseID", adBigInt '
    .Append "CompanyID", adBigInt '
    .Append "ContactID", adBigInt '
    .Append "City", adChar, 50
    .Append "Country", adChar, 50
    .Append "DateEntered", adDBTimeStamp
    .Append "EmplID", adBigInt '
    .Append "EntryDate", adDBTimeStamp '
    .Append "Fax", adChar, 50
    .Append "FirstName", adChar, 50
    .Append "LastName", adChar, 50
    .Append "NoteDate", adDBDate '
    .Append "Note", adChar, 500 '
    .Append "Phone", adChar, 100
    .Append "ProductID", adBigInt '
    .Append "Skill", adChar, 50
    .Append "sName", adChar, 200
    .Append "State", adChar, 2
    .Append "Title", adChar, 50
    .Append "Training", adChar, 50
    .Append "Type", adChar, 8
    .Append "Zip", adChar, 15
End With

rs.Open

Set CreateRS = rs
Set rs = Nothing

End Function
