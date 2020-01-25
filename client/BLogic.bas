Attribute VB_Name = "BLogic"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EBC8650381"
Option Explicit
Dim RS As Recordset
Dim Employees As New Collection
Dim Customers As New CCustomers
Dim Contacts As New Collection
Dim Products As New Collection
Dim CallCodes As New Collection
Dim Entries As New Collection
Dim Entry As CEntry

Const NOCUSTOMER As Long = 6
Const NOCONTACT As Long = 10
Const NOPRODUCT As Long = 14
Const NOCODE As Long = 6

Dim rsData As adodb.Recordset
Function GetData(lb As ListBox, oData As Object, Optional Filter As Long) As Integer
Dim sQuery As String

On Error GoTo GetDataError
GetData = 0 ' assume failure

If Not (TypeOf oData Is CLinks) Then
    sQuery = MakeQuery(lb.Index, Filter)
Else
    sQuery = MakeQuery(4)
End If

'Pass string and recordset to DBLogic module
If GetListRS(sQuery, oData) = 0 Then MsgBox "Error getting customers"

'Load Customer listbox
If Not (TypeOf oData Is CLinks) Then
    If LoadListBox(lb, oData) = 0 Then MsgBox "Couldn't load listbox", , "BLogic-GetCustomers"
End If

GetData = 1
Exit Function

GetDataError:
MsgBox Err.Number & vbCrLf & Err.Description, , "GetCustomers"
GetData = 0
End Function
'##ModelId=39EBC86701D6
Function GetCustomers2(lb As ListBox, oData As Object, Optional Filter As Long) As Integer
Dim sQuery As String
'Dim sFilter As String

On Error GoTo GetCustomersError
GetCustomers2 = 0 ' assume failure

sQuery = MakeQuery(lb.Index)

'Pass string and recordset to DBLogic module
If GetListRS(sQuery, oData) = 0 Then MsgBox "Error getting customers"

'Load Customer listbox
If Filter = 0 Then
    If LoadListBox(lb, oData) = 0 Then MsgBox "Couldn't load listbox", , "BLogic-GetCustomers"
End If

GetCustomers2 = 1
Exit Function

GetCustomersError:
MsgBox Err.Number & vbCrLf & Err.Description, , "GetCustomers"
GetCustomers2 = 0
End Function
'##ModelId=39EBC867035C
Function GetContacts2(lb As ListBox, oData As Object, Optional Filter As Long) As Integer
Dim sQuery As String

On Error GoTo ErrorHandler
GetContacts2 = 0 ' assume failure

'Make sql select string
sQuery = MakeQuery(lb.Index)

'Pass string and recordset to DBLogic module
If GetListRS(sQuery, oData) = 0 Then MsgBox "Error getting customers"

'Load Customer listbox
If Filter = 0 Then
    If LoadListBox(lb, oData) = 0 Then MsgBox "Couldn't load listbox", , "BLogic-GetCustomers"
End If
GetContacts2 = 1
Exit Function

ErrorHandler:
MsgBox Err.Number & vbCrLf & Err.Description, , "GetContacts"
GetContacts2 = 0
End Function
'##ModelId=39EBC86800F1
Function GetCallCodes2(lb As ListBox, colData As Object, Optional Filter As Long) As Integer
Dim sQuery As String
Dim sSelect As String
Dim sFrom As String
Dim sWhere As String
Dim sOrderBy As String
Dim sFilter As String

On Error GoTo ErrorHandler
GetCallCodes2 = 0 ' assume failure

'Make sql select string

If Filter <> 0 Then
    sFilter = " AND ([CallCode]![ID]=" & CStr(Filter) & ")"
Else
    sFilter = ""
End If

'Make sql select string
sSelect = "SELECT [CallCode]![CallType] AS sName, CallCode.ID "
sFrom = "From CallCode "
sWhere = "WHERE (([CallCode]![ID]<>6)" & sFilter & ") "
sOrderBy = "ORDER BY [CallCode]![CallType];"

sQuery = sSelect & sFrom & sWhere & sOrderBy

'Pass string and recordset to DBLogic module
If GetListRS(sQuery, colData) = 0 Then MsgBox "Error getting customers"

'Load Callcode listbox
If Filter = 0 Then
    If LoadListBox(lb, colData) = 0 Then MsgBox "Couldn't load listbox", , "BLogic-GetCustomers"
End If

GetCallCodes2 = 1
Exit Function

ErrorHandler:
MsgBox Err.Number & vbCrLf & Err.Description, , "GetContacts"
GetCallCodes2 = 0
End Function
'##ModelId=39EBC8680277
Function GetProducts2(lb As ListBox, colData As Object) As Integer
Dim sQuery As String

On Error GoTo ErrorHandler
GetProducts2 = 0 ' assume failure

'Make sql select string
sQuery = "SELECT [Product]![ProductName] AS sName, Product.ID "
sQuery = sQuery & "From Product "
sQuery = sQuery & "WHERE ([Product]![ID]<>14) "
sQuery = sQuery & "ORDER BY [Product]![ProductName];"

'Pass string and recordset to DBLogic module
If GetListRS(sQuery, colData) = 0 Then MsgBox "Error getting customers"

'Load Customer listbox
If LoadListBox(lb, colData) = 0 Then MsgBox "Couldn't load listbox", , "BLogic-GetCustomers"

GetProducts2 = 1
Exit Function

ErrorHandler:
MsgBox Err.Number & vbCrLf & Err.Description, , "GetContacts"
GetProducts2 = 0
End Function

'##ModelId=39EBC868039A
Function GetLinks(colData As Object) As Integer
Dim sQuery As String
Dim sSelect As String
Dim sFrom As String
Dim sWhere As String
Dim sOrderBy As String

On Error GoTo ErrorHandler
GetLinks = 0 ' assume failure

'Pass string and recordset to DBLogic module
sQuery = MakeQuery(4)
'If GetLbData(lstItem(2).Index, colData, 1) = 0 Then MsgBox "Error getting contacts"

GetLinks = 1
Exit Function

ErrorHandler:
MsgBox Err.Number & vbCrLf & Err.Description, , "GetContacts"
GetLinks = 0
End Function

'##ModelId=39EBC8690066
Function AddCall(iCustomerID As Long, iContactID As Long, iCallCodeID As Long, iProductID As Long, iEmplID As Long, dNoteDate As Date, sNote As String, iCallTime As Integer) As Integer
'This function gathers call information from the form
'Then it accesses the DB layer to append the call
Dim dEntryDate As Date
Dim sTable As String

AddCall = 0

sTable = "SupportCalls"
dEntryDate = Now()

'Validate call data
If (iCustomerID < 1) Or (iContactID < 1) Or (sNote = "") Then
    MsgBox "Please select the customer, contact, product, code, and enter note.", vbCritical, "BLogic - AddCall"
    AddCall = 0
    Exit Function
End If

If AppendCall(sTable, iCustomerID, iContactID, iCallCodeID, iProductID, iEmplID, dNoteDate, sNote, dEntryDate, iCallTime) = 0 Then MsgBox "Failed to add record", vbExclamation, "AddCall Error"

AddCall = 1
End Function
'##ModelId=39EBC86A003F
Function AddCustomer(CompanyName As String, Address As String, City As String, State As String, Zip As String, Country As String, Phone As String, Fax As String, iEmplID As Long) As Long
Dim iTemp As Long

AddCustomer = 0

'-------------------------------------------------
'Validate field data
If CompanyName = "" Then Exit Function
If Address = "" Then Address = " "
If City = "" Then Exit Function
If State = "" Then Exit Function
If Zip = "" Then Zip = " "
If Country = "" Then Country = " "
If Phone = "" Then Exit Function
If Fax = "" Then Fax = " "
If iEmplID = 0 Then Exit Function
'-------------------------------------------------
AddCustomer = AppendCustomer(CompanyName, Address, City, State, Zip, Country, Phone, Fax, iEmplID)

If AddCustomer = 0 Then
    MsgBox "Failed to add customer", vbInformation, "BLogic - AddCustomer"
    Exit Function
End If

'AddCustomer = iTemp

End Function
'##ModelId=39EBC86B00A5
Function AddContact(FirstName As String, LastName As String, Title As String, iEmplID As Long) As Long
Dim iTemp As Long

AddContact = 0

'-------------------------------------------------
'Validate field data
If FirstName = "" Then Exit Function
If LastName = "" Then Exit Function
If Title = "" Then Title = " "
If iEmplID = 0 Then Exit Function
'-------------------------------------------------
AddContact = AppendContact(FirstName, LastName, Title, iEmplID)
If AddContact = 0 Then
    MsgBox "Failed to add contact", vbInformation, "BLogic - AddContact"
    Exit Function
End If

'AddContact = iTemp

End Function
'##ModelId=39EBC86B02C2
Function AddCompanyLink(iCompany As Long, iContact As Long, EmplID As Long) As Long
Dim iTemp As Long

On Error GoTo ErrorHandler

AddCompanyLink = 0

'-------------------------------------------------
'Validate field data
If iCompany = 0 Then Exit Function
If iContact = 0 Then Exit Function
If EmplID = 0 Then Exit Function
'-------------------------------------------------
AddCompanyLink = AppendCompanyLink(iCompany, iContact, EmplID)
If AddCompanyLink = 0 Then
    MsgBox "Failed to add link between contact and company", vbInformation, "BLogic - AddCompanyLink"
    Exit Function
End If
Exit Function
ErrorHandler:
MsgBox "Error entering Link; Err: " & Err.Number & vbCrLf & "Error Description: " & Err.Description, vbExclamation, "AddCompanyLink"
Resume Next

End Function

'##ModelId=39EBC86C0088
Function AddCallCode(CallCode As String, iEmplID As Long) As Long
Dim iTemp As Long

AddCallCode = 0

'-------------------------------------------------
'Validate field data
If CallCode = "" Then Exit Function
If iEmplID = 0 Then Exit Function
'-------------------------------------------------
AddCallCode = AppendCallCode(CallCode, iEmplID)
If AddCallCode = 0 Then
    MsgBox "Failed to add new call code.", vbInformation, "BLogic - AddContact"
    Exit Function
End If

'AddCallCode = iTemp

End Function
'##ModelId=39EBC86C01BF
Function GetHistory(iID As Long, oCalls As Object)
'Dim sQry As String
'Dim sSelect As String
'Dim sFrom As String
'Dim sWhere As String
'Dim sOrderBy As String
'
'GetHistory = 0
'
''get the customer call history
'    sSelect = "SELECT Employees.LastName AS LastName, "
'    sSelect = sSelect & "[Contact]![FirstName] & ' ' & "
'    sSelect = sSelect & "[Contact]![LastName] & ', ' & "
'    sSelect = sSelect & "[Contact]![Title] AS ContactName, "
'    sSelect = sSelect & "Company.CompanyName AS CompanyName, "
'    sSelect = sSelect & "Product.ProductName AS ProductName, "
'    sSelect = sSelect & "CallCode.CallType AS CallType, "
'    sSelect = sSelect & "SupportCalls.NoteDate AS NoteDate, "
'    sSelect = sSelect & "SupportCalls.Note AS sNote, "
'    sSelect = sSelect & "SupportCalls.CallTime AS iCallTime "
'
'    sFrom = sFrom & "FROM Product INNER JOIN "
'    sFrom = sFrom & "(Employees INNER JOIN "
'    sFrom = sFrom & "(Contact INNER JOIN "
'    sFrom = sFrom & "(Company INNER JOIN "
'    sFrom = sFrom & "(CallCode INNER JOIN SupportCalls "
'    sFrom = sFrom & "ON CallCode.ID = SupportCalls.CallCodeID) "
'    sFrom = sFrom & "ON Company.ID = SupportCalls.CustomerID) "
'    sFrom = sFrom & "ON Contact.ID = SupportCalls.ContactID) "
'    sFrom = sFrom & "ON Employees.ID = SupportCalls.EmployeeID) "
'    sFrom = sFrom & "ON Product.ID = SupportCalls.ProductID "
'
'    sWhere = sWhere & "Where (((Company.ID) = " & iID & ")) "
'
'    sOrderBy = sOrderBy & "ORDER BY SupportCalls.NoteDate DESC;"
'
'    sQry = sSelect & sFrom & sWhere & sOrderBy
'
'    If GetHistoryRS(sQry, oCalls) = 0 Then MsgBox "Error getting call history", vbCritical, "BLogic - GetHistory"
GetHistory = 1
End Function
'##ModelId=39EBC86C02F5
Function GetEmployees(cboBox As ComboBox, oData As Object) As Integer
'Dim sQuery As String
'
'On Error GoTo ErrorHandler
'GetEmployees = 0 ' assume failure
'
''Make sql select string
'sQuery = "SELECT Employees.ID AS ID, "
'sQuery = sQuery & "IIf(Len([Employees]![FirstName])=1,[Employees]![MiddleName],[Employees]![FirstName]) & ' ' & [Employees]![LastName] AS sName "
'sQuery = sQuery & "FROM Employees "
'sQuery = sQuery & "ORDER BY Employees.LastName;"
'
''Pass string and recordset to DBLogic module
'If GetEmployeeRS(sQuery, oData) = 0 Then MsgBox "Error getting customers"
'
''Load employees combobox
'If LoadComboBox(cboBox, oData) = 0 Then MsgBox "Couldn't load combobox", , "BLogic-GetEmployees"
'
'GetEmployees = 1
'Exit Function
'
'ErrorHandler:
'MsgBox Err.Number & vbCrLf & Err.Description, , "GetEmployees"
GetEmployees = 0
End Function
'##ModelId=39EBC86D0044
Function AddEntries(oCol As Object) As Integer
'This function gathers call information from the form
'Then it accesses the DB layer to append the call
'Dim sTable As String
'
'AddEntries = 0
'
'sTable = "Calls"
'sTable = "SupportCalls"
'
'If AppendEntries(sTable, oCol) = 0 Then MsgBox "Failed to add entries", vbExclamation, "AddEntries Error"
'
AddEntries = 1
End Function
