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
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B4FC3840164"
Option Explicit

Const NOCUSTOMER As Long = 6
Const NOCONTACT As Long = 10
Const NOPRODUCT As Long = 14
Const NOCODE As Long = 6

Const clCOMPANYBYID As Integer = 0 'Version 3 'OK
'Const clCOMPANYBYDATE As Integer = 0
'Const clCONTACT As Integer = 1 'Version 3
Const clCONTACTBYID As Integer = 1 'Version 3
Const clPRODUCT As Integer = 2 'Version 3 'OK
Const clCALLCODE As Integer = 3 'Version 3 'OK
Const clCALL As Integer = 4 'Version 3 'unused
Const clHISTORY2 As Integer = 4
Const clLINK As Integer = 5 'Version 3 'OK
Const clHISTORY As Integer = 6 'Version 3
Const clCALLTYPE As Integer = 7
Const clEMPLOYEE As Integer = 8 'Version 3 'OK
Const clCOMPANYUPDATE As Integer = 9 'Version 3 'OK
Const clCONTACTUPDATE As Integer = 10 'Version 3 'OK
Const clPRODUCTHISTORY As Integer = 11
Const clCONTACT2 As Integer = 12 'Version 3
Const clNEWCONTACTBYID As Integer = 12 'Version 3
Const clNEWCOMPANYBYID As Integer = 13 'Version 3
Const clCOMPANYBYDATE As Integer = 13 'Version 3
'Const clCOMPANYBYID As Integer = 13
Const clCONTACT3 As Integer = 14 'Version 3
Const clCONTACTBYCOMPID As Integer = 14 'Version 3
Const clCONTACT4 As Integer = 22
Const clCASEID As Integer = 15 'Version 3
Const clCOMPANY4 As Integer = 16
Const clNEWCALLS As Integer = 17
Const clNEWCOMPANIES As Integer = 18
Const clNEWCOMPANIESNOCALLS As Integer = 19
Const clREFRESHCOMPANIES As Integer = 20
Const clREFRESHCONTACTS As Integer = 21

Dim rsData As ADODB.Recordset
Dim rsLBData As ADODB.Recordset

Dim WithEvents DBServer As CallTrackerDBServer.DBServer
Attribute DBServer.VB_VarHelpID = -1

'local variable(s) to hold property value(s)
Private mvarStatus As Integer 'local copy

Public Event OnUpdateDone()
Public Event OnNewCustomer()

Public Property Get Status() As Integer
    Status = mvarStatus
End Property
Private Function MakeQuery(Index As Integer, Optional Filter As Long) As String
Dim sQuery As String
Dim sSelect As String
Dim sFrom As String
Dim sWhere As String
Dim sOrderBy As String
Dim sFilter As String

On Error GoTo MakeQueryErrorHandler
MakeQuery = "" ' assume failure

Select Case Index
    Case 0
    'This query selects SUMMARY company info for a treeview
    'It can select all companies or a specific company
        If Filter <> 0 Then
            sFilter = " AND ([Company]![ID]=" & CStr(Filter) & ")"
        Else
            sFilter = ""
        End If
        
        sSelect = "SELECT IIf([Type]='D',1,IIf([Type]='E',2,IIf([Type]='F',3, "
        sSelect = sSelect & "IIf([Type]='I',4,IIf([Type]='O',5,IIf([Type]='R',6, "
        sSelect = sSelect & "IIf([Type]='S',7,IIf([Type]='U',8,8)))))))) AS cType, "
        sSelect = sSelect & "'p' & cstr([Company]![ID]) AS ParentID, "
        sSelect = sSelect & "'p' & cstr([Company]![ID]) AS ChildID, "
        sSelect = sSelect & "[Company]![CompanyName] & ', ' & "
        sSelect = sSelect & "[Company]![Address] & ', ' & [Company]![City] & ', ' & "
        sSelect = sSelect & "[Company]![State] AS sName, [Company]![CompanyName] & ', ' & "
        sSelect = sSelect & "[Company]![Address] & ', ' & [Company]![City] & ', ' & "
        sSelect = sSelect & "[Company]![State] AS SortKey "
        
        sFrom = "FROM Company "
        
        sWhere = "WHERE ((('p' & [Company]![ID])<>('p' & " & NOCUSTOMER & "))" & sFilter & ") "
        
        sOrderBy = "ORDER BY "
        sOrderBy = sOrderBy & "[Company]![CompanyName], [Company]![State], "
        sOrderBy = sOrderBy & "[Company]![City], [Company]![Address]; "
    Case 1 'working
    'This query selects SUMMARY CONTACT info for a treeview
    'It can select all CONTACTS or a specific CONTACT
        If Filter <> 0 Then
            sFilter = " AND ([Contact]![ID]=" & CStr(Filter) & ")"
        Else
            sFilter = ""
        End If
        
        sSelect = "SELECT IIf(InStr(1,[Training],'N'),18, "
        sSelect = sSelect & "IIf(InStr(1,[Training],'C'),15, "
        sSelect = sSelect & "IIf(InStr(1,[Training],'L') Or "
        sSelect = sSelect & "InStr(1,[Training],'R'),12, "
        sSelect = sSelect & "IIf(InStr(1,[Training],'A') Or "
        sSelect = sSelect & "InStr(1,[Training],'D') Or "
        sSelect = sSelect & "InStr(1,[Training],'S'),9,0))))+[Skill]-1 AS cType, "
        sSelect = sSelect & "'p' & cstr([Contact Link]![CompanyID]) AS ParentID, "
        sSelect = sSelect & "IIf([Contact]![ID] Is Null,'','c' & "
        sSelect = sSelect & "cstr([Contact]![ID])) AS ChildID, "
        sSelect = sSelect & "[Contact]![FirstName] & ' ' & [Contact]![LastName] & ', ' & "
        sSelect = sSelect & "[Contact]![Title] & ' ' & [Contact]![Phone] AS sName, "
        sSelect = sSelect & "[Company]![CompanyName] & ', ' & [Company]![Address] & ', ' & "
        sSelect = sSelect & "[Company]![City] & ', ' & [Company]![State] & ', ' & "
        sSelect = sSelect & "[Company]![ZIP] & ', ' & [Contact]![FirstName] & ' ' & "
        sSelect = sSelect & "[Contact]![LastName] & ', ' & [Contact]![Title] & ' ' & "
        sSelect = sSelect & "[Contact]![Phone] AS SortKey "
        
        sFrom = "FROM Company INNER JOIN "
        sFrom = sFrom & "(Contact RIGHT JOIN [Contact Link] "
        sFrom = sFrom & "ON Contact.ID = [Contact Link].ContactID) "
        sFrom = sFrom & "ON Company.ID = [Contact Link].CompanyID "
        
'        sWhere = "WHERE ((('c' & [Contact]![ID])<>('c' & " & NOCONTACT & "))" & sFilter & ") "
        sWhere = "WHERE ((([Contact]![ID])<>(" & NOCONTACT & "))" & sFilter & ") "
        
        sOrderBy = "ORDER BY [Company]![CompanyName], [Company]![State], "
        sOrderBy = sOrderBy & "[Company]![City], [Company]![Address], "
        sOrderBy = sOrderBy & "[Company]![ZIP], [Contact]![FirstName], "
        sOrderBy = sOrderBy & "[Contact]![LastName], [Contact]![Title], "
        sOrderBy = sOrderBy & "[Contact]![Phone]; "
    Case clPRODUCT 'working
    'This query selects PRODUCTS for a listview
        sSelect = "SELECT [Product]![ProductName] AS sName, Product.ID "
        sFrom = "From Product "
        sWhere = "WHERE ([Product]![ID]<>" & NOPRODUCT & ") "
        sOrderBy = "ORDER BY [Product]![ProductName];"
    Case clCALLCODE 'working
    'This query selects CALLCODES for a listview
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
    'This query selects ALL CONTACT LINK data
        sSelect = "SELECT [Contact Link].* "
        sFrom = "FROM [Contact Link];"
        sWhere = ""
        sOrderBy = ""
    Case clEMPLOYEE 'working
    'This query selects EMPLOYEES for a listview
        sSelect = "SELECT Employees.ID AS ID, "
        sSelect = sSelect & "IIf(Len([Employees]![FirstName])=1,[Employees]![MiddleName],[Employees]![FirstName]) & ' ' & [Employees]![LastName] AS sName "
        sFrom = "FROM Employees "
        sWhere = ""
        sOrderBy = "ORDER BY Employees.LastName;"
    Case clHISTORY 'working
    'This query selects the CALL HISTORY for one company
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
    Case 9 'get the company data fields
    'This query selects ALL COMPANY info for editing company info
    'It can select one or all companies
        If Filter <> 0 Then
            sFilter = " AND ([Company]![ID]=" & CStr(Filter) & ")"
        Else
            sFilter = ""
        End If
        
        sSelect = "SELECT * "
        sFrom = "FROM Company "
        sWhere = "WHERE (([Company]![ID]<>" & NOCUSTOMER & ")" & sFilter & ") "
        sOrderBy = "ORDER BY [Company]![CompanyName], [Company]![State], [Company]![City];"
    Case 10
    'This query selects ALL CONTACT info for editing Contact info
    'It can select one or all Contacts
        If Filter <> 0 Then
            sFilter = " ([Contact]![ID]=" & CStr(Filter) & ")"
'            sFilter = " AND ([Contact]![ID]=" & CStr(Filter) & ")"
        Else
            sFilter = ""
        End If
        
        sSelect = "SELECT * "
        sFrom = "FROM Contact "
        sWhere = "WHERE (" & sFilter & ") "
'        sWhere = "WHERE (([Contact]![ID]<>" & NOCONTACT & ")" & sFilter & ") "
        sOrderBy = "ORDER BY [Contact]![FirstName] & ' ' & [Contact]![LastName] & ', ' & [Contact]![Title];"
    Case clPRODUCTHISTORY
    'This query selects PRODUCT CALL HISTORY info for a chosen product
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
    Case clCOMPANYBYDATE
    'This query selects SUMMARY company info for a treeview
    'It can select all companies or a specific company
    'This query is EXACTLY like clCOMPANY
        If Filter <> 0 Then
            sFilter = " AND ([Company]![ID]=" & CStr(Filter) & ")"
        Else
            sFilter = ""
        End If
        
        sSelect = "SELECT IIf([Type]='D',1,IIf([Type]='E',2,IIf([Type]='F',3, "
        sSelect = sSelect & "IIf([Type]='I',4,IIf([Type]='O',5,IIf([Type]='R',6, "
        sSelect = sSelect & "IIf([Type]='S',7,IIf([Type]='U',8,8)))))))) AS cType, "
        sSelect = sSelect & "'p' & cstr([Company]![ID]) AS ParentID, "
        sSelect = sSelect & "'p' & cstr([Company]![ID]) AS ChildID, "
        sSelect = sSelect & "[Company]![CompanyName] & ', ' & "
        sSelect = sSelect & "[Company]![Address] & ', ' & [Company]![City] & ', ' & "
        sSelect = sSelect & "[Company]![State] AS sName, [Company]![CompanyName] & ', ' & "
        sSelect = sSelect & "[Company]![Address] & ', ' & [Company]![City] & ', ' & "
        sSelect = sSelect & "[Company]![State] AS SortKey "
        
        sFrom = "FROM Company "
        
        sWhere = "WHERE ((('p' & [Company]![ID])<>('p' & " & NOCUSTOMER & "))" & sFilter & ") "
        
        sOrderBy = "ORDER BY "
        sOrderBy = sOrderBy & "[Company]![CompanyName] & ', ' & [Company]![Address] & ', ' & "
        sOrderBy = sOrderBy & "[Company]![City] & ', ' & [Company]![State]; "
    Case clCONTACT2
    'This query selects SUMMARY CONTACT info for a treeview
    'It can select all CONTACTS or a specific CONTACT
    'This query is EXACTLY like clCOMPANY except for 'c' prefix on ID
        If Filter <> 0 Then
            sFilter = " AND ([Contact]![ID]=" & CStr(Filter) & ")"
        Else
            sFilter = ""
        End If
        
        sSelect = "SELECT IIf(InStr(1,[Training],'N'),18, "
        sSelect = sSelect & "IIf(InStr(1,[Training],'C'),15, "
        sSelect = sSelect & "IIf(InStr(1,[Training],'L') Or "
        sSelect = sSelect & "InStr(1,[Training],'R'),12, "
        sSelect = sSelect & "IIf(InStr(1,[Training],'A') Or "
        sSelect = sSelect & "InStr(1,[Training],'D') Or "
        sSelect = sSelect & "InStr(1,[Training],'S'),9,0))))+[Skill]-1 AS cType, "
        sSelect = sSelect & "'p' & cstr([Contact Link]![CompanyID]) AS ParentID, "
        sSelect = sSelect & "IIf([Contact]![ID] Is Null,'','c' & "
        sSelect = sSelect & "cstr([Contact]![ID])) AS ChildID, "
        sSelect = sSelect & "[Contact]![FirstName] & ' ' & [Contact]![LastName] & ', ' & "
        sSelect = sSelect & "[Contact]![Title] & ' ' & [Contact]![Phone] AS sName, "
        sSelect = sSelect & "[Company]![CompanyName] & ', ' & [Company]![Address] & ', ' & "
        sSelect = sSelect & "[Company]![City] & ', ' & [Company]![State] & ', ' & "
        sSelect = sSelect & "[Company]![ZIP] & ', ' & [Contact]![FirstName] & ' ' & "
        sSelect = sSelect & "[Contact]![LastName] & ', ' & [Contact]![Title] & ' ' & "
        sSelect = sSelect & "[Contact]![Phone] AS SortKey "
        
        sFrom = "FROM Company INNER JOIN "
        sFrom = sFrom & "(Contact RIGHT JOIN [Contact Link] "
        sFrom = sFrom & "ON Contact.ID = [Contact Link].ContactID) "
        sFrom = sFrom & "ON Company.ID = [Contact Link].CompanyID "
        
        sWhere = "WHERE ((('c' & [Contact]![ID])<>('c' & " & NOCONTACT & "))" & sFilter & ") "
        
        sOrderBy = "ORDER BY [Company]![CompanyName], [Company]![Address], "
        sOrderBy = sOrderBy & "[Company]![City], [Company]![State], "
        sOrderBy = sOrderBy & "[Company]![ZIP], [Contact]![FirstName], "
        sOrderBy = sOrderBy & "[Contact]![LastName], [Contact]![Title], "
        sOrderBy = sOrderBy & "[Contact]![Phone]; "
    Case clCONTACT3
    'This query selects SUMMARY CONTACT info for a treeview
    'It can select all CONTACTS or a specific COMPANY
        If Filter <> 0 Then
            sFilter = " AND ([Company]![ID]=" & CStr(Filter) & ")"
        Else
            sFilter = ""
        End If
        
        sSelect = "SELECT IIf(InStr(1,[Training],'N'),18, "
        sSelect = sSelect & "IIf(InStr(1,[Training],'C'),15, "
        sSelect = sSelect & "IIf(InStr(1,[Training],'L') Or "
        sSelect = sSelect & "InStr(1,[Training],'R'),12, "
        sSelect = sSelect & "IIf(InStr(1,[Training],'A') Or "
        sSelect = sSelect & "InStr(1,[Training],'D') Or "
        sSelect = sSelect & "InStr(1,[Training],'S'),9,0))))+[Skill]-1 AS cType, "
        sSelect = sSelect & "'p' & cstr([Contact Link]![CompanyID]) AS ParentID, "
        sSelect = sSelect & "IIf([Contact]![ID] Is Null,'','c' & "
        sSelect = sSelect & "cstr([Contact]![ID])) AS ChildID, "
        sSelect = sSelect & "[Contact]![FirstName] & ' ' & [Contact]![LastName] & ', ' & "
        sSelect = sSelect & "[Contact]![Title] & ' ' & [Contact]![Phone] AS sName, "
        sSelect = sSelect & "[Company]![CompanyName] & ', ' & [Company]![Address] & ', ' & "
        sSelect = sSelect & "[Company]![City] & ', ' & [Company]![State] & ', ' & "
        sSelect = sSelect & "[Company]![ZIP] & ', ' & [Contact]![FirstName] & ' ' & "
        sSelect = sSelect & "[Contact]![LastName] & ', ' & [Contact]![Title] & ' ' & "
        sSelect = sSelect & "[Contact]![Phone] AS SortKey "
        
        sFrom = "FROM Company INNER JOIN "
        sFrom = sFrom & "(Contact RIGHT JOIN [Contact Link] "
        sFrom = sFrom & "ON Contact.ID = [Contact Link].ContactID) "
        sFrom = sFrom & "ON Company.ID = [Contact Link].CompanyID "
        
        sWhere = "WHERE ((('c' & [Contact]![ID])<>('c' & " & NOCONTACT & "))" & sFilter & ") "
        
        sOrderBy = "ORDER BY [Company]![CompanyName], [Company]![Address], "
        sOrderBy = sOrderBy & "[Company]![City], [Company]![State], "
        sOrderBy = sOrderBy & "[Company]![ZIP], [Contact]![FirstName], "
        sOrderBy = sOrderBy & "[Contact]![LastName], [Contact]![Title], "
        sOrderBy = sOrderBy & "[Contact]![Phone]; "
        
    Case clCONTACT4 'Get company phone/fax for new CONTACT entry
    'This query selects SUMMARY CONTACT info for a treeview
    'It can select all CONTACTS or a specific COMPANY
        If Filter <> 0 Then
            sFilter = " AND ([Company]![ID]=" & CStr(Filter) & ")"
        Else
            sFilter = ""
        End If
        
        sSelect = "SELECT IIf(InStr(1,[Training],'N'),18, "
        sSelect = sSelect & "IIf(InStr(1,[Training],'C'),15, "
        sSelect = sSelect & "IIf(InStr(1,[Training],'L') Or "
        sSelect = sSelect & "InStr(1,[Training],'R'),12, "
        sSelect = sSelect & "IIf(InStr(1,[Training],'A') Or "
        sSelect = sSelect & "InStr(1,[Training],'D') Or "
        sSelect = sSelect & "InStr(1,[Training],'S'),9,0))))+[Skill]-1 AS cType, "
        sSelect = sSelect & "'p' & cstr([Contact Link]![CompanyID]) AS ParentID, "
        sSelect = sSelect & "IIf([Contact]![ID] Is Null,'','c' & "
        sSelect = sSelect & "cstr([Contact]![ID])) AS ChildID, "
        sSelect = sSelect & "[Contact]![FirstName] & ' ' & [Contact]![LastName] & ', ' & "
        sSelect = sSelect & "[Contact]![Title] & ' ' & [Contact]![Phone] AS sName, "
        sSelect = sSelect & "[Company]![CompanyName] & ', ' & [Company]![Address] & ', ' & "
        sSelect = sSelect & "[Company]![City] & ', ' & [Company]![State] & ', ' & "
        sSelect = sSelect & "[Company]![ZIP] & ', ' & [Contact]![FirstName] & ' ' & "
        sSelect = sSelect & "[Contact]![LastName] & ', ' & [Contact]![Title] & ' ' & "
        sSelect = sSelect & "[Contact]![Phone] AS SortKey "
        
        sFrom = "FROM Company INNER JOIN "
        sFrom = sFrom & "(Contact RIGHT JOIN [Contact Link] "
        sFrom = sFrom & "ON Contact.ID = [Contact Link].ContactID) "
        sFrom = sFrom & "ON Company.ID = [Contact Link].CompanyID "
        
        sWhere = "WHERE ((('c' & [Contact]![ID])<>('c' & " & NOCONTACT & "))" & sFilter & ") "
        
        sOrderBy = "ORDER BY [Company]![CompanyName], [Company]![Address], "
        sOrderBy = sOrderBy & "[Company]![City], [Company]![State], "
        sOrderBy = sOrderBy & "[Company]![ZIP], [Contact]![FirstName], "
        sOrderBy = sOrderBy & "[Contact]![LastName], [Contact]![Title], "
        sOrderBy = sOrderBy & "[Contact]![Phone]; "
        
    Case clCOMPANY4 'Get the new COMPANYS since the last get.
    'This query selects NEW SUMMARY COMPANY info for a treeview
    'by the date entered
        If Filter <> 0 Then
            sFilter = " AND ([Company]![DateEntered] > #" & CDate(Filter) & "#)"
        Else
            sFilter = ""
        End If
        
        sSelect = "SELECT IIf([Type]='D',1,IIf([Type]='E',2,IIf([Type]='F',3, "
        sSelect = sSelect & "IIf([Type]='I',4,IIf([Type]='O',5,IIf([Type]='R',6, "
        sSelect = sSelect & "IIf([Type]='S',7,IIf([Type]='U',8,8)))))))) AS cType, "
        sSelect = sSelect & "'p' & cstr([Company]![ID]) AS ParentID, "
        sSelect = sSelect & "'p' & cstr([Company]![ID]) AS ChildID, "
        sSelect = sSelect & "[Company]![CompanyName] & ', ' & "
        sSelect = sSelect & "[Company]![Address] & ', ' & [Company]![City] & ', ' & "
        sSelect = sSelect & "[Company]![State] AS sName, [Company]![CompanyName] & ', ' & "
        sSelect = sSelect & "[Company]![Address] & ', ' & [Company]![City] & ', ' & "
        sSelect = sSelect & "[Company]![State] AS SortKey "
        
        sFrom = "FROM Company "
        
        sWhere = "WHERE ((('p' & [Company]![ID])<>('p' & " & NOCUSTOMER & "))" & sFilter & ") "
        
        sOrderBy = "ORDER BY "
        sOrderBy = sOrderBy & "[Company]![CompanyName] & ', ' & [Company]![Address] & ', ' & "
        sOrderBy = sOrderBy & "[Company]![City] & ', ' & [Company]![State]; "
    
    Case clCASEID 'NOT WORKING!
    'Selects the company ID from CASEID
        sSelect = "SELECT SupportCalls.CompanyID AS CompanyID "
        sFrom = "FROM SupportCalls "
        sWhere = "WHERE (((SupportCalls.CaseID)=" & Filter & ")) "
        sOrderBy = "GROUP BY SupportCalls.CompanyID; "
    Case clNEWCALLS
    'Under construction
'        If Filter <> 0 Then
'            sFilter = " AND ([Company]![ID]=" & CStr(Filter) & ")"
'        Else
'            sFilter = ""
'        End If
        
        sSelect = "SELECT CallCode.CallType, "
        sSelect = sSelect & "[Company]![CompanyName] & ' (' & CStr([Company]![ID]) & ')' AS Company, "
        sSelect = sSelect & "Company.City, Company.State, "
        sSelect = sSelect & "[Contact]![FirstName] & IIf([Contact]![LastName]<>'--',' ' & [Contact]![LastName],'') AS Contact, "
        sSelect = sSelect & "Employees.LastName , Product.ProductName, "
        sSelect = sSelect & "SupportCalls.Note, SupportCalls.NoteDate "
        
        sFrom = "FROM CallCode INNER JOIN (Product INNER JOIN "
        sFrom = sFrom & "(Employees INNER JOIN (Contact INNER JOIN "
        sFrom = sFrom & "(Company INNER JOIN SupportCalls "
        sFrom = sFrom & "ON Company.ID = SupportCalls.CompanyID) "
        sFrom = sFrom & "ON Contact.ID = SupportCalls.ContactID) "
        sFrom = sFrom & "ON Employees.ID = SupportCalls.EmplID) "
        sFrom = sFrom & "ON Product.ID = SupportCalls.ProductID) "
        sFrom = sFrom & "ON CallCode.ID = SupportCalls.CallCodeID "
        
'        sWhere = "Where (([SupportCalls]![DateEntered] < CStr(Filter)))" & sFilter & ") """
        
        sOrderBy = "ORDER BY "
        sOrderBy = sOrderBy & "SupportCalls.DateEntered DESC;"
    
    Case clNEWCOMPANIES
    Case clNEWCOMPANIESNOCALLS
End Select

MakeQuery = sSelect & sFrom & sWhere & sOrderBy

On Error GoTo 0
Exit Function

MakeQueryErrorHandler:
mvarStatus = Err.Number 'GetLbData started
End Function
'===================================================================
Public Function GetLbData(iDataIndex As Integer, Optional Filter As Long) As ADODB.Recordset
'==========================================================
'Function is called by presentation layer which passes a
'request key (int) and filter (integer key) for company or contact
'Filter is used for showing filter lists of contacts for a company
'
'The return is disconnected ADO recordset
'==========================================================

Dim sQuery As String
Dim iCounter As Long

On Error GoTo GetDataError
'initialize
sQuery = ""
Set GetLbData = Nothing
Set rsLBData = Nothing
Set rsLBData = New ADODB.Recordset

sQuery = MakeQuery2(iDataIndex, Filter)

'Pass string and recordset to DBLogic module
If DBServer.GetListRS(sQuery, rsLBData, iDataIndex) = 1 Then
    Set GetLbData = rsLBData
    If GetLbData Is Nothing Then
        MsgBox "GetLBData not created", vbCritical, "BL-GetLBData"
        GoTo GetDataError
    End If
End If

If iDataIndex = clCONTACT2 Then
'    MsgBox GetLbData!cType & vbCrLf & rsLBData!cType
End If

On Error GoTo 0

Exit Function

GetDataError:
    mvarStatus = Err.Number 'GetLbData started
    MsgBox "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, , "BL-GetLBData"
End Function
'===================================================================
Public Function GetLbData2(iDataIndex As Integer, Optional Filter As String) As ADODB.Recordset
'==========================================================
'Function is called by presentation layer which passes a
'request key (int) and filter (integer key) for company or contact
'Filter is used for showing filter lists of contacts for a company
'
'The return is disconnected ADO recordset
'==========================================================

Dim sQuery As String
Dim iCounter As Long

On Error GoTo GetDataError
'initialize
sQuery = ""
Set GetLbData2 = Nothing
Set rsLBData = Nothing
Set rsLBData = New ADODB.Recordset

sQuery = MakeQuery2(iDataIndex, Filter)

'Pass string and recordset to DBLogic module
If DBServer.GetListRS(sQuery, rsLBData, iDataIndex) = 1 Then
    Set GetLbData2 = rsLBData
    If GetLbData2 Is Nothing Then
        MsgBox "GetLBData not created", vbCritical, "BL-GetLBData"
        GoTo GetDataError
    End If
End If

If iDataIndex = clCONTACT2 Then
'    MsgBox GetLbData!cType & vbCrLf & rsLBData!cType
End If

On Error GoTo 0

Exit Function

GetDataError:
    mvarStatus = Err.Number 'GetLbData started
    MsgBox "Error " & Err.Number & " from " & Err.Source & vbCrLf & Err.Description, , "BL-GetLBData"
End Function
Public Function AddCall(iCompanyID As Long, iContactID As Long, iCallCodeID As Long, iProductID As Long, iEmplID As Long, dNoteDate As Date, sNote As String, iCallTime As Integer, iCaseID As Long) As Long
'========================================================================
'This function gathers call information from the form
'Then accesses the DB layer to append the call
'========================================================================
Dim dEntryDate As Date
Dim sTable As String
'Dim NewCalls As CCalls
Dim rsTemp As ADODB.Recordset

On Error GoTo AddCallErrorHandler

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
    mvarStatus = Err.Number
    Exit Function
End If

On Error GoTo 0
Exit Function

AddCallErrorHandler:
    mvarStatus = Err.Number

End Function
Public Function AddCall2(iCompanyID As Long, iContactID As Long, iCallCodeID As Long, iProductID As Long, iEmplID As Long, dNoteDate As Date, sNote As String, iCallTime As Integer, iCaseID As Long, Optional lRecordID As Long) As Long
'========================================================================
'This function gathers call information from the form
'Then accesses the DB layer to append the call
'========================================================================
Dim dEntryDate As Date
Dim sTable As String
'Dim NewCalls As CCalls
Dim rsTemp As ADODB.Recordset

On Error GoTo AddCallErrorHandler

AddCall2 = 0 'assume failure

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

Set rsTemp = CreateRS(clCALL)

rsTemp.AddNew
    If lRecordID <> 0 Then rsTemp!ID = lRecordID Else rsTemp!ID = 0
'    rsTemp!ID = 0
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

AddCall2 = DBServer.AppendData(sTable, rsTemp)

If AddCall2 = 0 Then
    MsgBox "Failed to add call", vbExclamation, "BLogic::AddCall2 - Error"
    mvarStatus = Err.Number
    Exit Function
End If

On Error GoTo 0
Exit Function

AddCallErrorHandler:
    mvarStatus = Err.Number

End Function
Public Function AddCustomer(CompanyName As String, Address As String, City As String, State As String, Zip As String, Country As String, Phone As String, Fax As String, iEmplID As Long, cType As String, Optional ID As Long) As Long
Dim sTable As String
Dim rsTemp As ADODB.Recordset
Dim iTemp As Long

On Error GoTo AddCustomerErrorHandler

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

Set rsTemp = CreateRS(clCOMPANYBYDATE)

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
    mvarStatus = Err.Number
    Exit Function
End If

On Error GoTo 0
Exit Function

AddCustomerErrorHandler:
    mvarStatus = Err.Number
End Function
Public Function AddContact(FirstName As String, LastName As String, Title As String, iEmplID As Long, Phone As String, Skill As String, Training As String, Optional ID As Long) As Long
Dim sTable As String
Dim rsTemp As ADODB.Recordset
Dim iTemp As Long

On Error GoTo AddContactErrorHandler

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

Set rsTemp = CreateRS(clCONTACTBYCOMPID)

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

'Do While Not DBServer_onappenddone
'Loop

AddContact = DBServer.AppendData(sTable, rsTemp)

If AddContact = 0 Then
    MsgBox "Failed to add contact", vbInformation, "BL - AddContact"
    mvarStatus = Err.Number
    Exit Function
End If

Set rsTemp = Nothing

On Error GoTo 0
Exit Function

AddContactErrorHandler:
    mvarStatus = Err.Number

End Function
Public Function AddCompanyLink(iCompanyID As Long, iContactID As Long, iEmplID As Long) As Long
Dim sTable As String
Dim rsTemp As ADODB.Recordset
Dim iTemp As Long

On Error GoTo AddCompanyLinkErrorHandler
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
    MsgBox "Error entering Link;" & vbCrLf & "Err: " & Err.Number & vbCrLf & "Error Description: " & Err.Description, vbExclamation, "AddCompanyLink"
    mvarStatus = Err.Number
    Exit Function
End If

Set rsTemp = Nothing

On Error GoTo 0
Exit Function

AddCompanyLinkErrorHandler:
    MsgBox "Error entering Link;" & vbCrLf & "Err: " & Err.Number & vbCrLf & "Error Description: " & Err.Description, vbExclamation, "AddCompanyLink"
    mvarStatus = Err.Number
End Function
Public Function AddCallCode(sCallType As String, iEmplID As Long) As Long
Dim sTable As String
Dim rsTemp As ADODB.Recordset
Dim iTemp As Long

On Error GoTo AddCallCodeErrorHandler

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
    mvarStatus = Err.Number
    Exit Function
End If

On Error GoTo 0
Exit Function

AddCallCodeErrorHandler:
    mvarStatus = Err.Number
End Function
Public Function GetHistory(iID As Long) As ADODB.Recordset
Dim rsData As ADODB.Recordset
On Error GoTo GetHistoryError
'Pass string and recordset to DBLogic module
If DBServer.GetHistoryRS(MakeQuery2(clHISTORY, iID), rsData) Then
    Set GetHistory = rsData
    If GetHistory Is Nothing Then
        MsgBox "GetHistory not created", vbCritical, "BLServer - GetHistory"
        GoTo GetHistoryError
    End If
End If

On Error GoTo 0
Exit Function

GetHistoryError:
    mvarStatus = Err.Number
    Set rsData = Nothing
End Function
Public Function GetProductHistory(iID As Long) As ADODB.Recordset
Dim rsData As ADODB.Recordset

On Error GoTo GetProductHistoryError

'Pass string and recordset to DBLogic module
If DBServer.GetHistoryRS(MakeQuery(clPRODUCTHISTORY, iID), rsData) Then
    Set GetProductHistory = rsData
    If GetProductHistory Is Nothing Then
        MsgBox "GetProductHistory not created", vbCritical, "BLServer - GetProductHistory"
        GoTo GetProductHistoryError
    End If
End If

On Error GoTo 0
Exit Function

GetProductHistoryError:
    mvarStatus = Err.Number
    Set rsData = Nothing
End Function
Public Function AddEntries(oCol As Object) As Integer
'This function gathers call information from the form
'Then it accesses the DB layer to append the calls
Dim sTable As String

On Error GoTo AddEntriesError

AddEntries = 9999
mvarStatus = 9999
sTable = "SupportCalls"

If oCol Is Nothing Then
    MsgBox "Collection of calls empty", vbCritical, "BL-AddEntries"
    GoTo AddEntriesError
End If
If DBServer.AppendEntries(sTable, oCol) <> 0 Then
    MsgBox "Failed to add entries", vbExclamation, "BL-AddEntries"
    GoTo AddEntriesError
End If

AddEntries = 0

On Error GoTo 0
Exit Function

AddEntriesError:
    mvarStatus = Err.Number
End Function

Public Function AddEntries2(rs As ADODB.Recordset) As Integer
'This function gathers call information from the form
'Then it accesses the DB layer to append the calls
Dim sTable As String
Dim rsCalls As ADODB.Recordset

On Error GoTo AddEntries2Error

Set rsCalls = rs
AddEntries2 = 9999
mvarStatus = 9999
sTable = "SupportCalls"

If rsCalls Is Nothing Then
    MsgBox "Collection of calls empty", vbCritical, "BL-AddEntries2"
    GoTo AddEntries2Error
End If
If DBServer.AppendEntries(sTable, rsCalls) <> 0 Then
    MsgBox "Failed to add entries", vbExclamation, "BL-AddEntries2"
    GoTo AddEntries2Error
End If

AddEntries2 = 0

On Error GoTo 0
Exit Function

AddEntries2Error:
    mvarStatus = Err.Number
End Function
Private Sub Class_Initialize()
If (DBServer Is Nothing) Then
    Set DBServer = CreateObject("CallTrackerDBServer.DBServer")
    mvarStatus = 0 'DB Server started
End If
End Sub
Private Sub Class_Terminate()
If Not rsLBData Is Nothing Then Set rsLBData = Nothing
If Not rsData Is Nothing Then Set rsData = Nothing

If Not (DBServer Is Nothing) Then
    Set DBServer = Nothing
    mvarStatus = 0 'DB Server stopped
End If
End Sub
Private Function CreateRS(iIndex As Integer) As ADODB.Recordset
Dim rs As ADODB.Recordset
'========================================================================
'This function creates an ADO recordset used to pass contact/customer data to
'the DBServer.
'iIndex selects the proper rs (currently unused).
'========================================================================
On Error GoTo CreateRSErrorHandler

Set rs = New ADODB.Recordset

If rs Is Nothing Then
    mvarStatus = Err.Number
    Exit Function
End If

With rs
    Set .ActiveConnection = Nothing
    .CursorLocation = adUseClient
    .LockType = adLockBatchOptimistic

    With .Fields
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
        .Append "OpenCall", adInteger '
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
    .Open
End With

Set CreateRS = rs
Set rs = Nothing

On Error GoTo 0
Exit Function

CreateRSErrorHandler:
mvarStatus = Err.Number
Set rs = Nothing
End Function

Private Sub DBServer_OnAppendDone()
    RaiseEvent OnUpdateDone
End Sub

Private Sub DBServer_OnNewCustomerAppendDone()
    RaiseEvent OnNewCustomer
End Sub
Private Function MakeQuery2(Index As Integer, Optional Filter As Variant) As String
Dim sQuery As String
Dim sSelect As String
Dim sFrom As String
Dim sWhere As String
Dim sOrderBy As String
Dim sFilter As String

On Error GoTo MakeQuery2ErrorHandler
MakeQuery2 = "" ' assume failure

Select Case Index
    Case clCOMPANYBYID, clNEWCOMPANYBYID
    'This query selects SUMMARY company info for a treeview
    'It selects all recent companies or a specific company by ID
        If Filter <> 0 Then
            sFilter = " AND ([Company]![ID]=" & CStr(Filter) & ")"
        Else
            sFilter = ""
        End If
        
        sSelect = SelectCompany
        
        sFrom = "FROM Company "
        
        sWhere = "WHERE ((([Company]![ID])<>(" & NOCUSTOMER & "))" & sFilter & ") "
        
        sOrderBy = OrderByAddress
    
    Case clCOMPANYUPDATE
    'This query selects ALL COMPANY info for editing company record
    'It can select one or all companies
        If Filter <> 0 Then
            sFilter = " AND ([Company]![ID]=" & CStr(Filter) & ")"
        Else
            sFilter = ""
        End If
        
        sSelect = "SELECT * "
        
        sFrom = "FROM Company "
        
        sWhere = "WHERE ((([Company]![ID])<>(" & NOCUSTOMER & "))" & sFilter & ") "
        
        sOrderBy = OrderByAddress
    
    Case clCOMPANYBYDATE
    'This query selects SUMMARY company info for a treeview
    'It can select all companies or a specific company BY ID
        If Filter <> 0 Then
            sFilter = " AND ([Company]![DateEntered]>#" & Now() - CInt(Filter) & "#)" '(Now() - cdate(Filter)) & ")"
        Else
            sFilter = ""
        End If
        
        sSelect = SelectCompany
        
        sFrom = "FROM Company "
        
        sWhere = "WHERE ((([Company]![ID])<>(" & NOCUSTOMER & "))" & sFilter & ") "
        
        sOrderBy = OrderByAddress
        
    Case clCOMPANY4 'Get the new COMPANYS since the last get.
    'This query selects NEW SUMMARY COMPANY info for a treeview
    'by the date entered
        If Filter <> 0 Then
            sFilter = " AND ([Company]![DateEntered] > #" & CDate(Filter) & "#)"
        Else
            sFilter = ""
        End If
        
        sSelect = SelectCompany
        
        sFrom = "FROM Company "
        
        sWhere = "WHERE ((([Company]![ID])<>(" & NOCUSTOMER & "))" & sFilter & ") "
        
        sOrderBy = OrderByAddress
    
    Case clNEWCOMPANIES
    'Not Completed (filter not being used yet)
    'Selects companies with calls newer than filter or since last year
'        If Filter <> 0 And Filter <> "" Then
'            sFilter = " ([SupportCalls]![NoteDate]>#" & Format(Now() - CLng(Filter), "mm/dd/yy") & "#)"
'        Else
''            sFilter = " ((SupportCalls.NoteDate)>#" & Format(Now() - 365, "mm/dd/yy") & "#)"
'            sFilter = " OR ((([Company]![DateEntered]) > Now() - 7))"
'        End If
        
        sSelect = SelectFirstCompany

        sFrom = FromCompanyCalls
        
        sWhere = "WHERE ((([SupportCalls]![NoteDate]) > (Now() - 60))) Or (((Company.DateEntered) > Now() - 30)) "

        sOrderBy = GroupByFirstCompany
        
    Case clREFRESHCOMPANIES
    'Selects companies whose names contain the filter string
    'Used in Call Tracker 2
        sSelect = SelectFirstCompany
'        sSelect = SelectCompany

        sFrom = "FROM Company "
        
'        sWhere = "WHERE ((([Company]![CompanyName]) Like '%" & Filter & "%')) "
        sWhere = "GROUP BY ((IIf([Type]='D',1,"
        sWhere = sWhere & "IIf([Type]='E',2,"
        sWhere = sWhere & "IIf([Type]='F',3,"
        sWhere = sWhere & "IIf([Type]='I',4,"
        sWhere = sWhere & "IIf([Type]='O',5,"
        sWhere = sWhere & "IIf([Type]='R',6,"
        sWhere = sWhere & "IIf([Type]='S',7,"
        sWhere = sWhere & "IIf([Type]='U',8,8)))))))))), "
        
        sWhere = sWhere & "('p' & CStr([Company]![ID])), "
        
        sWhere = sWhere & "('p' & CStr([Company]![ID])), "
        
        sWhere = sWhere & "([Company]![CompanyName] & ', ' & "
        sWhere = sWhere & "[Company]![Address] & ', ' & "
        sWhere = sWhere & "[Company]![City] & ', ' & "
        sWhere = sWhere & "[Company]![State]), "
        sWhere = sWhere & "([Company]![CompanyName])"
        
        sWhere = sWhere & "Having (((([Company]![CompanyName])) "
        sWhere = sWhere & "Like  '%" & Filter & "%' ))"

        sOrderBy = OrderByAddress
        
    Case clCONTACTBYID, clNEWCONTACTBYID
    'This query selects SUMMARY CONTACT info for a treeview
    'It can select all CONTACTS or a specific CONTACT BY ID
        If Filter <> 0 And Filter <> "" Then
            sFilter = "WHERE ([Contact]![ID]=" & CStr(Filter) & ") "
        Else
            sFilter = ""
        End If
        
        sSelect = SelectContact
        
        sFrom = FromCompanyContact
        
        sWhere = sFilter
        
        sOrderBy = OrderByAddressContact
        
    Case clCONTACTBYCOMPID
    'This query selects SUMMARY CONTACT info for a treeview
    'It can select all CONTACTS or a specific COMPANY
        If Filter <> 0 Then
            sFilter = " AND ([Company]![ID]=" & CStr(Filter) & ")"
        Else
            sFilter = ""
        End If
        
        sSelect = SelectContact
        
        sFrom = FromCompanyContact
        
        sWhere = "WHERE (([Contact]![ID]<>" & NOCONTACT & ")" & sFilter & ") "
        
        sOrderBy = OrderByAddressContact
        
    Case 25 'clCONTACTssinceDATe
    'This query selects SUMMARY CONTACT info for a treeview
    'It selects all recent CONTACTS or a specific CONTACT by DATE
    'Call Tracker 2 does not use this query
        If Filter <> 0 Then
            sFilter = " AND ([Contact]![DateEntered]>#" & (CDate(Filter)) & "#)" '(Now() - cdate(Filter)) & ")"
        Else
            sFilter = ""
        End If
        
        sSelect = SelectContact
        
        sFrom = FromCompanyContact
        
        sWhere = "WHERE ((([Contact]![ID])<>(" & NOCONTACT & "))" & sFilter & ") "
        
        sOrderBy = OrderByAddressContact
        
    Case clCONTACT4
    'Get company phone/fax for new CONTACT entry
    'This query selects SUMMARY CONTACT info for a treeview
    'It can select all CONTACTS or a specific COMPANY
        If Filter <> 0 Then
            sFilter = " AND ([Company]![ID]=" & CStr(Filter) & ")"
        Else
            sFilter = ""
        End If
        
        sSelect = SelectContact
        
        sFrom = FromCompanyContact
        
        sWhere = "WHERE ((([Contact]![ID])<>(" & NOCONTACT & "))" & sFilter & ") "
        
        sOrderBy = OrderByAddressContact
        
    Case clCONTACTUPDATE 'get the contact data fields
    'This query selects ALL CONTACT info for editing Contact info
    'It can select one or all Contacts by ID
    
        sSelect = "SELECT * "
        sFrom = "FROM Contact "
        
        If Filter <> 0 Then 'get matching contact
            sWhere = "WHERE ([Contact]![ID]=" & CStr(Filter) & ") "
        Else 'get all contacts
            sWhere = "WHERE ([Contact]![ID]<>" & NOCONTACT & ") "
        End If
        
        sOrderBy = "ORDER BY [Contact]![FirstName] & ' ' & [Contact]![LastName] & ', ' & [Contact]![Title];"
    
    Case clPRODUCT 'working
    'This query selects PRODUCTS for a listview
        sSelect = "SELECT [Product]![ProductName] AS sName, Product.ID "
        sFrom = "From Product "
        sWhere = "WHERE ([Product]![ID]<>" & NOPRODUCT & ") "
        sOrderBy = "ORDER BY [Product]![Status], [Product]![ProductName];"
        
    Case clCALLCODE 'working
    'This query selects CALLCODES for a listview
    'Call Tracker 2 does not use this query
        sSelect = "SELECT [CallCode]![CallType] AS sName, [CallCode]![ID] AS ID "
        sFrom = "From CallCode "
        sWhere = "WHERE ([CallCode]![ID]<>" & NOCODE & ") "
        sOrderBy = "ORDER BY [CallCode]![CallType];"
        
    Case clLINK 'working
    'This query selects ALL CONTACT LINK data
    'It connects people to companies
        sSelect = "SELECT [Contact Link].* "
        sFrom = "FROM [Contact Link];"
        sWhere = ""
        sOrderBy = ""
        
    Case clEMPLOYEE 'working
    'This query selects EMPLOYEES for a listview
        sSelect = "SELECT Employees.ID AS ID, "
        sSelect = sSelect & "IIf(Len([Employees]![FirstName])=1,"
        sSelect = sSelect & "[Employees]![MiddleName],[Employees]![FirstName])"
        sSelect = sSelect & " & ' ' & [Employees]![LastName] AS sName "
        sFrom = "FROM Employees "
        sWhere = ""
        sOrderBy = "ORDER BY Employees.LastName;"
        
    Case clHISTORY 'working
    'This query selects the CALL HISTORY for one company by ID
    'Data fills the flexgrid
        sSelect = SelectCalls
        
        sFrom = FromAllTables
    
        sWhere = "Where (((Company.ID) = " & Filter & ")) "
        
        sOrderBy = OrderByCallDate
        
    Case clHISTORY2 'working
    'This query selects the CALL HISTORY for one company by CASE ID
    'Data fills the flexgrid
        If Filter <> "0" Then
            sFilter = " AND ((SupportCalls.CaseID) = " & Filter & ") "
        Else
            sFilter = ""
        End If
        
        sSelect = SelectCalls
        
        sFrom = FromAllTables
    
        sWhere = "WHERE ((([SupportCalls]![NoteDate]) > Now() - 180)"
        sWhere = sWhere & sFilter & ") "
        
        sOrderBy = OrderByCallDate
        
    Case clPRODUCTHISTORY
    'This query selects PRODUCT CALL HISTORY info for a chosen product
    'Used for analysis, not call entry
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
        
        sFrom = FromAllTables
    
        sWhere = "Where (((Product.ID) = " & Filter & ")) "
        
        sOrderBy = OrderByCallDate
        
    Case clCASEID 'NOT WORKING!
    'Selects the company ID from CASEID
        If Filter = 0 Then
            Filter = "'%'"
        End If
        sSelect = "SELECT SupportCalls.CompanyID AS CompanyID "
        sFrom = "FROM SupportCalls "
        sWhere = "WHERE (((SupportCalls.CaseID) LIKE " & Filter & ")) "
        sOrderBy = "GROUP BY SupportCalls.CompanyID; "
    Case clNEWCALLS
    'Under construction
    'Selects calls newer than a selected date (FILTER)
        If Filter <> 0 Then
            sFilter = " ((([SupportCalls]![NoteDate]) > Now() - " & CLng(Filter) & "))"
        Else
            sFilter = " ((([SupportCalls]![NoteDate]) > Now() - 60))"
        End If
        
        sSelect = "SELECT CallCode.CallType, "
        sSelect = sSelect & "[Company]![CompanyName] & ' (' & CStr([Company]![ID]) & ')' AS Company, "
        sSelect = sSelect & "Company.City, Company.State, "
        sSelect = sSelect & "[Contact]![FirstName] & IIf([Contact]![LastName]<>'--',' ' & [Contact]![LastName],'') AS Contact, "
        sSelect = sSelect & "Employees.LastName , Product.ProductName, "
        sSelect = sSelect & "SupportCalls.Note, SupportCalls.NoteDate "
        
        sFrom = FromAllTables
        
        sWhere = "Where (" & sFilter & ") "
             
        sOrderBy = "ORDER BY "
        sOrderBy = sOrderBy & "SupportCalls.DateEntered DESC;"
    
End Select

MakeQuery2 = sSelect & sFrom & sWhere & sOrderBy
'MsgBox sSelect & vbCrLf & vbCrLf & sFrom & vbCrLf & vbCrLf & sWhere & vbCrLf & vbCrLf & sOrderBy
On Error GoTo 0
Exit Function

MakeQuery2ErrorHandler:
mvarStatus = Err.Number 'GetLbData started
End Function

