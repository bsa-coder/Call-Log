Attribute VB_Name = "MQueryStrings"
Function SelectContact() As String
Dim sSelect As String
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
    sSelect = sSelect & "[Contact]![DateEntered] AS EDate "
SelectContact = sSelect
End Function
Function SelectCompany() As String
'Get company name + address (for TVW)
Dim sSelect As String
SelectCompany = ""
    sSelect = "SELECT IIf([Type]='D',1,IIf([Type]='E',2,IIf([Type]='F',3, "
    sSelect = sSelect & "IIf([Type]='I',4,IIf([Type]='O',5,IIf([Type]='R',6, "
    sSelect = sSelect & "IIf([Type]='S',7,IIf([Type]='U',8,8)))))))) AS cType, "
    sSelect = sSelect & "'p' & cstr([Company]![ID]) AS ParentID, "
    sSelect = sSelect & "'p' & cstr([Company]![ID]) AS ChildID, "
    sSelect = sSelect & "[Company]![CompanyName] & ', ' & "
    sSelect = sSelect & "[Company]![Address] & ', ' & [Company]![City] & ', ' & "
    sSelect = sSelect & "[Company]![State] AS sName, "
    sSelect = sSelect & "[Company]![Address] AS Address, [Company]![City] AS City, "
    sSelect = sSelect & "[Company]![ID] AS ID, [Company]![State] AS State "
SelectCompany = sSelect
End Function
Function SelectCompanyDate() As String
Dim sSelect As String
SelectCompanyDate = ""
    sSelect = "SELECT IIf([Type]='D',1,IIf([Type]='E',2,IIf([Type]='F',3, "
    sSelect = sSelect & "IIf([Type]='I',4,IIf([Type]='O',5,IIf([Type]='R',6, "
    sSelect = sSelect & "IIf([Type]='S',7,IIf([Type]='U',8,8)))))))) AS cType, "
    sSelect = sSelect & "'p' & cstr([Company]![ID]) AS ParentID, "
    sSelect = sSelect & "'p' & cstr([Company]![ID]) AS ChildID, "
    sSelect = sSelect & "[Company]![CompanyName] & ', ' & "
    sSelect = sSelect & "[Company]![Address] & ', ' & [Company]![City] & ', ' & "
    sSelect = sSelect & "[Company]![State] AS sName, "
    sSelect = sSelect & "[Company]![Address] AS Address, [Company]![City] AS City, "
    sSelect = sSelect & "[Company]![State] AS State,[Company]![DateEntered] AS EDate "
SelectCompanyDate = sSelect
End Function
Function SelectFirstCompany() As String
Dim sSelect As String
'    sSelect = "SELECT First((IIf([Type]='D',1,"
'    sSelect = sSelect & "IIf([Type]='E',2,IIf([Type]='F',3,"
'    sSelect = sSelect & "IIf([Type]='I',4,IIf([Type]='O',5,"
'    sSelect = sSelect & "IIf([Type]='R',6,IIf([Type]='S',7,"
'    sSelect = sSelect & "IIf([Type]='U',8,8)))))))))) AS cType,"
'
'    sSelect = sSelect & "First('p' & CStr([Company]![ID])) AS ParentID,"
'
'    sSelect = sSelect & "First('p' & CStr([Company]![ID])) AS ChildID,"
'
'    sSelect = sSelect & "First([Company]![CompanyName] & ', ' & "
'    sSelect = sSelect & "[Company]![Address] & ', ' & "
'    sSelect = sSelect & "[Company]![City] & ', ' & "
'    sSelect = sSelect & "[Company]![State]) AS sName "
    sSelect = "SELECT First((IIf([Type]='D',1,IIf([Type]='E',2,IIf([Type]='F',3,IIf([Type]='I',4,IIf([Type]='O',5,IIf([Type]='R',6,IIf([Type]='S',7,IIf([Type]='U',8,8)))))))))) AS cType, First('p' & CStr([Company]![ID])) AS ParentID, First('p' & CStr([Company]![ID])) AS ChildID, First([Company]![CompanyName] & ', ' & [Company]![Address] & ', ' & [Company]![City] & ', ' & [Company]![State]) AS sName "
    
SelectFirstCompany = sSelect

End Function
Function SelectCalls() As String
Dim sSelect As String
    sSelect = "SELECT Employees.LastName AS LastName, "
    sSelect = sSelect & "SupportCalls.EmplID AS EmplID, "
    sSelect = sSelect & "Company.CompanyName AS CompanyName, "
    sSelect = sSelect & "[Contact]![FirstName] & IIf([Contact]![LastName]<>'--',' ' & [Contact]![LastName],'') AS ContactName, "
    sSelect = sSelect & "SupportCalls.ContactID AS ContactID, "
    sSelect = sSelect & "[Contact]![Phone] AS Phone, "
    sSelect = sSelect & "Product.ProductName AS ProductName, "
    sSelect = sSelect & "SupportCalls.ProductID AS ProductID, "
    sSelect = sSelect & "SupportCalls.CallCodeID AS CallCodeID, "
    sSelect = sSelect & "CallCode.CallType AS CallType, "
    sSelect = sSelect & "SupportCalls.Note AS sNote, "
    sSelect = sSelect & "SupportCalls.NoteDate AS NoteDate, "
    sSelect = sSelect & "SupportCalls.CallTime AS iCallTime, "
    sSelect = sSelect & "SupportCalls.OpenCall AS CallOpen, "
    sSelect = sSelect & "SupportCalls.ID AS RecordID, "
    sSelect = sSelect & "SupportCalls.CaseID AS CaseID "
SelectCalls = sSelect
End Function
Function FromCompanyContact() As String
Dim sFrom As String
    sFrom = "FROM Company INNER JOIN "
    sFrom = sFrom & "(Contact RIGHT JOIN [Contact Link] "
    sFrom = sFrom & "ON Contact.ID = [Contact Link].ContactID) "
    sFrom = sFrom & "ON Company.ID = [Contact Link].CompanyID "
FromCompanyContact = sFrom
End Function
Function FromCompanyCalls2() As String
Dim sFrom As String
    sFrom = "FROM Company INNER JOIN SupportCalls ON "
    sFrom = sFrom & "Company.ID = SupportCalls.CompanyID "
FromCompanyCalls2 = sFrom
End Function
Function FromCompanyCalls() As String
Dim sFrom As String
    sFrom = "FROM Company LEFT JOIN SupportCalls ON "
    sFrom = sFrom & "Company.ID = SupportCalls.CompanyID "
FromCompanyCalls = sFrom
End Function
Function FromAllTables() As String
Dim sFrom As String
    sFrom = "FROM "
    sFrom = sFrom & "(Company INNER JOIN "
    sFrom = sFrom & "(Contact INNER JOIN "
    sFrom = sFrom & "(Employees INNER JOIN "
    sFrom = sFrom & "(CallCode INNER JOIN "
    sFrom = sFrom & "(Product INNER JOIN "
    sFrom = sFrom & "SupportCalls "
    sFrom = sFrom & "ON Product.ID = SupportCalls.ProductID) "
    sFrom = sFrom & "ON CallCode.ID = SupportCalls.CallCodeID) "
    sFrom = sFrom & "ON Employees.ID = SupportCalls.EmplID) "
    sFrom = sFrom & "ON Contact.ID = SupportCalls.ContactID) "
    sFrom = sFrom & "ON Company.ID = SupportCalls.CompanyID) "
FromAllTables = sFrom
End Function
Function GroupByFirstCompany() As String
Dim sOrderBy As String
'    sOrderBy = "GROUP BY ('p' & CStr([Company]![ID])) "
'    sOrderBy = sOrderBy & "ORDER BY First([Company]![CompanyName] "
'    sOrderBy = sOrderBy & "& ' (' & CStr([Company]![ID]) & ')');"
    
    sOrderBy = "GROUP BY ('p' & CStr([Company]![ID])) "

    sOrderBy = sOrderBy & "ORDER BY First([Company]![CompanyName] "
    sOrderBy = sOrderBy & "& ', ' & [Company]![Address] & ', ' & [Company]![City] "
    sOrderBy = sOrderBy & "& ', ' & [Company]![State]), "
    sOrderBy = sOrderBy & "First([Company]![CompanyName] & ' (' & CStr([Company]![ID]) & ')');"
GroupByFirstCompany = sOrderBy
End Function
Function OrderByFirstCompany() As String
Dim sOrderBy As String
    sOrderBy = "ORDER BY First([Company]![CompanyName] "
    sOrderBy = sOrderBy & "& ' (' & CStr([Company]![ID]) & ')');"
OrderByFirstCompany = sOrderBy
End Function
Function OrderByAddress() As String
Dim sOrderBy As String
    sOrderBy = "ORDER BY "
    sOrderBy = sOrderBy & "[Company]![CompanyName] & ', ' & [Company]![Address] & ', ' & "
    sOrderBy = sOrderBy & "[Company]![City] & ', ' & [Company]![State]; "
OrderByAddress = sOrderBy
End Function
Function OrderByAddressContact() As String
Dim sOrderBy As String
    sOrderBy = "ORDER BY [Company]![CompanyName], [Company]![Address], "
    sOrderBy = sOrderBy & "[Company]![City], [Company]![State], "
    sOrderBy = sOrderBy & "[Company]![ZIP], [Contact]![FirstName], "
    sOrderBy = sOrderBy & "[Contact]![LastName], [Contact]![Title], "
    sOrderBy = sOrderBy & "[Contact]![Phone]; "
OrderByAddressContact = sOrderBy
End Function
Function OrderByCallDate() As String
Dim sOrderBy As String
    sOrderBy = "ORDER BY SupportCalls.NoteDate DESC;"
OrderByCallDate = sOrderBy
End Function

