VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CCustomers"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EBC87000FC"
'local variable(s) to hold property value(s)
'##ModelId=39EBC87001A7
Private mvarsName As String 'local copy
'##ModelId=39EBC8700201
Private mvarsAddress As String 'local copy
'##ModelId=39EBC8700265
Private mvarID As Integer 'local copy
'##ModelId=39EBC87002D3
Private mvarsCity As String 'local copy
'##ModelId=39EBC8700341
Private mvarsState As String 'local copy
'##ModelId=39EBC87003AF
Private mvarsZIP As String 'local copy
'##ModelId=39EBC871003F
Private mvarsPhone As String 'local copy
'##ModelId=39EBC87100C2
Private mvarsFAX As String 'local copy
'##ModelId=39EBC871014E
Private mvarCCustomers As CCustomers
'local variable(s) to hold property value(s)
'##ModelId=39EBC871019E
Private mvarsCountry As String 'local copy
'local variable(s) to hold property value(s)
Private mvarsType As String 'local copy
Private mvariType As Integer 'local copy
Public Property Let sType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sType = 5
'    If vData = "" Then
'        mvarsType = "U" 'default is to be a user
    If InStr(1, "UIDOSREF", Left(vData, 1)) <> 0 Then 'U/D/S/O/R/I
            mvarsType = vData
    Else
            mvarsType = "U"
    End If
    Select Case mvarsType
        Case "U" 'end user
            mvariType = 0
        Case "I" 'integrator
            mvariType = 1
        Case "D" 'distributor
            mvariType = 2
        Case "O" 'oem
            mvariType = 3
        Case "S" 'service house
            mvariType = 4
        Case "R" 'rep
            mvariType = 5
        Case "E" 'Eurotherm drives company
            mvariType = 6
        Case "F" 'Invensys sister company
            mvariType = 7
        Case Else
            mvariType = 0
    End Select
End Property
Public Property Get sType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sType
    sType = mvarsType
End Property
Public Property Let iType(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sType = 5
    If Not (vData < 0 Or vData > 7) Then
        mvariType = vData
    Else
        mvariType = 0
    End If
    Select Case mvariType
        Case 0 'end user
            mvarsType = "U"
        Case 1 'integrator
            mvarsType = "I"
        Case 2 'distributor
            mvarsType = "D"
        Case 3 'oem
            mvarsType = "O"
        Case 4 'service house
            mvarsType = "S"
        Case 5 'rep
            mvarsType = "R"
        Case 6 'Eurotherm drives company
            mvarsType = "E"
        Case 7 'Invensys sister company
            mvarsType = "F"
        Case Else
            mvarsType = "U"
    End Select
End Property
Public Property Get iType() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sType
    iType = mvariType
End Property
'##ModelId=39EBC871022A
Public Property Let sCountry(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCountry = 5
    If vData <> "" Then
        mvarsCountry = vData
    End If
End Property
'##ModelId=39EBC871039D
Public Property Get sCountry() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCountry
    sCountry = mvarsCountry
End Property
'##ModelId=39EBC8720181
Public Property Get CCustomers() As CCustomers
    If mvarCCustomers Is Nothing Then
        Set mvarCCustomers = New CCustomers
    End If

    Set CCustomers = mvarCCustomers
End Property
'##ModelId=39EBC8720055
Public Property Set CCustomers(vData As CCustomers)
    Set mvarCCustomers = vData
End Property

Private Sub Class_Initialize()
    mvarsName = "--"
    mvarsAddress = "--"
    mvarID = 0
    mvarsCity = "--"
    mvarsState = "--"
    mvarsZIP = "--"
    mvarsPhone = "000-000-0000"
    mvarsFAX = "000-000-0000"
    mvarsCountry = "--"
    mvarsType = "U"
End Sub

'##ModelId=39EBC872022C
Private Sub Class_Terminate()
    Set mvarCCustomers = Nothing
End Sub
'##ModelId=39EBC872029A
Public Property Let sFAX(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sFAX = 5
If vData <> "" Then
    mvarsFAX = vData
'Else
'    mvarsFAX = "000-000-0000"
End If
End Property
'##ModelId=39EBC87203E4
Public Property Get sFAX() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sFAX
    sFAX = mvarsFAX
End Property
'##ModelId=39EBC87300B1
Public Property Let sPhone(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sPhone = 5
If vData <> "" Then
    mvarsPhone = vData
'Else
'    mvarsPhone = "000-000-0000"
End If
End Property
'##ModelId=39EBC873020F
Public Property Get sPhone() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sPhone
    sPhone = mvarsPhone
End Property
'##ModelId=39EBC87302CD
Public Property Let sZIP(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sZIP = 5
If vData <> "" Then
    mvarsZIP = vData
'Else
'    mvarsZIP = "--"
End If
End Property
'##ModelId=39EBC874004E
Public Property Get sZIP() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sZIP
    sZIP = mvarsZIP
End Property
'##ModelId=39EBC8740116
Public Property Let sState(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sState = 5
If vData <> "" Then
    mvarsState = vData
'Else
'    mvarsState = "--"
End If
End Property
'##ModelId=39EBC8740293
Public Property Get sState() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sState
    sState = mvarsState
End Property
'##ModelId=39EBC874036F
Public Property Let sCity(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCity = 5
If vData <> "" Then
    mvarsCity = vData
'Else
'    mvarsCity = "--"
End If
End Property
'##ModelId=39EBC8750118
Public Property Get sCity() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCity
    sCity = mvarsCity
End Property
'##ModelId=39EBC87501FE
Public Property Let ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ID = 5
If vData <> 0 Then
    mvarID = vData
'Else
'    mvarID = "--"
End If
End Property
'##ModelId=39EBC87503A2
Public Property Get ID() As Long
Attribute ID.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ID
    ID = mvarID
End Property
'##ModelId=39EBC87600A1
Public Property Let sAddress(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sAddress = 5
If vData <> "" Then
    mvarsAddress = vData
'Else
'    mvarsAddress = "--"
End If
End Property
'##ModelId=39EBC876024F
Public Property Get sAddress() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sAddress
    sAddress = mvarsAddress
End Property
'##ModelId=39EBC8760354
Public Property Let sName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sName = 5
If vData <> "" Then
    mvarsName = vData
'Else
'    mvarsName = "--"
End If
End Property
'##ModelId=39EBC8770124
Public Property Get sName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sName
    sName = mvarsName
End Property



