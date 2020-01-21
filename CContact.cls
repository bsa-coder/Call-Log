VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CContacts"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarsName As String 'local copy
Private mvarID As Integer 'local copy
Private mvarsLastName As String 'local copy
Private mvarsFirstName As String 'local copy
Private mvarsTitle As String 'local copy
Private mvarCContacts As CContacts
Private mvarsPhone As String 'local copy
Private mvarsSkill As String 'local copy
Private mvarsTraining As String 'local copy
Public Property Let sTraining(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sTraining = 5
If InStr(1, "ADSRLC", Left(vData, 1)) <> 0 Then
    mvarsTraining = vData
End If
End Property
Public Property Get sTraining() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sTraining
    sTraining = mvarsTraining
End Property
Public Property Let sSkill(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSkill = 5
If InStr(1, "123", Left(vData, 1)) <> 0 Then
    mvarsSkill = vData
End If
End Property
Public Property Get sSkill() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSkill
    sSkill = mvarsSkill
End Property
Public Property Let sPhone(ByVal vData As String)
Attribute sPhone.VB_Description = "list of contact numbers, Phone, Fax, Cell, Pager"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sPhone = 5
If vData <> "" Then
    mvarsPhone = vData
End If
End Property
Public Property Get sPhone() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sPhone
    sPhone = mvarsPhone
End Property
Public Property Get CContacts() As CContacts
    If mvarCContacts Is Nothing Then
        Set mvarCContacts = New CContacts
    End If
    Set CContacts = mvarCContacts
End Property
Public Property Set CContacts(vData As CContacts)
    Set mvarCContacts = vData
End Property

Private Sub Class_Initialize()

mvarsTitle = "--"
mvarsName = "--"
mvarsLastName = "--"
mvarsFirstName = "--"
mvarsTitle = "--"
mvarsPhone = "000-000-0000"
mvarsSkill = "1"
mvarsTraining = "N"

End Sub

Private Sub Class_Terminate()
    Set mvarCContacts = Nothing
End Sub
Public Property Let sTitle(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sTitle = 5
If vData <> "" Then
    mvarsTitle = vData
End If
End Property
Public Property Get sTitle() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sTitle
    sTitle = mvarsTitle
End Property
Public Property Let sFirstName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sFirstName = 5
If vData <> "" Then
    mvarsFirstName = vData
End If
End Property
Public Property Get sFirstName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sFirstName
    sFirstName = mvarsFirstName
End Property
Public Property Let sLastName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sLastName = 5
If vData <> "" Then
    mvarsLastName = vData
End If
End Property
Public Property Get sLastName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sLastName
    sLastName = mvarsLastName
End Property
Public Property Let ID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ID = 5
    mvarID = vData
End Property
Public Property Get ID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ID
    ID = mvarID
End Property
Public Property Let sName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sName = 5
If vData <> "" Then
    mvarsName = vData
End If
End Property
Public Property Get sName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sName
    sName = mvarsName
End Property



