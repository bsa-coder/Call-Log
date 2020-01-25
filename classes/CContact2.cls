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
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EBC8770355"
'local variable(s) to hold property value(s)
'##ModelId=39EBC878000D
Private mvarsName As String 'local copy
'##ModelId=39EBC8780068
Private mvarID As Integer 'local copy
'##ModelId=39EBC87800C2
Private mvarsLastName As String 'local copy
'##ModelId=39EBC8780144
Private mvarsFirstName As String 'local copy
'##ModelId=39EBC87801A8
Private mvarsTitle As String 'local copy
'##ModelId=39EBC8780222
Private mvarCContacts As CContacts


'##ModelId=39EBC8780343
Public Property Get CContacts() As CContacts
    If mvarCContacts Is Nothing Then
        Set mvarCContacts = New CContacts
    End If


    Set CContacts = mvarCContacts
End Property


'##ModelId=39EBC8780234
Public Property Set CContacts(vData As CContacts)
    Set mvarCContacts = vData
End Property
'##ModelId=39EBC87803C5
Private Sub Class_Terminate()
    Set mvarCContacts = Nothing
End Sub



'##ModelId=39EBC8790037
Public Property Let sTitle(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sTitle = 5
    mvarsTitle = vData
End Property


'##ModelId=39EBC879013B
Public Property Get sTitle() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sTitle
    sTitle = mvarsTitle
End Property



'##ModelId=39EBC87901C8
Public Property Let sFirstName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sFirstName = 5
    mvarsFirstName = vData
End Property


'##ModelId=39EBC87902D6
Public Property Get sFirstName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sFirstName
    sFirstName = mvarsFirstName
End Property



'##ModelId=39EBC8790376
Public Property Let sLastName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sLastName = 5
    mvarsLastName = vData
End Property


'##ModelId=39EBC87A00BB
Public Property Get sLastName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sLastName
    sLastName = mvarsLastName
End Property



'##ModelId=39EBC87A015B
Public Property Let ID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ID = 5
    mvarID = vData
End Property


'##ModelId=39EBC87A029B
Public Property Get ID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ID
    ID = mvarID
End Property



'##ModelId=39EBC87A0350
Public Property Let sName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sName = 5
    mvarsName = vData
End Property


'##ModelId=39EBC87B00B2
Public Property Get sName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sName
    sName = mvarsName
End Property



