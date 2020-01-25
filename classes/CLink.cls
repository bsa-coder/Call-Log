VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CLinks"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EBC88300B4"
'local variable(s) to hold property value(s)
'##ModelId=39EBC8830154
Private mvarID As Integer 'local copy
'##ModelId=39EBC88301AE
Private mvarCompanyID As Integer 'local copy
'##ModelId=39EBC8830208
Private mvarContactID As Integer 'local copy
'##ModelId=39EBC883026E
Private mvarCLinks As CLinks


'##ModelId=39EBC8830348
Public Property Get CLinks() As CLinks
    If mvarCLinks Is Nothing Then
        Set mvarCLinks = New CLinks
    End If


    Set CLinks = mvarCLinks
End Property


'##ModelId=39EBC8830280
Public Property Set CLinks(vData As CLinks)
    Set mvarCLinks = vData
End Property
'##ModelId=39EBC88303C1
Private Sub Class_Terminate()
    Set mvarCLinks = Nothing
End Sub



'##ModelId=39EBC884001F
Public Property Let ContactID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ContactID = 5
    mvarContactID = vData
End Property


'##ModelId=39EBC8840105
Public Property Get ContactID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ContactID
    ContactID = mvarContactID
End Property



'##ModelId=39EBC884017D
Public Property Let CompanyID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CompanyID = 5
    mvarCompanyID = vData
End Property


'##ModelId=39EBC8840278
Public Property Get CompanyID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CompanyID
    CompanyID = mvarCompanyID
End Property



'##ModelId=39EBC8840304
Public Property Let ID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ID = 5
    mvarID = vData
End Property


'##ModelId=39EBC8850020
Public Property Get ID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ID
    ID = mvarID
End Property



