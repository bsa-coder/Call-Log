VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CEmployees"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EBC89101AE"
'local variable(s) to hold property value(s)
'##ModelId=39EBC8910317
Private mvarsName As String 'local copy
'##ModelId=39EBC8910385
Private mvarID As Long 'local copy
'##ModelId=39EBC892000D
Private mvarCEmployees As CEmployees
'local variable(s) to hold property value(s)
Private mvarsTraining As String 'local copy
Private mvarsSkill As String 'local copy
Private mvarsPhone As String 'local copy
'##ModelId=39EBC892010F
Public Property Get CEmployees() As CEmployees
    If mvarCEmployees Is Nothing Then
        Set mvarCEmployees = New CEmployees
    End If


    Set CEmployees = mvarCEmployees
End Property

'##ModelId=39EBC892001F
Public Property Set CEmployees(vData As CEmployees)
    Set mvarCEmployees = vData
End Property
'##ModelId=39EBC8920191
Private Sub Class_Terminate()
    Set mvarCEmployees = Nothing
End Sub

'##ModelId=39EBC89201EC
Public Property Let ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ID = 5
    mvarID = vData
End Property

'##ModelId=39EBC89202FA
Public Property Get ID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ID
    ID = mvarID
End Property

'##ModelId=39EBC8920390
Public Property Let sName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sName = 5
    mvarsName = vData
End Property

'##ModelId=39EBC89300CB
Public Property Get sName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sName
    sName = mvarsName
End Property



