VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCallCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CCallCodes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarID As Integer 'local copy
Private mvarsName As String 'local copy
Private mvarCCallCodes As CCallCodes


Public Property Get CCallCodes() As CCallCodes
    If mvarCCallCodes Is Nothing Then
        Set mvarCCallCodes = New CCallCodes
    End If


    Set CCallCodes = mvarCCallCodes
End Property


Public Property Set CCallCodes(vData As CCallCodes)
    Set mvarCCallCodes = vData
End Property
Private Sub Class_Terminate()
    Set mvarCCallCodes = Nothing
End Sub



Public Property Let sName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sName = 5
    mvarsName = vData
End Property


Public Property Get sName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sName
    sName = mvarsName
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



