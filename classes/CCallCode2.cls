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
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EBC87D0304"
'##ModelId=39EBC87D03AE
Private mvarID As Integer 'local copy
'##ModelId=39EBC87E0020
Private mvarsName As String 'local copy
'##ModelId=39EBC87E0084
Private mvarCCallCodes As CCallCodes


'##ModelId=39EBC87E0161
Public Property Get CCallCodes() As CCallCodes
    If mvarCCallCodes Is Nothing Then
        Set mvarCCallCodes = New CCallCodes
    End If


    Set CCallCodes = mvarCCallCodes
End Property


'##ModelId=39EBC87E008E
Public Property Set CCallCodes(vData As CCallCodes)
    Set mvarCCallCodes = vData
End Property
'##ModelId=39EBC87E01D9
Private Sub Class_Terminate()
    Set mvarCCallCodes = Nothing
End Sub



'##ModelId=39EBC87E021F
Public Property Let sName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sName = 5
    mvarsName = vData
End Property


'##ModelId=39EBC87E0305
Public Property Get sName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sName
    sName = mvarsName
End Property



'##ModelId=39EBC87E0391
Public Property Let ID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ID = 5
    mvarID = vData
End Property


'##ModelId=39EBC87F00A4
Public Property Get ID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ID
    ID = mvarID
End Property



