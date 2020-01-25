VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CProducts"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EBC8870131"
'##ModelId=39EBC88701F0
Private mvarID As Integer 'local copy
'##ModelId=39EBC8870254
Private mvarsName As String 'local copy
'##ModelId=39EBC88702C2
Private mvarCProducts As CProducts


'##ModelId=39EBC8880010
Public Property Get CProducts() As CProducts
    If mvarCProducts Is Nothing Then
        Set mvarCProducts = New CProducts
    End If


    Set CProducts = mvarCProducts
End Property


'##ModelId=39EBC8870308
Public Property Set CProducts(vData As CProducts)
    Set mvarCProducts = vData
End Property
'##ModelId=39EBC8880093
Private Sub Class_Terminate()
    Set mvarCProducts = Nothing
End Sub



'##ModelId=39EBC88800ED
Public Property Let sName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sName = 5
    mvarsName = vData
End Property


'##ModelId=39EBC88801E7
Public Property Get sName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sName
    sName = mvarsName
End Property



'##ModelId=39EBC888027D
Public Property Let ID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ID = 5
    mvarID = vData
End Property


'##ModelId=39EBC8880396
Public Property Get ID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ID
    ID = mvarID
End Property



