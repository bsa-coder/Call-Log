VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCallCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"CCallCode"
Attribute VB_Ext_KEY = "Member0" ,"CCallCode"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EBC87F019E"
'local variable to hold collection
'##ModelId=39EBC87F028E
Private mCol As Collection

'##ModelId=39EBC87F02FD
Public Function Add(sName As String, ID As Integer, Optional CCallCodes As CCallCodes, Optional sKey As String) As CCallCode
    'create a new object
    Dim objNewMember As CCallCode
    Set objNewMember = New CCallCode


    'set the properties passed into the method
    objNewMember.ID = ID
    objNewMember.sName = sName
    Set objNewMember.CCallCodes = CCallCodes
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

'##ModelId=39EBC8800127
Public Property Get Item(vntIndexKey As Variant) As CCallCode
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



'##ModelId=39EBC8800218
Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


'##ModelId=39EBC88002A4
Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


'##ModelId=39EBC880036C
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


'##ModelId=39EBC8810010
Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


'##ModelId=39EBC881006B
Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

