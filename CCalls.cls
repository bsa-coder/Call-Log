VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCalls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"CCall"
Attribute VB_Ext_KEY = "Member0" ,"CCall"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EBC88F0101"
'local variable to hold collection
'##ModelId=39EBC88F01DD
Private mCol As Collection

'##ModelId=39EBC88F0237
Public Function Add(sLastName As String, sCompanyName As String, sContactName As String, sProductName As String, sCallType As String, dNoteDate As String, sNote As String, iCallTime As Integer, Optional CCalls As CCalls, Optional sKey As String) As CCall
    'create a new object
    Dim objNewMember As CCall
    Set objNewMember = New CCall

    'set the properties passed into the method
    objNewMember.sLastName = sLastName
    objNewMember.sCompanyName = sCompanyName
    objNewMember.sContactName = sContactName
    objNewMember.sProductName = sProductName
    objNewMember.sCallType = sCallType
    objNewMember.dNoteDate = dNoteDate
    objNewMember.sNote = sNote
    objNewMember.iCallTime = iCallTime
    Set objNewMember.CCalls = CCalls
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

'##ModelId=39EBC8900211
Public Property Get Item(vntIndexKey As Variant) As CCall
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



'##ModelId=39EBC89002CF
Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


'##ModelId=39EBC8900347
Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


'##ModelId=39EBC89003E7
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


'##ModelId=39EBC8910078
Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


'##ModelId=39EBC89100BE
Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

