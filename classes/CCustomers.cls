VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"CCustomer"
Attribute VB_Ext_KEY = "Member0" ,"CCustomer"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EBC86D0238"
'local variable to hold collection
'##ModelId=39EBC86D030B
Private mCol As Collection
'##ModelId=39EBC86D036F
Public Function Add(sName As String, ID As Integer, Optional sAddress As String, _
Optional sCity As String, Optional sState As String, Optional sCountry As String, _
Optional sZIP As String, Optional sPhone As String, Optional sFAX As String, _
Optional CCustomers As CCustomers, Optional sKey As String, Optional sType As String) As CCustomer
    'create a new object
    Dim objNewMember As CCustomer
    Set objNewMember = New CCustomer

    'set the properties passed into the method
    objNewMember.sName = sName
    objNewMember.sAddress = sAddress
    objNewMember.ID = ID
    objNewMember.sCity = sCity
    objNewMember.sState = sState
    objNewMember.sCountry = sCountry
    objNewMember.sZIP = sZIP
    objNewMember.sPhone = sPhone
    objNewMember.sFAX = sFAX
    objNewMember.sType = sType
    Set objNewMember.CCustomers = CCustomers
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function
'##ModelId=39EBC86F00A1
Public Property Get Item(vntIndexKey As Variant) As CCustomer
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property
'##ModelId=39EBC86F0187
Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property
'##ModelId=39EBC86F0209
Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub
'##ModelId=39EBC86F02C8
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property
'##ModelId=39EBC86F0354
Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub
'##ModelId=39EBC86F03AE
Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

