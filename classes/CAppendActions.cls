VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAppendActions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EE7A6800C5"
'Collection class CAppendActions (generated)


'local variable to hold collection
'##ModelId=39EE7A6801DD
Private mCol As Collection

'used when referencing an element in the collection
'vntIndexKey contains either the Index or Key, which is why
'it is declared as a Variant
'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
'##ModelId=39EE7A6A0168
Public Property Get Item(vntIndexKey As Variant) As CAppendAction
Attribute Item.VB_UserMemId = 0
   Set Item = mCol(vntIndexKey)
End Property

'used when removing an element from the collection
'vntIndexKey contains either the Index or Key, which is why
'it is declared as a Variant
'Syntax: x.Remove
'##ModelId=39EE7A6A0059
Public Sub Remove(vntIndexKey As Variant)
   mCol.Remove vntIndexKey
End Sub

'used when adding a new item to the collection.
'Syntax: x.Add xyz
'##ModelId=39EE7A690120
Public Sub Add(Item As CAppendAction, Optional Key As String, Optional Before As Variant, Optional After As Variant)
    If IsMissing(Key) Then
        mCol.Add Item
    Else
        mCol.Add Item, Key
    End If

End Sub

'used when retrieving the number of elements in the collection.
'Syntax: Debug.Print x.Count
'##ModelId=39EE7A69006C
Public Property Get Count() As Long
   Count = mCol.Count
End Property

'this property allows you to enumerate
'this collection with the For...Each syntax
'##ModelId=39EE7A680364
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
   Set NewEnum = mCol.[_NewEnum]
End Property

'creates the collection when this class is created
'##ModelId=39EE7A680309
Private Sub Class_Initialize()
   Set mCol = New Collection
End Sub

'destroys collection when this class is terminated
'##ModelId=39EE7A6802AF
Private Sub Class_Terminate()
   Set mCol = Nothing
End Sub
