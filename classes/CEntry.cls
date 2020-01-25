VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B4FC392013C"
Option Explicit

'Dim ETime As Integer
'Dim EDate As Date
'Dim EmplID As String
'Dim EIndex As Integer

'##ModelId=3B4FC3920196
Dim DEDate As Date 'Date of call
'##ModelId=3B4FC39201DC
Dim sEmplID As String 'Employee ID
'##ModelId=3B4FC3920222
Dim iEIndex As Integer 'index of call, temporary
'##ModelId=3B4FC3920272
Dim iETime As Integer 'Length of call in minutes
'##ModelId=3B4FC39202C2
Dim sCaller As String 'type of caller 3 char code
'##ModelId=3B4FC392031D
Dim sProduct As String 'call subject
'Dim dDeDate As Date 'Data Entry Date

'##ModelId=3B4FC3930057
Property Get Caller() As String
    Caller = sCaller
End Property
'##ModelId=3B4FC3920381
Property Let Caller(Caller As String)
    sCaller = Caller
End Property


'##ModelId=3B4FC3930197
Property Get EDate() As Date
    EDate = DEDate
End Property
'##ModelId=3B4FC3930210
Property Get DateEntered() As Date
    DateEntered = Now()
End Property

'##ModelId=3B4FC39300BB
Property Let EDate(EDate As Date)
    DEDate = EDate
End Property
'##ModelId=3B4FC393036E
Property Get EIndex() As Integer

EIndex = iEIndex

End Property
'##ModelId=3B4FC3930288
Property Let EIndex(EIndex As Integer)

iEIndex = EIndex

End Property

'##ModelId=3B4FC3940012
Property Let EmplID(EmplID As String)
    sEmplID = EmplID
End Property

'##ModelId=3B4FC394010D
Property Get EmplID() As String
    EmplID = sEmplID
End Property


'##ModelId=3B4FC39402A7
Property Get ETime() As Integer
    ETime = iETime
End Property
'##ModelId=3B4FC39401A3
Property Let ETime(ETime As Integer)
    iETime = ETime
End Property



'##ModelId=3B4FC395006E
Property Get Product() As String
    Product = sProduct
End Property
'##ModelId=3B4FC394033D
Property Let Product(Product As String)
    sProduct = Product
End Property



