VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CCalls"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EBC88900EE"
'local variable(s) to hold property value(s)
'##ModelId=39EBC889018E
Private mvarsCompanyName As String 'local copy
'##ModelId=39EBC88901D5
Private mvarsProduct As String 'local copy
'##ModelId=39EBC8890225
Private mvarsCallCode As String 'local copy
'##ModelId=39EBC889027F
Private mvarsCallNote As String 'local copy
'##ModelId=39EBC88902D9
Private mvardDate As Date 'local copy
'##ModelId=39EBC889033D
Private mvarsContactName As String 'local copy
'##ModelId=39EBC88903AB
Private mvardNoteDate As String 'local copy
'##ModelId=39EBC88A0031
Private mvarsNote As String 'local copy
'##ModelId=39EBC88A00A0
Private mvariCallTime As Integer 'local copy
'##ModelId=39EBC88A0122
Private mvarCCalls As CCalls
'##ModelId=39EBC88A012C
Private mvarsLastName As String 'local copy
'##ModelId=39EBC88A01AE
Public Property Let sLastName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sLastName = 5
    mvarsLastName = vData
End Property


'##ModelId=39EBC88A0294
Public Property Get sLastName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sLastName
    sLastName = mvarsLastName
End Property
'##ModelId=39EBC88B003D
Public Property Get CCalls() As CCalls
    If mvarCCalls Is Nothing Then
        Set mvarCCalls = New CCalls
    End If


    Set CCalls = mvarCCalls
End Property


'##ModelId=39EBC88A0320
Public Property Set CCalls(vData As CCalls)
    Set mvarCCalls = vData
End Property
'##ModelId=39EBC88B00C9
Private Sub Class_Terminate()
    Set mvarCCalls = Nothing
End Sub
'##ModelId=39EBC88B0137
Public Property Let sNote(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNote = 5
    mvarsNote = vData
End Property
'##ModelId=39EBC88B0246
Public Property Get sNote() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNote
    sNote = mvarsNote
End Property
'##ModelId=39EBC88B02E6
Public Property Let iCallTime(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNote = 5
    mvariCallTime = vData
End Property
'##ModelId=39EBC88C002A
Public Property Get iCallTime() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNote
    iCallTime = mvariCallTime
End Property
'##ModelId=39EBC88C00D5
Public Property Let dNoteDate(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dNoteDate = 5
    mvardNoteDate = vData
End Property
'##ModelId=39EBC88C020B
Public Property Get dNoteDate() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dNoteDate
    dNoteDate = mvardNoteDate
End Property
'##ModelId=39EBC88C02BF
Public Property Let sContactName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sContactName = 5
    mvarsContactName = vData
End Property


'##ModelId=39EBC88D0022
Public Property Get sContactName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sContactName
    sContactName = mvarsContactName
End Property



'##ModelId=39EBC88D00E0
Public Property Let sCallType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCallCode = 5
    mvarsCallCode = vData
End Property


'##ModelId=39EBC88D0253
Public Property Get sCallType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCallCode
    sCallType = mvarsCallCode
End Property



'##ModelId=39EBC88D031B
Public Property Let sProductName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProduct = 5
    mvarsProduct = vData
End Property


'##ModelId=39EBC88E00AF
Public Property Get sProductName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProduct
    sProductName = mvarsProduct
End Property



'##ModelId=39EBC88E0182
Public Property Let sCompanyName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCompanyName = 5
    mvarsCompanyName = vData
End Property


'##ModelId=39EBC88E0308
Public Property Get sCompanyName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCompanyName
    sCompanyName = mvarsCompanyName
End Property



