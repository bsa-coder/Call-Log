VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAppendAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"39EE7220015C"
'Append class for all database append requests to request.
'Methods within the class perform appends.
Option Base 0
Option Explicit

'set this to 0 to disable debug code in this class
#Const DebugMode = 1
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=39EE7A670019
    Private mlClassDebugID As Long
#End If

'Integer to show server status for diagnostics
'##ModelId=39EE73B60138
Private mStatus As Integer ' = 0

'Code to choose the action type
'##ModelId=39EE7511008D
Private mActionType As Integer ' = 0

'Contains the data to be appended
'##ModelId=39EE7519034B
Private mActionObject As Object

'##ModelId=39EE73980370
Private Function AppendCallCode(CallCodeId As Long, EmplID As Long) As Integer
    On Error GoTo AppendCallCodeErr

    'your code goes here...

    Exit Function
AppendCallCodeErr:
    Call RaiseError(MyUnhandledError, "CAppendAction:AppendCallCode Method")
End Function

'##ModelId=39EE738C034A
Private Function AppendContact(FirstName As String, LastName As String, Title As String, EmplID As Long) As Integer
    On Error GoTo AppendContactErr

    'your code goes here...

    Exit Function
AppendContactErr:
    Call RaiseError(MyUnhandledError, "CAppendAction:AppendContact Method")
End Function

'##ModelId=39EE736D015B
Private Function AppendLink(CustomerID As Long, ContactId As Long, EmplID As Long) As Integer
    On Error GoTo AppendLinkErr

    'your code goes here...

    Exit Function
AppendLinkErr:
    Call RaiseError(MyUnhandledError, "CAppendAction:AppendLink Method")
End Function

'##ModelId=39EE737702A0
Private Function AppendCall() As Integer
    On Error GoTo AppendCallErr

    'your code goes here...

    Exit Function
AppendCallErr:
    Call RaiseError(MyUnhandledError, "CAppendAction:AppendCall Method")
End Function

'##ModelId=39EE73810290
Private Function AppendCustomer() As Integer
    On Error GoTo AppendCustomerErr

    'your code goes here...

    Exit Function
AppendCustomerErr:
    Call RaiseError(MyUnhandledError, "CAppendAction:AppendCustomer Method")
End Function

'##ModelId=39EE74A8015E
Private Function SelectAction() As Integer
    On Error GoTo SelectActionErr

    'your code goes here...

    Exit Function
SelectActionErr:
    Call RaiseError(MyUnhandledError, "CAppendAction:SelectAction Method")
End Function

#If DebugMode Then
    '##ModelId=39EE7A6700EB
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

'##ModelId=39EE7A660252
Public Property Get ActionObject() As Object
    Set ActionObject = mActionObject
End Property

'##ModelId=39EE7A6600E0
Public Property Set ActionObject(ByVal Value As Object)
    Set mActionObject = Value
End Property

'##ModelId=39EE7A6503E1
Public Property Get ActionType() As Integer
   Let ActionType = mActionType
End Property

'##ModelId=39EE7A650283
Public Property Let ActionType(ByVal Value As Integer)
    Let mActionType = Value
End Property

'##ModelId=39EE7A6501B1
Public Property Get Status() As Integer
   Let Status = mStatus
End Property


'##ModelId=39EE7A65005C
Public Property Let Status(ByVal Value As Integer)
    Let mStatus = Value
End Property
