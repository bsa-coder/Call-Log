VERSION 5.00
Begin VB.Form fQuery 
   Caption         =   "Product Query"
   ClientHeight    =   7920
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12000
   Icon            =   "fTreeviewTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Load Companies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      TabIndex        =   18
      Top             =   4800
      Width           =   1455
   End
   Begin VB.OptionButton optEntryDate 
      Caption         =   "One Month"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   8640
      TabIndex        =   16
      Top             =   1560
      Width           =   1455
   End
   Begin VB.OptionButton optEntryDate 
      Caption         =   "One Week"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   8640
      TabIndex        =   15
      Top             =   1320
      Width           =   1455
   End
   Begin VB.OptionButton optEntryDate 
      Caption         =   "One Day"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   8640
      TabIndex        =   14
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdNewContacts 
      Caption         =   "New Contacts"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      TabIndex        =   12
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdNewCompanies 
      Caption         =   "New Companies"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      TabIndex        =   11
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdNewCalls 
      Caption         =   "New Calls"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      TabIndex        =   10
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdClearCode 
      Caption         =   "Clear Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10440
      TabIndex        =   9
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Frame frmCustomer 
      Caption         =   "Query Results"
      Height          =   5535
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   10215
      Begin VB.TextBox txtCallHistory 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   240
         Width           =   9975
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   10440
      TabIndex        =   6
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H80000012&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10440
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame frmCall 
      Caption         =   "Search Criteria"
      Height          =   2235
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.OptionButton optEntryDate 
         Caption         =   "All"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   8520
         TabIndex        =   13
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ListBox lstCallCodes 
         Height          =   1620
         Left            =   2160
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstProducts 
         Height          =   1620
         ItemData        =   "fTreeviewTest.frx":0442
         Left            =   120
         List            =   "fTreeviewTest.frx":0444
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Record Entries"
         Enabled         =   0   'False
         Height          =   255
         Left            =   8520
         TabIndex        =   17
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblCallCode 
         Caption         =   "Call Code"
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblProduct 
         Caption         =   "Product"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "fQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsGeneric As ADODB.Recordset
Dim rsLinks As ADODB.Recordset
Dim mlstProducts As ListBox
Dim mlstCallCodes As ListBox

Const clCOMPANY As Integer = 0
Const clNEWCALLS As Integer = 17
Const clNEWCOMPANIES As Integer = 18
Const clNEWCONTACTS As Integer = 19

'##ModelId=3A0F61DC0025
Dim WithEvents BLServer As CallTrackerBLServer.BLServer
Attribute BLServer.VB_VarHelpID = -1

Sub GetProductHistory(ProductID As Long, tb As TextBox, Optional lCallCode As Long)
Dim rsCalls As ADODB.Recordset
Dim sText As String
Dim CallTime As Integer

'Create BL Server if none exists
If BLServer Is Nothing Then
    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
End If

'Get all the calls for the company
If Not (rsCalls Is Nothing) Then
    If rsCalls!ID <> ProductID Then
        Set rsCalls = BLServer.GetProductHistory(ProductID)
    End If
Else
    Set rsCalls = BLServer.GetProductHistory(ProductID)
End If

tb.Text = ""

'If there were calls, fill collection with calls
If Not (rsCalls Is Nothing) Then
    Do Until rsCalls.EOF
        If lCallCode = 0 Then
            sText = rsCalls!NoteDate & " (" & rsCalls!iCallTime & " min) " & rsCalls!ContactName & vbCrLf
            sText = sText & "-----------------------------------------------" & vbCrLf
            sText = sText & rsCalls!ProductName & "  :  "
            sText = sText & rsCalls!CallType & " (" & rsCalls!LastName & ")" & " Case: " & rsCalls!CaseID & vbCrLf
            sText = sText & "..............................................." & vbCrLf
            sText = sText & rsCalls!sNote & vbCrLf
            sText = sText & "===============================================" & vbCrLf
        Else
'=========================
'This section can be combined with the previous for a better sub
'       If (lCallCode = 0) or (rsCalls!CallCodeID = lCallCode) Then
            If rsCalls!CallCodeID = lCallCode Then
                sText = rsCalls!NoteDate & " (" & rsCalls!iCallTime & " min) " & rsCalls!ContactName & vbCrLf
                sText = sText & "-----------------------------------------------" & vbCrLf
                sText = sText & rsCalls!ProductName & "  :  "
                sText = sText & rsCalls!CallType & " (" & rsCalls!LastName & ")" & " Case: " & rsCalls!CaseID & vbCrLf
                sText = sText & "..............................................." & vbCrLf
                sText = sText & rsCalls!sNote & vbCrLf
                sText = sText & "===============================================" & vbCrLf
            Else
                sText = ""
            End If
        End If
        tb.Text = tb.Text & sText
        
        rsCalls.MoveNext
    Loop
Else
    tb.Text = ""
End If

End Sub
Private Sub cmdCancel_Click(Index As Integer)
Unload Me
End Sub
Private Sub cmdClose_Click()
Unload Me
End Sub
Property Set Products(lb As ListBox)
'Set mlstProducts = lb
Dim iCount As Integer
For iCount = 0 To lb.ListCount - 1
        lb.ListIndex = iCount
        lstProducts.AddItem lb.Text
        lstProducts.ItemData(lstProducts.NewIndex) = lb.ItemData(lb.ListIndex)
Next iCount
End Property
Property Set CallCodes(lb As ListBox)
Dim iCount As Integer
For iCount = 0 To lb.ListCount - 1
        lb.ListIndex = iCount
        lstCallCodes.AddItem lb.Text
        lstCallCodes.ItemData(lstCallCodes.NewIndex) = lb.ItemData(lb.ListIndex)
Next iCount
End Property

Private Sub cmdClearCode_Click()
lstCallCodes.ListIndex = -1
lstProducts.ListIndex = -1
End Sub

Private Sub cmdNewCalls_Click()
Dim rsCalls As ADODB.Recordset
Dim sText As String
Dim CallTime As Integer

'Create BL Server if none exists
If BLServer Is Nothing Then
    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
End If

'Get all the calls for the company
If Not (rsCalls Is Nothing) Then
    Set rsCalls = BLServer.GetLbData(clNEWCALLS)
End If

DisplayCalls

End Sub

Private Sub cmdSearch_Click()
If lstProducts.ListIndex = -1 Then
    MsgBox "Please select a product.", vbExclamation, "CL-FQuery::cmdSearch"
Else
    If lstCallCodes.ListIndex = -1 Then
        GetProductHistory lstProducts.ItemData(lstProducts.ListIndex), txtCallHistory
    Else
        GetProductHistory lstProducts.ItemData(lstProducts.ListIndex), txtCallHistory, lstCallCodes.ItemData(lstCallCodes.ListIndex)
    End If
End If
End Sub

Private Sub Command1_Click()
FMain.LoadAllCustomers (clCOMPANY)
End Sub

Private Sub Form_Load()
Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
If BLServer Is Nothing Then
    MsgBox "Server not created"
    Unload Me
End If
End Sub

Private Sub Form_Terminate()
If Not BLServer Is Nothing Then Set BLServer = Nothing
If Not fQuery Is Nothing Then Set fQuery = Nothing
End Sub

Private Sub DisplayCalls()
Dim sQuery As String
Dim sSelect As String
Dim sFrom As String
Dim sWhere As String
Dim sOrderBy As String
Dim sFilter As String
Dim iDateOffset As Integer
Dim rs As ADODB.Recordset

'Get the date range
For iCount = optEntryDate.LBound To optEntryDate.UBound
    If optEntryDate(iCount).Value = True Then Exit For
Next

Select Case iCount
    Case 0 'day
        iDateOffset = 1
    Case 1 'week
        iDateOffset = 7
    Case 2 'month
        iDateOffset = 30
    Case 3 'all
        iDateOffset = CInt(Now())
End Select
    
'create the select statement for existing recordset
sSelect = "SELECT * "
sSelect = sSelect & " "

sFrom = "FROM rscalls "

sWhere = "WHERE (rscalls!NoteDate = " & Now() - 7 & ") "

sOrderBy = "ORDER BY "
sOrderBy = sOrderBy & "[rscalls]![DateEntered], [rscalls]![Address]; "

sQuery = sSelect & sFrom & sWhere & sOrderBy

rs.ActiveConnection = rsCalls
rs.Open sQuery, Nothing, adOpenForwardOnly, adLockReadOnly

tb.Text = ""

'If there were calls, show them
If Not (rsCalls Is Nothing) Then
    Do Until rsCalls.EOF
        If lCallCode = 0 Then
            sText = rsCalls!NoteDate & " (" & rsCalls!iCallTime & " min) " & rsCalls!ContactName & vbCrLf
            sText = sText & "-----------------------------------------------" & vbCrLf
            sText = sText & rsCalls!ProductName & "  :  "
            sText = sText & rsCalls!CallType & " (" & rsCalls!LastName & ")" & " Case: " & rsCalls!CaseID & vbCrLf
            sText = sText & "..............................................." & vbCrLf
            sText = sText & rsCalls!sNote & vbCrLf
            sText = sText & "===============================================" & vbCrLf
        Else
'=========================
'This section can be combined with the previous for a better sub
'       If (lCallCode = 0) or (rsCalls!CallCodeID = lCallCode) Then
            If rsCalls!CallCodeID = lCallCode Then
                sText = rsCalls!NoteDate & " (" & rsCalls!iCallTime & " min) " & rsCalls!ContactName & vbCrLf
                sText = sText & "-----------------------------------------------" & vbCrLf
                sText = sText & rsCalls!ProductName & "  :  "
                sText = sText & rsCalls!CallType & " (" & rsCalls!LastName & ")" & " Case: " & rsCalls!CaseID & vbCrLf
                sText = sText & "..............................................." & vbCrLf
                sText = sText & rsCalls!sNote & vbCrLf
                sText = sText & "===============================================" & vbCrLf
            Else
                sText = ""
            End If
        End If
        tb.Text = tb.Text & sText
        
        rsCalls.MoveNext
    Loop
Else
    tb.Text = ""
End If

End Sub
