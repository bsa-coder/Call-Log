VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fQuery 
   Caption         =   "Product Query"
   ClientHeight    =   7920
   ClientLeft      =   1455
   ClientTop       =   735
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   12000
   Begin VB.Frame frmStartDate 
      Height          =   495
      Left            =   120
      TabIndex        =   25
      Top             =   2640
      Width           =   10215
      Begin VB.OptionButton optEntryDate 
         Caption         =   "One Month"
         Height          =   255
         Index           =   30
         Left            =   9000
         TabIndex        =   32
         Top             =   185
         Width           =   1095
      End
      Begin VB.OptionButton optEntryDate 
         Caption         =   "One Week"
         Height          =   255
         Index           =   7
         Left            =   7800
         TabIndex        =   31
         Top             =   185
         Width           =   1095
      End
      Begin VB.OptionButton optEntryDate 
         Caption         =   "One Day"
         Height          =   255
         Index           =   1
         Left            =   6720
         TabIndex        =   30
         Top             =   185
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.ComboBox cboStartDateYear 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         TabIndex        =   29
         Text            =   "2002"
         Top             =   130
         Width           =   855
      End
      Begin VB.ComboBox cboStartDateMonth 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   28
         Text            =   "December"
         Top             =   130
         Width           =   1335
      End
      Begin VB.ComboBox cboStartDateDay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   27
         Text            =   "30"
         Top             =   130
         Width           =   615
      End
      Begin VB.ComboBox cboStartDateLogic 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Text            =   "Logic"
         Top             =   130
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Record Entries"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5040
         TabIndex        =   33
         Top             =   190
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdDataType 
      Caption         =   "Employees"
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
      Index           =   8
      Left            =   10440
      TabIndex        =   19
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdDataType 
      Caption         =   "Links"
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
      Index           =   12
      Left            =   10440
      TabIndex        =   18
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdDataType 
      Caption         =   "Company Detail"
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
      Index           =   9
      Left            =   10440
      TabIndex        =   17
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdDataType 
      Caption         =   "Call Codes"
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
      Index           =   3
      Left            =   10440
      TabIndex        =   16
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdDataType 
      Caption         =   "Products"
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
      Index           =   2
      Left            =   10440
      TabIndex        =   15
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdDataType 
      Caption         =   "New Contacts"
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
      Index           =   1
      Left            =   10440
      TabIndex        =   14
      Top             =   2760
      Width           =   1455
   End
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
      Height          =   495
      Left            =   8640
      TabIndex        =   10
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdDataType 
      Caption         =   "New Companies"
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
      Index           =   0
      Left            =   10440
      TabIndex        =   9
      Top             =   2040
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
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Frame frmCustomer 
      Caption         =   "Query Results"
      Height          =   2895
      Left            =   120
      TabIndex        =   7
      Top             =   4920
      Width           =   10215
      Begin MSFlexGridLib.MSFlexGrid fgdDetail 
         DragIcon        =   "fQuery.frx":0000
         Height          =   2535
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   4471
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         ForeColorSel    =   -2147483635
         WordWrap        =   -1  'True
         FocusRect       =   2
         HighLight       =   2
         GridLinesFixed  =   1
         MergeCells      =   2
         AllowUserResizing=   1
         FormatString    =   ""
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
      Height          =   615
      Index           =   1
      Left            =   10440
      TabIndex        =   6
      Top             =   7200
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
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.ListBox lstEmployees 
         Height          =   1620
         Left            =   4320
         MultiSelect     =   1  'Simple
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton cmdDataType 
         Caption         =   "Call Detail"
         Height          =   375
         Index           =   17
         Left            =   4320
         TabIndex        =   20
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CommandButton cmdDataType 
         Caption         =   "Companies"
         Height          =   495
         Index           =   18
         Left            =   8520
         TabIndex        =   21
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton cmdDataType 
         Caption         =   "Companies"
         Height          =   495
         Index           =   16
         Left            =   8520
         TabIndex        =   13
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton cmdDataType 
         Caption         =   "Contacts"
         Height          =   375
         Index           =   14
         Left            =   2160
         TabIndex        =   12
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CommandButton cmdDataType 
         Caption         =   "Companies"
         Height          =   375
         Index           =   13
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ListBox lstCallCodes 
         Height          =   1620
         Left            =   2160
         MultiSelect     =   1  'Simple
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstProducts 
         Height          =   1620
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblEmployees 
         Caption         =   "Employees"
         Height          =   255
         Left            =   4320
         TabIndex        =   24
         Top             =   240
         Width           =   1815
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
Option Explicit

Dim rsGeneric As adodb.Recordset
Dim rsLinks As adodb.Recordset
Dim mlstProducts As ListBox
Dim mlstCallCodes As ListBox
Dim dDateFilter As Date

Const clCOMPANY As Integer = 0
Const clCONTACT As Integer = 1
Const clPRODUCT As Integer = 2
Const clCALLCODE As Integer = 3
Const clNEWCALLS As Integer = 17
Const clNEWCOMPANIES As Integer = 18
Const clNEWCONTACTS As Integer = 19

'##ModelId=3A0F61DC0025
Dim WithEvents BLServer As CallTrackerBLServer.BLServer
Attribute BLServer.VB_VarHelpID = -1

Sub GetProductHistory(ProductID As Long, tb As TextBox, Optional lCallCode As Long)
Dim rsCalls As adodb.Recordset
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
Property Set Products(rs As adodb.Recordset)
With rs
    .MoveLast
    .MoveFirst
    While Not .EOF
        lstProducts.AddItem .Fields("sName").Value
        lstProducts.ItemData(lstProducts.NewIndex) = .Fields("Id").Value
        .MoveNext
    Wend
End With
End Property
Property Set CallCodes(rs As adodb.Recordset) 'lb As ListBox)
With rs
    .MoveLast
    .MoveFirst
    While Not .EOF
        lstCallCodes.AddItem .Fields("sName").Value
        lstCallCodes.ItemData(lstCallCodes.NewIndex) = .Fields("Id").Value
        .MoveNext
    Wend
End With
End Property
Property Set Employees(rs As adodb.Recordset) 'lb As ListBox)
With rs
    .MoveLast
    .MoveFirst
    While Not .EOF
        lstEmployees.AddItem .Fields("sName").Value
        lstEmployees.ItemData(lstEmployees.NewIndex) = .Fields("Id").Value
        .MoveNext
    Wend
End With
End Property

Private Sub cmdClearCode_Click()
lstCallCodes.ListIndex = -1
lstProducts.ListIndex = -1
End Sub

Private Sub cmdDataType_Click(Index As Integer)
Dim iCount As Integer

    Select Case Index
        Case 0 'New Companies
            On Error Resume Next
            For iCount = optEntryDate.LBound To optEntryDate.UBound
                If optEntryDate(iCount).Value = True Then Exit For
            Next iCount
            
            On Error GoTo 0
            If Not GetData(Index, CStr(iCount * 100)) Then
                MsgBox "No new companies in date range.", vbInformation, "Query Engine"
                fgdDetail.Clear
            End If
        Case clCOMPANY, clCONTACT
            If Not GetData(Index, CStr(dDateFilter)) Then MsgBox "No new companies in date range.", vbInformation, "Query Engine"
        Case Else
            If Not GetData(Index) Then MsgBox "Unable to get data.", vbInformation, "Query Engine"
    End Select
End Sub

Private Sub cmdSearch_Click()
'If lstProducts.ListIndex = -1 Then
'    MsgBox "Please select a product.", vbExclamation, "CL-FQuery::cmdSearch"
'Else
'    If lstCallCodes.ListIndex = -1 Then
'        GetProductHistory lstProducts.ItemData(lstProducts.ListIndex), txtCallHistory
'    Else
'        GetProductHistory lstProducts.ItemData(lstProducts.ListIndex), txtCallHistory, lstCallCodes.ItemData(lstCallCodes.ListIndex)
'    End If
'End If

Load Form1
Form1.Show

End Sub

Private Sub Command1_Click()
MsgBox "FMain.LoadAllCustomers clCOMPANY"
Form_Load
End Sub

Private Sub Form_Load()
Dim iCount As Integer
Dim iCounter As Integer
Dim sGridString As String
Dim rs As adodb.Recordset


Set BLServer = CreateObject("CallTrackerBLServer.BLServer")

If BLServer Is Nothing Then
    MsgBox "Server not created"
    Unload Me
End If

'Fill Product Listbox
If Not GetData(clPRODUCT) Then MsgBox "Unable to load products."
Set Me.Products = rsGeneric

'Fill Callcode Listbox
If Not GetData(clCALLCODE) Then MsgBox "Unable to load call codes."
Set Me.CallCodes = rsGeneric

'Fill Employee Listbox
If Not GetData(8) Then MsgBox "Unable to load employees."
Set Me.Employees = rsGeneric

'Fill Date selection combo boxes
With Me
    FillDateCombo .cboStartDateLogic, .cboStartDateDay, .cboStartDateMonth, .cboStartDateYear
End With

With fgdDetail
    .MergeCells = flexMergeRestrictColumns
    
    For iCount = 0 To .Cols - 1
        .MergeCol(iCount) = True     ' Allow merge on Columns 0 thru 3
        .ColAlignment(iCount) = 1
    Next iCount
    
End With

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
Dim rs As adodb.Recordset
Dim tb As TextBox

Set tb = txtCallHistory

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
Function GetData(iData As Integer, Optional Filter As String) As Boolean
On Error GoTo ErrorHandler

'Create BL Server if none exists
If BLServer Is Nothing Then
    Set BLServer = CreateObject("CallTrackerBLServer.BLServer")
End If

'Get the requested data
If Not (BLServer Is Nothing) Then
'    Set rsGeneric = BLServer.GetLbData(iData)
'MsgBox Filter
    If Not (Filter = "") Then
        Set rsGeneric = BLServer.GetLbData(iData, CLng(Filter))
    Else
        Set rsGeneric = BLServer.GetLbData(iData)
    End If
    If rsGeneric.Fields.Count = 0 Then Exit Function
End If

'Load search detail
GetData = ShowDetail(rsGeneric)
Exit Function

ErrorHandler:
    Set rsGeneric = BLServer.GetLbData(iData)
    Resume Next
End Function
Function ShowDetail(rs As adodb.Recordset) As Boolean
Dim sGridString As String
Dim iCount As Integer
Dim iCounter As Integer
Dim iNumberOfFields As Integer

On Error GoTo ErrorHandler

ShowDetail = False
sGridString = ""
iNumberOfFields = rs.Fields.Count - 1
    
With fgdDetail
    .Redraw = False
    .Clear
    .Cols = rs.Fields.Count
    .Rows = 0
    .AddItem sGridString
    
'Load search detail
    For iCount = 1 To rs.RecordCount
        For iCounter = 0 To iNumberOfFields
            .ColAlignment(iCounter) = 1
            sGridString = sGridString & rs.Fields(iCounter).Value & vbTab
            .MergeCol(iCounter) = True     ' Allow merge on Columns 0 thru 3
        Next iCounter
        .AddItem sGridString
        rs.MoveNext
        sGridString = ""
    Next iCount

    .FixedRows = 1
'Load heading row
    For iCounter = 0 To iNumberOfFields
        .TextMatrix(0, iCounter) = rs.Fields(iCounter).Name
    Next iCounter
    
    .Redraw = True
End With

ShowDetail = True
Exit Function

ErrorHandler:
End Function
Sub DoSort()
    
    fgdDetail.Col = 0
    fgdDetail.ColSel = fgdDetail.Cols - 1
    fgdDetail.Sort = 1 ' Generic Ascending
    
End Sub
Private Sub fgdDetail_DragDrop(Source As VB.Control, x As Single, y As Single)
    If fgdDetail.Tag = "" Then Exit Sub
    fgdDetail.Redraw = False
    fgdDetail.ColPosition(Val(fgdDetail.Tag)) = fgdDetail.MouseCol
    DoSort
    fgdDetail.Redraw = True
End Sub

Private Sub fgdDetail_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    fgdDetail.Tag = ""
    If fgdDetail.MouseRow <> 0 Then Exit Sub
    fgdDetail.Tag = Str(fgdDetail.MouseCol)
    fgdDetail.Drag 1
End Sub

Private Function FillList(lb As ListBox, rs As adodb.Recordset) As Boolean
    While Not rs.EOF
        lb.AddItem rs.Fields("sName"), 0
        lb.ItemData(lb.ListIndex) = rs.Fields("Id").Value
        rs.MoveNext
    Wend
End Function

Private Sub optEntryDate_Click(Index As Integer)

    Select Case Index
        Case 1 'yesterday
        Case 7 'last week
        Case 30 'last month
        Case Else
    End Select
    dDateFilter = Now() - Index
'    MsgBox dDateFilter
    
End Sub
Private Sub FillDateCombo(ByRef cboLogic As ComboBox, ByRef cboDay As ComboBox, ByRef cboMonth As ComboBox, ByRef cboYear As ComboBox)
Dim iCount As Integer

'Fill Logic combo
cboLogic.AddItem "Before"
cboLogic.ItemData(cboLogic.NewIndex) = 1
cboLogic.AddItem "Equal"
cboLogic.ItemData(cboLogic.NewIndex) = 2
cboLogic.AddItem "After"
cboLogic.ItemData(cboLogic.NewIndex) = 3

'Fill Day combo
For iCount = 1 To 31
    cboDay.AddItem iCount
    cboDay.ItemData(cboDay.NewIndex) = iCount
Next iCount

'Fill Month combo
cboMonth.AddItem "January"
cboMonth.ItemData(cboMonth.NewIndex) = 1
cboMonth.AddItem "February"
cboMonth.ItemData(cboMonth.NewIndex) = 2
cboMonth.AddItem "March"
cboMonth.ItemData(cboMonth.NewIndex) = 3
cboMonth.AddItem "April"
cboMonth.ItemData(cboMonth.NewIndex) = 4
cboMonth.AddItem "May"
cboMonth.ItemData(cboMonth.NewIndex) = 5
cboMonth.AddItem "June"
cboMonth.ItemData(cboMonth.NewIndex) = 6
cboMonth.AddItem "July"
cboMonth.ItemData(cboMonth.NewIndex) = 7
cboMonth.AddItem "August"
cboMonth.ItemData(cboMonth.NewIndex) = 8
cboMonth.AddItem "September"
cboMonth.ItemData(cboMonth.NewIndex) = 9
cboMonth.AddItem "October"
cboMonth.ItemData(cboMonth.NewIndex) = 10
cboMonth.AddItem "November"
cboMonth.ItemData(cboMonth.NewIndex) = 11
cboMonth.AddItem "December"
cboMonth.ItemData(cboMonth.NewIndex) = 12

'Fill Year combo
cboYear.AddItem "1999"
cboYear.ItemData(cboYear.NewIndex) = 1999
cboYear.AddItem "2000"
cboYear.ItemData(cboYear.NewIndex) = 2000
cboYear.AddItem "2001"
cboYear.ItemData(cboYear.NewIndex) = 2001
cboYear.AddItem "2002"
cboYear.ItemData(cboYear.NewIndex) = 2002
cboYear.AddItem "2003"
cboYear.ItemData(cboYear.NewIndex) = 2003
cboYear.AddItem "2004"
cboYear.ItemData(cboYear.NewIndex) = 2004

End Sub
